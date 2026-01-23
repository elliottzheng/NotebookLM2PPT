import numpy as np
from scipy.interpolate import griddata

def inpaint_manual(img, mask, fill_color=(255, 255, 255), max_iter=20):
    """
    手动指定颜色的去水印函数
    
    参数:
        img: 图片数据 (H, W, 3)
        mask: 遮罩 (H, W)
        fill_color: 你想填入的颜色，元组或列表，例如 (255, 255, 255)。
                    注意：如果你用 cv2.imread 读取，顺序是 (B, G, R)！
                    如果你用 PIL 或 matplotlib，顺序是 (R, G, B)！
        max_iter: 边缘平滑次数。0 表示直接填充不平滑；建议 10-20 用于消除锯齿。
    """
    # 1. 格式预处理
    img_float = img.astype(np.float32)
    
    # 确保 Mask 是二维
    if mask.ndim == 3: mask = mask[:, :, 0]
    mask = mask > 100 # 二值化
    
    # 【步骤 A】膨胀 Mask
    # 这一步不能省，否则文字边缘的灰色像素会留下来，形成一圈脏轮廓
    # mask_dilated = simple_dilate(mask, iterations=3)
    mask_dilated = mask
    
    # 2. 性能优化：提取 ROI (只处理 Mask 区域)
    rows, cols = np.where(mask_dilated)
    if len(rows) == 0: return img
    
    pad = 2
    y1, y2 = max(0, rows.min()-pad), min(img.shape[0], rows.max()+pad+1)
    x1, x2 = max(0, cols.min()-pad), min(img.shape[1], cols.max()+pad+1)
    
    img_crop = img_float[y1:y2, x1:x2]
    mask_crop = mask_dilated[y1:y2, x1:x2]
    
    # --- 【步骤 B】暴力填充指定颜色 ---
    # 将 fill_color 转换为 numpy 数组，方便广播
    color_val = np.array(fill_color, dtype=np.float32)
    
    # 直接赋值！原本的水印直接被这个颜色覆盖
    # img_crop[mask_crop] 选中的是 (N, 3) 的像素，直接赋 (3,) 的颜色值会自动广播
    img_crop[mask_crop] = color_val
    
    # --- 【步骤 C】边缘羽化 (可选) ---
    # 如果 max_iter > 0，我们让填进去的颜色和周围稍微融合一下，消除边缘的“切割感”
    if max_iter > 0:
        mask_3d = mask_crop[:, :, np.newaxis]
        for i in range(max_iter):
            # 拉普拉斯平滑
            neighbor_sum = (
                img_crop[:-2, 1:-1] + 
                img_crop[2:,  1:-1] + 
                img_crop[1:-1, :-2] + 
                img_crop[1:-1, 2:]
            )
            avg = neighbor_sum / 4.0
            
            center_view = img_crop[1:-1, 1:-1]
            roi_mask = mask_3d[1:-1, 1:-1]
            
            # 更新颜色
            np.copyto(center_view, avg, where=roi_mask)

    # 3. 放回原图
    img_float[y1:y2, x1:x2] = img_crop
    
    return np.clip(img_float, 0, 255).astype(np.uint8)


def inpaint_numpy_onion(img, mask):
    """
    纯 NumPy 实现类似 Telea 的由外向内修复算法 (Onion-Peel)
    核心原理：每次只修复 Mask 最边缘的一圈像素，参考其周围已知的像素，
    修完一圈后，这一圈变回已知像素，继续修下一圈。
    这样能把背景的线条和渐变“长”进去，而不是糊成一团。
    """
    # 1. 预处理
    # 转为 float 方便计算
    img = img.astype(np.float32)
    
    # 确保 mask 是 0/1 的二维矩阵 (0=背景, 1=水印)
    if mask.ndim == 3: mask = mask[:, :, 0]
    mask = (mask > 100).astype(np.uint8)
    
    # 提取 ROI (只处理水印周围，加速运算)
    rows, cols = np.where(mask)
    if len(rows) == 0: return img.astype(np.uint8)
    
    pad = 2
    y1, y2 = max(0, rows.min()-pad), min(img.shape[0], rows.max()+pad+1)
    x1, x2 = max(0, cols.min()-pad), min(img.shape[1], cols.max()+pad+1)
    
    img_crop = img[y1:y2, x1:x2]
    mask_crop = mask[y1:y2, x1:x2]
    
    # remaining_mask 记录还需要修的区域，一开始等于 mask_crop
    remaining_mask = mask_crop.copy()
    
    # 2. 核心循环：只要还有没修完的地方，就一直剥洋葱
    # 为了防止死循环（孤岛像素），设置最大步数
    max_steps = max(img_crop.shape) 
    
    # 定义 3x3 邻域的切片偏移量
    offsets = [
        (-1, -1), (-1, 0), (-1, 1),
        ( 0, -1),          ( 0, 1),
        ( 1, -1), ( 1, 0), ( 1, 1)
    ]
    
    for step in range(max_steps):
        # 如果全都修完了，退出
        if np.sum(remaining_mask) == 0:
            break
            
        # --- A. 寻找当前的“边缘层” ---
        # 逻辑：像素本身是需要修的(1)，但它上下左右至少有一个像素是已知的(0)
        # 这一步可以用形态学腐蚀来实现：边缘 = Mask - 腐蚀后的Mask
        
        # 简单的 numpy 实现腐蚀 (Erosion)
        # 只有当周围全是 1 时，中心才是 1；只要有一个邻居是 0，中心变 0
        # 这样最外圈的 1 就会变成 0，剩下的 core 就是内层
        
        # 利用切片快速判断邻域是否全为 1
        eroded = remaining_mask.copy()
        # 只要有一个方向是0，就腐蚀掉
        # 也就是：up & down & left & right
        up    = remaining_mask[1:]
        down  = remaining_mask[:-1]
        left  = remaining_mask[:, 1:]
        right = remaining_mask[:, :-1]
        
        # 为了形状匹配，需要 pad 回去或者切片操作
        # 这里用更简单的逻辑：
        # 构造一个 sum_neighbors，计算每个像素周围有几个 1
        # 如果周围 1 的数量 < 8 (对于3x3)，说明它挨着 0，它就是边缘
        
        # 稍微麻烦一点但纯 numpy 的方法：
        is_border = np.zeros_like(remaining_mask, dtype=bool)
        
        # 计算 Mask 区域内，每一个像素周围 "已知像素(0)" 的数量
        # 我们利用 convolution 的思想
        # known_mask: 0=待修, 1=已知 (注意这里反转一下逻辑方便求和)
        known_mask = 1 - remaining_mask
        
        # 或者是更直接的方法：
        # 边缘 = Mask AND (NOT 腐蚀Mask)
        # 我们手动实现一次腐蚀
        m_pad = np.pad(remaining_mask, 1, mode='constant', constant_values=1)
        # 检查上下左右，如果全是1，则保持1（核心），否则是边缘
        sub_up    = m_pad[:-2, 1:-1]
        sub_down  = m_pad[2:,  1:-1]
        sub_left  = m_pad[1:-1, :-2]
        sub_right = m_pad[1:-1, 2:]
        
        core = remaining_mask & sub_up & sub_down & sub_left & sub_right
        border_mask = remaining_mask & (~core)
        
        # 如果这一轮找不到边缘（比如全是孤岛），强制结束防止死循环
        if np.sum(border_mask) == 0:
            # 最后的兜底：剩下的直接全填均值
            pass_mask = remaining_mask.astype(bool)
            if np.sum(pass_mask) > 0:
                # 找全局已知像素的均值
                valid_pixels = img_crop[~remaining_mask.astype(bool)]
                if len(valid_pixels) > 0:
                    fill_val = np.mean(valid_pixels, axis=0)
                else:
                    fill_val = np.array([255, 255, 255])
                img_crop[pass_mask] = fill_val
            break
            
        # --- B. 修复边缘层 ---
        # 对于 border_mask 中的每一个像素，计算它周围 "已知像素" 的加权平均
        
        # 获取边缘像素坐标
        border_y, border_x = np.where(border_mask)
        
        # 为了加速，我们不搞复杂的加权，直接算周围已知像素的均值
        # 这一步必须向量化，否则慢死
        
        # 累加器
        sum_color = np.zeros((len(border_y), 3), dtype=np.float32)
        count_valid = np.zeros((len(border_y), 1), dtype=np.float32)
        
        for dy, dx in offsets:
            ny, nx = border_y + dy, border_x + dx
            
            # 越界检查
            valid_idx = (ny >= 0) & (ny < img_crop.shape[0]) & \
                        (nx >= 0) & (nx < img_crop.shape[1])
            
            # 只取 valid_idx 为 True 的部分进行后续判断
            # 且该邻居必须是 "已知像素" (remaining_mask[ny, nx] == 0)
            
            # 这里的 numpy 索引比较绕，我们换个思路：
            # 直接全图卷积计算 Sum(Color * Known) / Sum(Known)
            pass
        
        # === 简易向量化替代方案 ===
        # 上面的 loop 比较难写，我们用全图移位操作
        # 1. 准备 Known Mask
        known = (1 - remaining_mask).astype(np.float32)
        known_3d = known[:, :, np.newaxis]
        
        # 2. 准备 Image * Known
        img_known = img_crop * known_3d
        
        # 3. 计算邻域总和 (Color) 和 邻域计数 (Weight)
        # 手动实现 3x3 卷积 (Sum)
        total_color = np.zeros_like(img_crop)
        total_weight = np.zeros_like(known_3d)
        
        for dy, dx in offsets:
            # 移位切片
            # src: [1:-1] based logic
            # 我们用 pad 方式更简单
            img_pad = np.pad(img_known, ((1,1),(1,1),(0,0)), 'constant')
            w_pad   = np.pad(known_3d,  ((1,1),(1,1),(0,0)), 'constant')
            
            # 比如 dy=-1, dx=-1，相当于取右下方的像素填到左上
            # 对应的切片是 [0:-2, 0:-2]
            # 为了通用性，我们直接 slice
            
            sy_start, sy_end = 1+dy, img_pad.shape[0]-1+dy
            sx_start, sx_end = 1+dx, img_pad.shape[1]-1+dx
            
            shifted_img = img_pad[sy_start:sy_end, sx_start:sx_end]
            shifted_w   = w_pad[sy_start:sy_end, sx_start:sx_end]
            
            total_color += shifted_img
            total_weight += shifted_w
            
        # 4. 计算平均值
        # 避免除以0
        total_weight[total_weight == 0] = 1.0 
        avg_color = total_color / total_weight
        
        # 5. 只更新 border_mask 区域
        # 注意：这里我们只把计算出来的 avg_color 填入 border_mask
        border_bool = border_mask.astype(bool)
        img_crop[border_bool] = avg_color[border_bool]
        
        # --- C. 更新 Mask (剥掉一层) ---
        # 这一圈修好了，它们变成了下一圈的“已知参考”
        remaining_mask[border_bool] = 0

    # 3. 放回原图
    # 记得转回 uint8
    img[y1:y2, x1:x2] = img_crop
    
    return np.clip(img, 0, 255).astype(np.uint8)




def inpaint_scipy_griddata(img, mask):
    """
    使用 SciPy 的 griddata 进行插值修复。
    这是数学层面最接近 '完美' 去除简单背景水印的方法。
    它会根据周围的像素拟合一个曲面来填补空洞。
    """
    # 预处理
    if mask.ndim == 3: mask = mask[:,:,0]
    mask = mask > 0
    
    # 坐标网格
    h, w = img.shape[:2]
    y, x = np.mgrid[0:h, 0:w]
    
    # 1. 找到所有 "已知像素" (背景) 的坐标和颜色
    # 为了速度，我们只取 mask 边缘一圈的背景，不用全图背景
    # 这里为了代码简单，取反 mask
    known_mask = ~mask
    
    # 优化：只取 ROI，不然全图计算太慢
    rows, cols = np.where(mask)
    if len(rows) == 0: return img
    pad = 5
    y_min, y_max = max(0, rows.min()-pad), min(h, rows.max()+pad)
    x_min, x_max = max(0, cols.min()-pad), min(w, cols.max()+pad)
    
    # 切片
    img_roi = img[y_min:y_max, x_min:x_max]
    mask_roi = mask[y_min:y_max, x_min:x_max]
    
    # 准备数据点
    # points: 已知像素坐标 (N, 2)
    # values: 已知像素颜色 (N, 3)
    known_y, known_x = np.where(~mask_roi)
    target_y, target_x = np.where(mask_roi)
    
    # 如果背景点太多，随机采样一下加速 (griddata O(N log N))
    if len(known_y) > 2000:
        idx = np.random.choice(len(known_y), 2000, replace=False)
        points = np.column_stack((known_y[idx], known_x[idx]))
        values = img_roi[known_y[idx], known_x[idx]]
    else:
        points = np.column_stack((known_y, known_x))
        values = img_roi[known_y, known_x]
        
    xi = np.column_stack((target_y, target_x))
    
    # 2. 核心：插值
    # method='linear' 速度快效果好 (线性过渡)
    # method='cubic' 最平滑但慢，且可能在边缘产生伪影
    interpolated = griddata(points, values, xi, method='linear')
    
    # 3. 填回图像
    img_roi[target_y, target_x] = interpolated
    img[y_min:y_max, x_min:x_max] = img_roi
    
    return img