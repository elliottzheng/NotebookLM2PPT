import numpy as np

def compute_edge_diversity_numpy(image_cv, left, top, right, bottom, tolerance=10):
    """
    使用 Numpy 替代 DBSCAN 计算边缘颜色一致性。
    tolerance: 容差，类似于 DBSCAN 的 eps。值越大，越忽略颜色微小差异。
    """
    left, top, right, bottom = round(left), round(top), round(right), round(bottom)
    
    # 边界检查，防止切片越界
    h, w, _ = image_cv.shape
    left = max(0, left); top = max(0, top)
    right = min(w, right); bottom = min(h, bottom)

    # 提取边缘像素
    top_edge = image_cv[top:top+1, left:right]      # 注意：top-1 可能越界，改为 top
    bottom_edge = image_cv[bottom-1:bottom, left:right] 
    left_edge = image_cv[top:bottom, left:left+1]
    right_edge = image_cv[top:bottom, right-1:right]
    
    edges = [top_edge, bottom_edge, left_edge, right_edge]
    
    # 过滤掉空切片 (防止 coordinates 重合导致 crash)
    valid_edges = [e.reshape(-1, 3) for e in edges if e.size > 0]
    if not valid_edges:
        return 1.0, np.array([255, 255, 255]) # 默认返回高多样性（不填充），白色

    flatten_points = np.concatenate(valid_edges, axis=0)

    if flatten_points.shape[0] == 0:
         return 1.0, np.array([255, 255, 255])

    # --- 核心逻辑替代 DBSCAN ---
    
    # 1. 颜色量化 (整除 tolerance)，相当于把相近颜色归桶
    quantized_points = flatten_points // tolerance
    
    # 2. 统计唯一颜色和数量
    unique_colors, counts = np.unique(quantized_points, axis=0, return_counts=True)
    
    # 3. 找到占比最大的颜色
    max_count_index = np.argmax(counts)
    max_count = counts[max_count_index]
    total_count = np.sum(counts)
    
    main_ratio = max_count / total_count
    
    # 4. 获取该“桶”内的平均颜色 (或者直接还原量化前的颜色)
    # 为了准确，我们可以取属于该桶的所有原始像素的平均值
    dominant_quantized_color = unique_colors[max_count_index]
    # 创建掩码找出原始像素
    mask = np.all((flatten_points // tolerance) == dominant_quantized_color, axis=1)
    main_color = np.mean(flatten_points[mask], axis=0)
    main_color = main_color.astype(np.uint8).tolist()

    return 1 - main_ratio, main_color

def compute_edge_average_color(image_cv, left, top, right, bottom):
    """
    计算边缘的平均颜色。
    """
    left, top, right, bottom = round(left), round(top), round(right), round(bottom)
    
    h, w, _ = image_cv.shape
    left = max(0, left); top = max(0, top)
    right = min(w, right); bottom = min(h, bottom)

    top_edge = image_cv[top:top+1, left:right]
    bottom_edge = image_cv[bottom-1:bottom, left:right]
    left_edge = image_cv[top:bottom, left:left+1]
    right_edge = image_cv[top:bottom, right-1:right]
    
    edges = [top_edge, bottom_edge, left_edge, right_edge]
    valid_edges = [e.reshape(-1, 3) for e in edges if e.size > 0]
    
    if not valid_edges:
        return np.array([255, 255, 255])

    flatten_points = np.concatenate(valid_edges, axis=0)
    average_color = np.mean(flatten_points, axis=0)
    average_color = average_color.astype(np.uint8).tolist()

    return average_color
