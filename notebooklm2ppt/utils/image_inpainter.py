# from pyinpaint import Inpaint
import numpy as np
from skimage.restoration import inpaint
from PIL import Image
from .edge_diversity import compute_edge_diversity_numpy, compute_edge_average_color
import math
from .inpaint_methods import inpaint_manual, inpaint_numpy_onion, inpaint_scipy_griddata


INPAINT_METHODS = [
    {
        'id': 'background_smooth',
        'name': '智能平滑（推荐）',  # 去掉“背景色”，强调智能
        'description': '综合效果最佳，适合大多数去除文字、水印的场景'
    },
    {
        'id': 'edge_mean_smooth',
        'name': '边缘均值填充',
        'description': '取周围像素平均色填充，适合纯色或简单背景'
    },
    {
        'id': 'background',
        'name': '极速纯色填充',
        'description': '直接填充单一背景色，仅适合极简底色，速度最快'
    },
    {
        'id': 'onion',
        'name': '逐层内缩修补', # 解释“洋葱皮”的原理
        'description': '由外向内逐层修补，适合细长划痕或线条修复'
    },
    {
        'id': 'griddata',
        'name': '渐变过渡插值', # 解释“网格插值”的效果
        'description': '计算平滑的曲面过渡，适合带有渐变的背景'
    },
    {
        'id': 'skimage',
        'name': '双调和光影修补', # 给 Biharmonic 一个听起来高级的名字
        'description': '计算量大，速度较慢，但能更好保持光影连续性'
    },
]

METHOD_ID_TO_NAME = {m['id']: m['name'] for m in INPAINT_METHODS}
METHOD_NAME_TO_ID = {m['name']: m['id'] for m in INPAINT_METHODS}


def get_method_names():
    """获取所有方法的中文名列表，用于GUI下拉框"""
    return [m['name'] for m in INPAINT_METHODS]


def get_method_id(method_name_or_id):
    """将方法名或ID转换为标准ID"""
    if method_name_or_id in METHOD_ID_TO_NAME:
        return method_name_or_id
    return METHOD_NAME_TO_ID.get(method_name_or_id, 'background_smooth')


def get_method_name_from_id(method_id):
    """将ID转换为中文名"""
    return METHOD_ID_TO_NAME.get(method_id, get_method_names()[0])


def inpaint_image(image_path, output_path, inpaint_method='skimage'):
    inpaint_method = get_method_id(inpaint_method)
    
    image = Image.open(image_path)
    image_defect = np.array(image)
    
    
    # [{\"width\":240,\"top\":1530,\"height\":65,\"left\":2620}]
    r1,r2,c1,c2 = 1536,1598,2627,2863

    old_width, old_height = 2867,1600

    image_width, image_height = image_defect.shape[1], image_defect.shape[0]
    ratio = image_width / old_width

    assert abs(ratio - (image_height / old_height)) < 0.01, "图片比例不对，无法修复"

    r1 = math.floor(r1 * ratio)
    r2 = math.ceil(r2 * ratio)
    c1 = math.floor(c1 * ratio)
    c2 = math.ceil(c2 * ratio)
    if inpaint_method =='skimage':
        dtype = bool
        fill_val = True
    else:
        dtype = np.uint8
        fill_val = 255

    mask = np.zeros(image_defect.shape[:-1], dtype=dtype)
    mask[r1:r2, c1:c2] = fill_val

    edge_diversity, fill_color = compute_edge_diversity_numpy(image_defect, c1, r1, c2, r2)

    if edge_diversity < 0.1 or inpaint_method == 'background': # 直接填充完事，速度最快
        print("直接填充",edge_diversity, fill_color)
        image_defect[r1:r2, c1:c2] = fill_color
        image_result = image_defect
    elif inpaint_method == 'skimage': # 速度最慢, 效果也一般
        image_result = inpaint.inpaint_biharmonic(image_defect, mask, channel_axis=-1)
        image_result = (image_result*255).astype("uint8")
    elif inpaint_method == 'onion':  # 效果一般，效果还行
        image_result = inpaint_numpy_onion(image_defect, mask)
    elif inpaint_method == 'griddata': # 跟onion差不多
        image_result = inpaint_scipy_griddata(image_defect, mask)
    elif inpaint_method == 'background_smooth': # 速度第二快，也平滑
        image_result = inpaint_manual(image_defect, mask, fill_color, max_iter=100)
    elif inpaint_method == 'edge_mean_smooth':
        fill_color = compute_edge_average_color(image_defect, c1, r1, c2, r2)
        image_result = inpaint_manual(image_defect, mask, fill_color, max_iter=100)
    else:
        raise ValueError(f"Unknown inpaint method: {inpaint_method}")


    Image.fromarray(image_result).save(output_path)
