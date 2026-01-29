import win32api

# 获取屏幕尺寸
screen_width = win32api.GetSystemMetrics(0)
screen_height = win32api.GetSystemMetrics(1)

def get_effective_top_left(top_left, width, height):
    """
    计算实际生效的 top_left 坐标。
    如果图片/区域已经填满或超过屏幕，或者设置的偏移导致超出屏幕，则进行修正。
    """
    effective_top_left = list(top_left)
    
    # 如果宽度已经占满或超过屏幕宽度，则水平偏移失效
    if width >= screen_width:
        effective_top_left[0] = 0
    # 如果高度已经占满或超过屏幕高度，则垂直偏移失效
    if height >= screen_height:
        effective_top_left[1] = 0
    
    # 越界检查：确保区域不会超出屏幕右侧和下方
    if effective_top_left[0] + width > screen_width:
        effective_top_left[0] = max(0, screen_width - width)
    if effective_top_left[1] + height > screen_height:
        effective_top_left[1] = max(0, screen_height - height)
        
    return tuple(effective_top_left)
