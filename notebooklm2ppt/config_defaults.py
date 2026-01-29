"""
统一的默认配置常量定义
"""

# 任务处理的出厂默认设置（仅在第一次使用时使用）
DEFAULT_TASK_SETTINGS = {
    "dpi": 150,
    "ratio": 0.8,
    "inpaint": True,
    "inpaint_method": "background_smooth",  # 需要根据具体方法调整
    "image_only": False,
    "force_regenerate": False,
    "unify_font": True,
    "font_name": "Calibri",
    "page_range": ""
}

# 自动化相关设置的出厂默认值
DEFAULT_AUTOMATION_SETTINGS = {
    "delay": 0,  # 等待时间（秒）
    "timeout": 50,  # 超时时间（秒）
    "done_offset": "",  # 按钮偏移
    "calibrate": True  # 自动校准
}

# GUI 相关的默认值
DEFAULT_GUI_VALUES = {
    "output_dir": "workspace",
    "language": "zh_cn"
}

# 获取合并后的完整默认设置（考虑用户上次的设置）
def get_default_settings(output_dir="workspace", inpaint_method="background_smooth", user_last_settings=None):
    """
    获取完整的任务默认设置，优先使用用户上次的设置
    
    Args:
        output_dir: 输出目录
        inpaint_method: 修复方法（会在运行时从翻译方法中获取第一个）
        user_last_settings: 用户上次使用的设置（从 config.json 读取）
    
    Returns:
        dict: 完整的设置字典
    """
    # 先使用出厂默认值
    settings = DEFAULT_TASK_SETTINGS.copy()
    
    # 如果有用户上次的设置，覆盖默认值
    if user_last_settings:
        settings.update(user_last_settings)
    
    # 始终使用当前的输出目录
    settings["output_dir"] = output_dir
    
    # 只有当用户上次没有保存修复方法时，才使用传入的默认方法
    if not user_last_settings or 'inpaint_method' not in user_last_settings:
        settings["inpaint_method"] = inpaint_method
    
    return settings

