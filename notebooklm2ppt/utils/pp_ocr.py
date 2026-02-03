# 综合OCR和布局解析脚本
import os
import base64
import requests
import json
from typing import Dict, Optional


class PP_OCR:
    """PDF OCR和布局解析处理器"""
    
    # API配置
    API_CONFIG = {
        "PP-OCRv5": {
            "url": "https://acz08b47m6tcndc6.aistudio-app.com/ocr",
            "params": {
                "useDocOrientationClassify": False,
                "useDocUnwarping": False,
                "useTextlineOrientation": False,
            }
        },
        "PaddleOCR-VL-1.5": {
            "url": "https://f26bq3abj7sal2f8.aistudio-app.com/layout-parsing",
            "params": {
                "useDocOrientationClassify": False,
                "useDocUnwarping": False,
                "useChartRecognition": False,
            }
        },
        "PP-StructureV3": {
            "url": "https://cczfe2v4c2qb0cv9.aistudio-app.com/layout-parsing",
            "params": {
                "useDocOrientationClassify": False,
                "useDocUnwarping": False,
                "useTextlineOrientation": False,
                "useChartRecognition": False,
            }
        }
    }
    
    def __init__(self, token: str):
        """
        初始化OCR处理器
        
        参数：
            token: API访问令牌
        """
        self.token = token
        self.headers = {
            "Authorization": f"token {token}",
            "Content-Type": "application/json"
        }
    
    def process_pdf(self, file_path: str, api_type: str, output_path: str) -> Dict:
        """
        处理PDF文件，调用指定的API
        
        参数：
            file_path: PDF文件路径
            api_type: 要调用的API类型，如 "PP-OCRv5", "PaddleOCR-VL-1.5", "PP-StructureV3"
            output_path: 输出文件路径
        
        返回：
            包含API调用结果的字典
        """
        
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        if api_type not in self.API_CONFIG:
            print(f"警告: 未知的API类型 '{api_type}'，跳过")
            return {"status": "failed", "error": f"未知的API类型 '{api_type}'"}
        
        # 创建输出目录
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 读取和编码PDF文件
        with open(file_path, "rb") as file:
            file_bytes = file.read()
            file_data = base64.b64encode(file_bytes).decode("ascii")
        
        config = self.API_CONFIG[api_type]
        
        print(f"正在调用 {api_type} API...")
        
        try:
            # 构建请求负载
            payload = {
                "file": file_data,
                "fileType": 0,  # 0: PDF, 1: 图像
            }
            payload.update(config["params"])
            
            # 发送请求
            response = requests.post(config["url"], json=payload, headers=self.headers)
            
            if response.status_code != 200:
                print(f"❌ {api_type} API 失败: HTTP {response.status_code}")
                return {"status": "failed", "code": response.status_code}
            
            # 获取结果
            result = response.json().get("result")
            
            # 保存结果
            with open(output_path, "w", encoding="utf-8") as json_file:
                json.dump(result, json_file, ensure_ascii=False, indent=4)
            
            print(f"✓ {api_type} API 处理成功，结果已保存到: {output_path}")
            return {"status": "success", "output_file": output_path}
                
        except Exception as e:
            print(f"❌ {api_type} API 出错: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def merge_results(self, vl_path: str, v5_path: str, output_path: str) -> None:
        """
        合并PaddleOCR-VL和PP-OCRv5的结果
        
        参数：
            vl_path: PaddleOCR-VL-1.5结果文件路径
            v5_path: PP-OCRv5结果文件路径
            output_path: 合并后的输出文件路径
        """
        print(f"Loading {vl_path}...")
        with open(vl_path, 'r', encoding='utf-8') as f:
            vl_data = json.load(f)
            
        print(f"Loading {v5_path}...")
        with open(v5_path, 'r', encoding='utf-8') as f:
            v5_data = json.load(f)
            
        # 创建合并后的字典
        merged_result = {}
        
        # 1. 复制所有 vl_data 的内容 (主要包含 layoutParsingResults)
        merged_result.update(vl_data)
        
        # 2. 复制 v5_data 中的 ocrResults
        if 'ocrResults' in v5_data:
            merged_result['ocrResults'] = v5_data['ocrResults']
            print(f"Added ocrResults from {v5_path}")
        
        # 3. 如果 vl_data 没 ocrResults 而 v5 有，确保不会丢失
        # 反之，如果 v5 有 layout 而 vl 没，也可以考虑合并，
        # 但以 vl 的 layout 为准。
        
        # 4. 检查 metadata
        if 'dataInfo' not in merged_result and 'dataInfo' in v5_data:
            merged_result['dataInfo'] = v5_data['dataInfo']
            
        # 保存结果
        print(f"Saving merged result to {output_path}...")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(merged_result, f, ensure_ascii=False, indent=4)
        
        print("Done!")
    
    def process_with_vl_and_v5(self, file_path: str, output_dir: str, overwrite: bool = False) -> Optional[str]:
        """
        使用 PaddleOCR-VL-1.5 + PP-OCRv5 组合处理PDF
        
        参数：
            file_path: PDF文件路径
            output_dir: 输出目录
        
        返回：
            合并结果文件路径，如果失败返回None
        """
        vl_output = os.path.join(output_dir, "result_vl.json")
        v5_output = os.path.join(output_dir, "result_v5.json")
        merged_output = os.path.join(output_dir, "result.json")
        if overwrite is False and os.path.exists(merged_output):
            print(f"合并结果已存在，跳过处理: {merged_output}")
            return merged_output
        
        # 调用 PaddleOCR-VL-1.5
        result = self.process_pdf(file_path, "PaddleOCR-VL-1.5", vl_output)
        if result.get("status") != "success":
            print("PaddleOCR-VL-1.5 处理失败，跳过合并步骤")
            return None
        
        # 调用 PP-OCRv5
        result = self.process_pdf(file_path, "PP-OCRv5", v5_output)
        if result.get("status") != "success":
            print("PP-OCRv5 处理失败，跳过合并步骤")
            return None
        
        # 合并结果
        self.merge_results(vl_output, v5_output, merged_output)
        return merged_output
    
    def process_with_structure(self, file_path: str, output_dir: str, overwrite: bool = False) -> Optional[str]:
        """
        使用 PP-StructureV3 处理PDF
        
        参数：
            file_path: PDF文件路径
            output_dir: 输出目录
        
        返回：
            结果文件路径，如果失败返回None
        """
        structure_output = os.path.join(output_dir, "result_structure.json")
        if overwrite is False and os.path.exists(structure_output):
            print(f"结构化结果已存在，跳过处理: {structure_output}")
            return structure_output
        result = self.process_pdf(file_path, "PP-StructureV3", structure_output)
        
        if result.get("status") != "success":
            print("PP-StructureV3 处理失败")
            return None
        
        return structure_output


def main():
    """主函数"""
    # 加载环境变量
    from dotenv import load_dotenv
    load_dotenv()
    TOKEN = os.getenv("TOKEN")
    
    if not TOKEN:
        raise ValueError("未找到 TOKEN 环境变量，请检查 .env 文件")

    # 配置文件路径
    file_path = r"examples/Floyd_算法的动态规划之魂.pdf"
    output_dir = "output/Floyd_算法的动态规划之魂"
    
    # 初始化处理器
    processor = PP_OCR(TOKEN)
    
    # 选择处理方法
    methods = ['PaddleOCR-VL-1.5+PP-OCRv5', 'PP-StructureV3']
    method = methods[1]  # 选择方法
    
    if method == 'PaddleOCR-VL-1.5+PP-OCRv5':
        processor.process_with_vl_and_v5(file_path, output_dir)
    elif method == 'PP-StructureV3':
        processor.process_with_structure(file_path, output_dir)
    else:
        raise ValueError(f"未知的方法: {method}")
    


if __name__ == "__main__":
    main()
