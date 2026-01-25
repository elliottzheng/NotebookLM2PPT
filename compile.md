# 编译指南
- 下载源码
- 创建一个环境，只安装本项目依赖和pyinstaller
- 参考配置upx配置upx，用于压缩exe文件
- 编译
```
pyinstaller --clean -F -w -n notebooklm2ppt --optimize=2 --collect-all spire.presentation main.py 
pyinstaller -D -w -n notebooklm2ppt --optimize=2 main.py
pyinstaller --clean -D -w -n notebooklm2ppt --optimize=2 --collect-all spire.presentation main.py

python -m nuitka --standalone main.py
git log -n 10 > log.txt
```