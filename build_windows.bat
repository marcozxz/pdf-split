@echo off
chcp 65001 >nul
echo ========================================
echo PDF智能处理工具 - Windows打包脚本
echo ========================================
echo.

echo [1/4] 检查Python环境...
python --version
if %errorlevel% neq 0 (
    echo 错误: 未找到Python，请先安装Python 3.9+
    pause
    exit /b 1
)

echo.
echo [2/4] 安装打包工具PyInstaller...
pip install pyinstaller

echo.
echo [3/4] 安装项目依赖...
pip install PyMuPDF rapidocr-onnxruntime==1.3.0 onnxruntime==1.15.1 numpy==1.26.4 opencv-python==4.8.1.78

echo.
echo [4/4] 开始打包...
pyinstaller --clean pdf_unified_tool.spec

echo.
echo ========================================
echo 打包完成！
echo 可执行文件位置: dist\PDF智能处理工具.exe
echo ========================================
pause
