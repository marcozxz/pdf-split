#!/bin/bash
echo "========================================"
echo "PDF智能处理工具 - macOS交叉编译Windows版"
echo "========================================"
echo ""

echo "[1/5] 检查依赖..."

# 检查是否安装 Homebrew
if ! command -v brew &> /dev/null; then
    echo "错误: 未安装Homebrew，请先安装: /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
    exit 1
fi

# 检查是否安装 Docker
if ! command -v docker &> /dev/null; then
    echo "未检测到Docker，正在安装..."
    echo "请手动下载并安装 Docker Desktop for Mac:"
    echo "https://docs.docker.com/desktop/install/mac-install/"
    echo ""
    echo "安装完成后重新运行此脚本"
    exit 1
fi

echo "✓ Docker已安装"
echo ""

echo "[2/5] 拉取Windows Python镜像..."
docker pull python:3.11-windows

echo ""
echo "[3/5] 在Docker容器中安装依赖并打包..."
docker run --rm -v "$(pwd)":/app -w /app python:3.11-windows cmd /c "pip install pyinstaller PyMuPDF rapidocr-onnxruntime==1.3.0 onnxruntime==1.15.1 numpy==1.26.4 opencv-python==4.8.1.78 && pyinstaller --clean PDF智能处理工具.spec"

echo ""
echo "[4/5] 打包完成！"
echo ""
echo "========================================"
echo "输出文件: dist/PDF智能处理工具.exe"
echo "========================================"
