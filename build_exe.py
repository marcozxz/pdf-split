"""
PDF工具集 - 打包脚本
运行方式: python build_exe.py [分类拆分|按页拆分|全部]
生成目录: dist/
"""

import subprocess
import sys
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TOOLS = {
    "分类拆分": {
        "script": "pdf_split_tool.py",
        "name": "PDF分类拆分工具",
    },
    "按页拆分": {
        "script": "pdf_page_split_tool.py",
        "name": "PDF按页拆分命名工具",
    },
}

COMMON_HIDDEN_IMPORTS = [
    "fitz",
    "rapidocr_onnxruntime",
    "rapidocr_onnxruntime.main",
    "rapidocr_onnxruntime.utils",
    "rapidocr_onnxruntime.utils.infer_engine",
    "rapidocr_onnxruntime.utils.parse_parameters",
    "rapidocr_onnxruntime.ch_ppocr_det",
    "rapidocr_onnxruntime.ch_ppocr_det.text_detect",
    "rapidocr_onnxruntime.ch_ppocr_rec",
    "rapidocr_onnxruntime.ch_ppocr_rec.text_recognize",
    "rapidocr_onnxruntime.ch_ppocr_cls",
    "rapidocr_onnxruntime.ch_ppocr_cls.text_cls",
    "PIL",
    "numpy",
    "onnxruntime",
]

# Collect data files (config, models) from rapidocr_onnxruntime
COMMON_COLLECT_DATA = [
    "rapidocr_onnxruntime",
]


def build_one(tool_key):
    tool = TOOLS[tool_key]
    script_path = os.path.join(SCRIPT_DIR, tool["script"])
    exe_name = tool["name"]

    if not os.path.exists(script_path):
        print(f"❌ 找不到脚本: {script_path}")
        return False

    print(f"\n{'='*50}")
    print(f"🔨 打包: {exe_name}")
    print(f"{'='*50}")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        f"--name={exe_name}",
        "--windowed",
        "--onefile",
        "--noconfirm",
        "--clean",
    ]
    for imp in COMMON_HIDDEN_IMPORTS:
        cmd.append(f"--hidden-import={imp}")
    for data in COMMON_COLLECT_DATA:
        cmd.append(f"--collect-data={data}")
    cmd.append(script_path)

    result = subprocess.run(cmd, cwd=SCRIPT_DIR)

    exe_path = os.path.join(SCRIPT_DIR, "dist", f"{exe_name}.exe")
    if result.returncode == 0 and os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"\n✅ 打包成功！")
        print(f"   文件: {exe_path}")
        print(f"   大小: {size_mb:.1f} MB")
        return True
    else:
        print(f"\n❌ 打包失败")
        return False


def main():
    # Check PyInstaller
    try:
        import PyInstaller
        print(f"✅ PyInstaller {PyInstaller.__version__} 已安装")
    except ImportError:
        print("📦 正在安装 PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller", "--quiet"])
        print("✅ PyInstaller 安装完成")

    # Parse argument
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg in TOOLS:
            targets = [arg]
        elif arg == "全部":
            targets = list(TOOLS.keys())
        else:
            print(f"用法: python build_exe.py [分类拆分|按页拆分|全部]")
            print(f"可用工具: {', '.join(TOOLS.keys())}, 全部")
            return
    else:
        print("PDF工具集 - 打包脚本")
        print("=" * 50)
        print("可用工具:")
        for key, tool in TOOLS.items():
            print(f"  {key}: {tool['name']}")
        print()
        print("用法: python build_exe.py [分类拆分|按页拆分|全部]")
        print()
        choice = input("请选择要打包的工具 (输入编号或名称，或输入'全部'): ").strip()
        if choice in TOOLS:
            targets = [choice]
        elif choice in ("全部", "all", "a"):
            targets = list(TOOLS.keys())
        else:
            print("❌ 无效选择")
            return

    results = {}
    for t in targets:
        results[t] = build_one(t)

    print(f"\n{'='*50}")
    print("📦 打包结果汇总")
    print(f"{'='*50}")
    for t, success in results.items():
        status = "✅ 成功" if success else "❌ 失败"
        print(f"  {TOOLS[t]['name']}: {status}")

    if all(results.values()):
        print(f"\n🎉 所有工具打包完成！输出目录: {os.path.join(SCRIPT_DIR, 'dist')}")


if __name__ == "__main__":
    main()
