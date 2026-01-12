import os
import sys
import shutil
import subprocess
import platform

def run_command(command, env=None):
    """运行命令"""
    print(f"执行命令: {command}")
    try:
        # 实时输出
        process = subprocess.Popen(
            command, 
            shell=True, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT,
            env=env,
            encoding='utf-8',
            errors='replace'
        )
        
        while True:
            line = process.stdout.readline()
            if line == '' and process.poll() is not None:
                break
            if line:
                print(line.strip())
                
        return process.returncode
    except Exception as e:
        print(f"执行出错: {e}")
        return 1

def main():
    print("=== 开始构建 Tungee CRM 爬虫 (GUI版) ===")
    
    # 1. 确保安装依赖
    print("\n[1/4] 检查依赖...")
    if run_command(f"{sys.executable} -m pip install -r requirements.txt") != 0:
        print("安装依赖失败")
        return

    # 2. 安装/下载 Playwright 浏览器
    print("\n[2/4] 下载 Chromium 浏览器到本地 'browsers' 目录...")
    
    # 设置 PLAYWRIGHT_BROWSERS_PATH 环境变量，让 playwright 下载到当前目录下的 browsers 文件夹
    env = os.environ.copy()
    project_dir = os.path.dirname(os.path.abspath(__file__))
    browsers_dir = os.path.join(project_dir, 'browsers')
    
    # 清理旧的 browsers 目录 (可选，为了保持干净)
    # if os.path.exists(browsers_dir):
    #     shutil.rmtree(browsers_dir)
    
    env['PLAYWRIGHT_BROWSERS_PATH'] = browsers_dir
    
    # 只安装 chromium
    cmd = f"{sys.executable} -m playwright install chromium"
    if run_command(cmd, env=env) != 0:
        print("下载浏览器失败")
        return

    # 3. 如果是 macOS，可能需要先移除 quarantine 属性 (针对 user 的情况)
    if platform.system() == 'Darwin':
        pass 

    # 4. 执行 PyInstaller 打包
    print("\n[3/4] 开始打包...")
    # 使用 --onedir 模式，因为我们需要包含 browsers 文件夹，且 --onefile 解压太慢
    # 但如果用户坚持单文件，依然需要把 browsers 文件夹放在 exe旁边
    
    # 简单的 PyInstaller 命令
    # -F: 单文件 (onefile)
    # -w: 无控制台窗口 (windowed) - 仅 Windows
    # -n: 名称
    
    pyinstaller_args = [
        "pyinstaller",
        "-F",  # 生成单文件，方便分发 (但也需要配套 browsers 文件夹)
        "--name=TungeeCrawler",
        "--clean",
        "--noconfirm",
    ]
    
    # 如果是 Windows, 加上 -w 隐藏控制台 (调试时可以去掉)
    if platform.system() == 'Windows':
        pyinstaller_args.append("-w")
    else:
        # Mac 上 -w 会生成 .app，这里我们只生成 unix 可执行文件方便侧式
        pass

    pyinstaller_args.append("crawler_gui.py")
    
    cmd = " ".join(pyinstaller_args)
    if run_command(cmd) != 0:
        print("打包失败")
        return

    # 5. 后处理
    print("\n[4/4] 后处理...")
    dist_dir = os.path.join(project_dir, 'dist')
    
    # 复制 browsers 文件夹到 dist 目录，方便用户直接打包分发
    print(f"复制 browsers 文件夹到 {dist_dir}...")
    target_browsers_dir = os.path.join(dist_dir, 'browsers')
    
    if os.path.exists(target_browsers_dir):
        shutil.rmtree(target_browsers_dir)
        
    try:
        shutil.copytree(browsers_dir, target_browsers_dir)
        print("复制成功!")
    except Exception as e:
        print(f"复制失败: {e}")

    print("\n" + "="*50)
    print("构建完成!")
    print(f"可执行文件位于: {dist_dir}")
    print(f"*** 重要提示 ***: 请务必将 'dist' 目录下的 'browsers' 文件夹与可执行文件放在同一目录下运行!")
    print("="*50)

if __name__ == "__main__":
    main()
