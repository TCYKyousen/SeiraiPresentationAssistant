import os
import sys
import shutil
import subprocess


def run():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(root_dir, "dist")
    build_dir = os.path.join(root_dir, "build")

    python_exe = sys.executable

    icons_dir = os.path.join(root_dir, "icons")
    logo_ico = os.path.join(icons_dir, "logo.ico")

    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--exclude-module",
        "PyQt5",
        "--exclude-module",
        "setuptools",
        "--exclude-module",
        "pkg_resources",
        "--onefile",
        "--windowed",
        "--name",
        "Kazuha",
        "--add-data",
        "version.json;.",
        "--add-data",
        "config;config",
        "--add-data",
        "plugins;plugins",
        "--add-data",
        "icons;icons",
        "--add-data",
        "ppt_assistant;ppt_assistant",
        "--add-data",
        "fonts;fonts",
        "main.py",
    ]

    if os.path.exists(logo_ico):
        cmd.insert(-1, "--icon")
        cmd.insert(-1, logo_ico)

    if os.path.isdir(dist_dir):
        shutil.rmtree(dist_dir, ignore_errors=True)
    if os.path.isdir(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    spec_path = os.path.join(root_dir, "Kazuha.spec")
    if os.path.exists(spec_path):
        os.remove(spec_path)

    env = os.environ.copy()
    env["PYINSTALLER_HIDE_PKGRES"] = "1"
    subprocess.check_call(cmd, cwd=root_dir, env=env)

    exe_name = "Kazuha.exe"
    src_exe = os.path.join(dist_dir, exe_name)
    dst_exe = os.path.join(root_dir, exe_name)

    if os.path.exists(src_exe):
        if os.path.exists(dst_exe):
            os.remove(dst_exe)
        shutil.move(src_exe, dst_exe)

    if os.path.isdir(dist_dir):
        shutil.rmtree(dist_dir, ignore_errors=True)
    if os.path.isdir(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    if os.path.exists(spec_path):
        os.remove(spec_path)


if __name__ == "__main__":
    run()
