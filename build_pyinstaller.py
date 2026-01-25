import os
import sys
import shutil
import subprocess


def _remove_pdb_files(root_dir):
    for dirpath, _, filenames in os.walk(root_dir):
        for name in filenames:
            if name.lower().endswith(".pdb"):
                try:
                    os.remove(os.path.join(dirpath, name))
                except OSError:
                    pass


def _prune_qt_translations(root_dir):
    keep_suffixes = ("_zh_cn.qm", "_en.qm", "_en_us.qm")
    for dirpath, _, filenames in os.walk(root_dir):
        if os.path.basename(dirpath).lower() == "translations":
            for name in filenames:
                lower = name.lower()
                if lower.endswith(".qm") and not lower.endswith(keep_suffixes):
                    try:
                        os.remove(os.path.join(dirpath, name))
                    except OSError:
                        pass


def _prune_qtwebengine_locales(root_dir):
    keep_locales = {"zh-CN.pak", "en-US.pak"}
    for dirpath, _, filenames in os.walk(root_dir):
        if os.path.basename(dirpath) == "qtwebengine_locales":
            for name in filenames:
                if name not in keep_locales:
                    try:
                        os.remove(os.path.join(dirpath, name))
                    except OSError:
                        pass


def run():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(root_dir, "dist")
    build_dir = os.path.join(root_dir, "build")

    venv_python = os.path.join(root_dir, ".venv", "Scripts", "python.exe")
    if not os.path.exists(venv_python):
        raise FileNotFoundError("未找到 .venv\\Scripts\\python.exe")
    python_exe = venv_python

    icons_dir = os.path.join(root_dir, "icons")
    logo_ico = os.path.join(icons_dir, "logo.ico")

    data_sep = os.pathsep
    add_data = [
        f"version.json{data_sep}.",
        f"config{data_sep}config",
        f"plugins{data_sep}plugins",
        f"icons{data_sep}icons",
        f"ppt_assistant{data_sep}ppt_assistant",
        f"fonts{data_sep}fonts",
    ]

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
        "--hidden-import",
        "PySide6.QtXml",
        "--onedir",
        "--windowed",
        "--name",
        "Kazuha",
    ]
    if os.path.exists(logo_ico):
        cmd += ["--icon", logo_ico]
    for entry in add_data:
        cmd += ["--add-data", entry]
    cmd += ["main.py"]

    if os.path.isdir(dist_dir):
        shutil.rmtree(dist_dir, ignore_errors=True)
    if os.path.isdir(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    for extra in [
        "main.build",
        "main.dist",
        "Kazuha.build",
        "Kazuha.dist",
    ]:
        extra_path = os.path.join(root_dir, extra)
        if os.path.isdir(extra_path):
            shutil.rmtree(extra_path, ignore_errors=True)
    subprocess.check_call(cmd, cwd=root_dir)

    output_dir = os.path.join(dist_dir, "Kazuha")
    if os.path.isdir(output_dir):
        _remove_pdb_files(output_dir)
        _prune_qt_translations(output_dir)
        _prune_qtwebengine_locales(output_dir)

    if os.path.isdir(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    for extra in [
        "main.build",
        "main.dist",
        "Kazuha.build",
        "Kazuha.dist",
    ]:
        extra_path = os.path.join(root_dir, extra)
        if os.path.isdir(extra_path):
            shutil.rmtree(extra_path, ignore_errors=True)
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        pyinstaller_cache = os.path.join(local_app_data, "pyinstaller")
        if os.path.isdir(pyinstaller_cache):
            shutil.rmtree(pyinstaller_cache, ignore_errors=True)


if __name__ == "__main__":
    run()
