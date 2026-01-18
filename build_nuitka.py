import os
import sys
import shutil
import subprocess

def run():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Nuitka build artifacts
    nuitka_build_dir = os.path.join(root_dir, "Kazuha.build")
    dist_dir = os.path.join(root_dir, "dist") # Nuitka might use this if standalone but we use onefile
    
    python_exe = sys.executable
    
    icons_dir = os.path.join(root_dir, "icons")
    logo_ico = os.path.join(icons_dir, "logo.ico")
    
    exe_name = "Kazuha.exe"
    exe_path = os.path.join(root_dir, exe_name)

    # Cleanup previous builds and artifacts
    for d in [dist_dir, nuitka_build_dir]:
        if os.path.isdir(d):
            print(f"Cleaning {d}...")
            shutil.rmtree(d, ignore_errors=True)
            
    if os.path.exists(exe_path):
        print(f"Removing existing {exe_name}...")
        os.remove(exe_path)

    # Construct Nuitka command
    cmd = [
        python_exe,
        "-m",
        "nuitka",
        "--standalone",
        "--onefile",
        "--show-progress",
        "--assume-yes-for-downloads",
        f"--output-filename={exe_name}",
        "--enable-plugin=pyside6",
        
        # Windows configuration
        "--windows-console-mode=disable",
        "--windows-uac-uiaccess", # Optional: might be needed for overlay? Original didn't specify uac-admin but uac-uiaccess helps with overlays
        
        # Excludes (matching original script's intent)
        "--nofollow-import-to=PyQt5",
        "--nofollow-import-to=setuptools",
        "--nofollow-import-to=pkg_resources",
        
        # Data files (preserving content)
        "--include-data-file=version.json=version.json",
        "--include-data-dir=config=config",
        "--include-data-dir=plugins=plugins",
        "--include-data-dir=icons=icons",
        "--include-data-dir=ppt_assistant=ppt_assistant",
        "--include-data-dir=fonts=fonts",
        
        "main.py",
    ]

    # Add icon if exists
    if os.path.exists(logo_ico):
        cmd.append(f"--windows-icon-from-ico={logo_ico}")
    else:
        print(f"Warning: Icon not found at {logo_ico}")

    print("Starting Nuitka build...")
    print("Command:", " ".join(cmd))
    
    env = os.environ.copy()
    # Ensure nuitka finds the packages in current directory
    env["PYTHONPATH"] = root_dir
    
    try:
        subprocess.check_call(cmd, cwd=root_dir, env=env)
        
        if os.path.exists(exe_path):
            print(f"Build successful! Executable created at: {exe_path}")
        else:
            print("Build command finished but executable not found.")
            sys.exit(1)
            
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error code {e.returncode}")
        sys.exit(e.returncode)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)
    finally:
        # Optional: cleanup build directory
        # if os.path.isdir(nuitka_build_dir):
        #     shutil.rmtree(nuitka_build_dir, ignore_errors=True)
        pass

if __name__ == "__main__":
    run()
