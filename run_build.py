"""Build script that sets up VS environment and runs pip install"""
import subprocess
import sys
import os

def main():
    # VS vcvars batch file path
    vcvars = r"C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"

    # Command to run
    build_cmd = f'"{vcvars}" && cd /d D:\\Code\\cpyhwpx && py -3 -m pip install . --no-build-isolation -v'

    print("Running build with VS2022 environment...")
    print(f"Command: {build_cmd}")
    print("-" * 60)

    # Run via cmd
    result = subprocess.run(
        f'cmd /c "{build_cmd}"',
        shell=True,
        cwd=r"D:\Code\cpyhwpx"
    )

    return result.returncode

if __name__ == "__main__":
    sys.exit(main())
