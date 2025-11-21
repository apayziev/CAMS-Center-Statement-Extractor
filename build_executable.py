"""
Local build script for creating executable.
Run this to build the executable on your local machine.
"""
import subprocess
import sys
import platform

def build_executable():
    """Build executable using PyInstaller with spec file."""
    
    print("Building CAMS Center Statement Extractor executable...")
    print(f"Platform: {platform.system()}")
    
    # Use the spec file for consistent builds
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "CAMSExtractor.spec"
    ]
    
    print(f"\nRunning: {' '.join(cmd)}\n")
    
    try:
        subprocess.run(cmd, check=True)
        print("\n✅ Build completed successfully!")
        print(f"Executable location: dist/CAMSExtractor{'.exe' if platform.system() == 'Windows' else ''}")
        return 0
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Build failed with error code {e.returncode}")
        return 1
    except FileNotFoundError:
        print("\n❌ Error: PyInstaller not found. Please install it with: pip install pyinstaller")
        return 1

if __name__ == "__main__":
    sys.exit(build_executable())
