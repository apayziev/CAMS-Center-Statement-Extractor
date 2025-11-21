"""
Local build script for creating executable.
Run this to build the executable on your local machine.
"""
import subprocess
import sys
import platform

def build_executable():
    """Build executable using PyInstaller."""
    
    print("Building CAMS Center Statement Extractor executable...")
    print(f"Platform: {platform.system()}")
    
    # Base PyInstaller command
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", "CAMSExtractor",
        "gui_settings.py"
    ]
    
    # Add console option for debugging (remove --windowed if you want to see console)
    # cmd.remove("--windowed")  # Uncomment this line to show console window
    
    print(f"\nRunning: {' '.join(cmd)}\n")
    
    try:
        result = subprocess.run(cmd, check=True)
        print("\n✅ Build completed successfully!")
        print(f"Executable location: dist/CAMSExtractor{'.exe' if platform.system() == 'Windows' else ''}")
        return 0
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Build failed with error code {e.returncode}")
        return 1
    except Exception as e:
        print(f"\n❌ Error: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(build_executable())
