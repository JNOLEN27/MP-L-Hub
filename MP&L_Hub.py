import subprocess
import sys

def ensure_dependencies():
    requiredpackages = ['PyQt5','bcrypt', 'pandas','numpy']
    missingpackages = []
    for package in requiredpackages:
        try:
            __import__(package)
            print(f"{package} is available")
        except ImportError:
            missingpackages.append(package)
            print(f"{package} is missing")
    
    if missingpackages:
        print(f"\nInstalling missing packages: {missingpackages}")
        for package in missingpackages:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])
                print(f"Installed {package}")
            except subprocess.CalledProcessError:
                print(f"Failed to install {package}")
                print("Please install manually: pip install --user " + package)
                return False
    return True

if __name__ == "__main__":
    print("Checking dependencies...")
    if ensure_dependencies():
        print("All dependencies ready!")
        from app.launcher.main import main
        main()
    else:
        print("Some dependencies failed to install. Please install manually.")
        input("Press Enter to exit...")
