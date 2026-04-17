import re
import sys
import subprocess
from pathlib import Path

def bump_version(current_version):
    """
    Increments the minor version. 
    If minor reaches 9, roll-over to major.
    Example: 2.3.0 -> 2.4.0
    Example: 2.9.0 -> 3.0.0
    """
    parts = current_version.split('.')
    if len(parts) >= 2:
        try:
            major = int(parts[0])
            minor = int(parts[1])
            
            if minor >= 9:
                major += 1
                minor = 0
            else:
                minor += 1
                
            parts[0] = str(major)
            parts[1] = str(minor)
            parts[2] = "0" # Reset patch version
            return ".".join(parts)
        except ValueError:
            return current_version
    return current_version

def update_file(file_path, patterns):
    """
    Updates a file using a list of (pattern, replacement) tuples.
    """
    if not file_path.exists():
        print(f"[WARNING] File not found: {file_path}")
        return False
        
    content = file_path.read_text(encoding='utf-8')
    new_content = content
    
    for pattern, replacement in patterns:
        # Escape backslashes in replacement for re.sub literal interpretation
        safe_replacement = replacement.replace('\\', '\\\\')
        new_content = re.sub(pattern, safe_replacement, new_content)
        
    if new_content != content:
        file_path.write_text(new_content, encoding='utf-8')
        print(f"[INFO] Updated {file_path.name}")
        return True
    return False

def main():
    root = Path(__file__).parent.parent
    config_file = root / "src" / "config.py"
    readme_file = root / "README.md"
    project_map_file = root / "PROJECT_MAP.md"
    spec_onefile = root / "build_onefile.spec"
    spec_standard = root / "build.spec"
    build_bat = root / "scripts" / "build_app.bat"

    # 1. Read current version
    if not config_file.exists():
        print("[ERROR] src/config.py not found.")
        sys.exit(1)
        
    config_content = config_file.read_text(encoding='utf-8')
    version_match = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', config_content)
    
    if not version_match:
        print("[ERROR] Could not find APP_VERSION in src/config.py")
        sys.exit(1)
        
    old_version = version_match.group(1)
    new_version = bump_version(old_version)
    
    print(f"[PROCESS] Bumping version: v{old_version} -> v{new_version}")

    # 2. Update all files
    # config.py
    update_file(config_file, [
        (rf'APP_VERSION\s*=\s*["\']{re.escape(old_version)}["\']', f'APP_VERSION = "{new_version}"')
    ])
    
    # README.md
    update_file(readme_file, [
        (rf'\(v{re.escape(old_version)}\)', f'(v{new_version})'),
        (rf'Release-v{re.escape(old_version)}', f'Release-v{new_version}')
    ])
    
    # PROJECT_MAP.md
    update_file(project_map_file, [
        (rf'\(v{re.escape(old_version)}\)', f'(v{new_version})')
    ])
    
    # Spec files
    update_file(spec_onefile, [
        (rf"name='OracleHCGenerator_v[^']+'", f"name='OracleHCGenerator_v{new_version}'")
    ])
    update_file(spec_standard, [
        (rf"name='OracleHCGenerator_v[^']+'", f"name='OracleHCGenerator_v{new_version}'")
    ])
    # 3. Running PyInstaller
    print(f"[INFO] Starting PyInstaller build for v{new_version}...")
    try:
        # Use venv python to run pyinstaller if possible
        pyinstaller_cmd = [sys.executable, "-m", "PyInstaller", "--clean", "build_onefile.spec"]
        subprocess.run(pyinstaller_cmd, check=True)
        print(f"[SUCCESS] Packaging complete: OracleHCGenerator_v{new_version}.exe")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] PyInstaller failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
