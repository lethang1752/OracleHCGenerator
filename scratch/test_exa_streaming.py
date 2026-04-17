import sys
import tarfile
import os
import shutil
from pathlib import Path
import json

print("Starting test script...")

# Add src to path
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

print(f"Python sys.path: {sys.path}")

from PyQt5.QtCore import QCoreApplication
app = QCoreApplication([])

try:
    from src.utils.exawatcher_runner import ExaWatcherGraphGenerator
    print("Import ExaWatcherGraphGenerator successful.")
except Exception as e:
    print(f"FAILED to import ExaWatcherGraphGenerator: {e}")
    sys.exit(1)

def create_mock_html(path, var_name, data):
    with open(path, 'w', encoding='utf-8') as f:
        f.write(f"<html><body><script>var {var_name} = {json.dumps(data)};</script></body></html>")

def create_mock_archive(file_path, html_files):
    print(f"Creating archive: {file_path}")
    with tarfile.open(file_path, "w:bz2") as tar:
        for p in html_files:
            tar.add(p, arcname=p.name)

def test_streaming():
    test_dir = Path("scratch/test_exa")
    print(f"Setting up test directory: {test_dir}")
    if test_dir.exists():
        try:
            shutil.rmtree(test_dir)
        except Exception as e:
            print(f"Warning cleanup: {e}")
    
    test_dir.mkdir(parents=True, exist_ok=True)

    # 1. Create Mock Data
    print("Creating mock data...")
    db_html = test_dir / "db_mp.html"
    mem_html = test_dir / "db_meminfo.html"
    cell_html = test_dir / "cell_iosummary.html"
    
    times = [f"2026-04-17T{h:02d}:00:00" for h in range(10)]
    cpu_data = {
        "xAxis": times,
        "series": {
            "all": [
                {"id": "idl", "name": "Idle", "items": [90, 85, 80, 70, 60, 50, 60, 70, 80, 90]}
            ]
        }
    }
    with open(db_html, 'w', encoding='utf-8') as f:
        f.write(f"var xAxis = {json.dumps(times)}; var series = {json.dumps(cpu_data['series'])};")

    mem_data = {
        "xAxis": times,
        "data": {
            "osmem": [{"name": "Used", "items": [10, 11, 12, 13, 14, 15, 14, 13, 12, 11]}]
        }
    }
    with open(mem_html, 'w', encoding='utf-8') as f:
        f.write(f"var xAxis = {json.dumps(times)}; var data = {json.dumps(mem_data['data'])};")

    io_data = {
        "xAxis": times,
        "data": {
            "flash": {
                "iops": [{"id": "iops", "items": [1000, 1100, 1200, 1300, 1400, 1500, 1400, 1300, 1200, 1100]}]
            }
        }
    }
    with open(cell_html, 'w', encoding='utf-8') as f:
        f.write(f"var xAxis = {json.dumps(times)}; var data = {json.dumps(io_data['data'])};")

    # 2. Package into .tar.bz2
    db_archive = test_dir / "db_logs.tar.bz2"
    cell_archive = test_dir / "cell_logs.tar.bz2"
    
    create_mock_archive(db_archive, [db_html, mem_html])
    create_mock_archive(cell_archive, [cell_html])

    # 3. Run Generator
    print("Running Generator...")
    output_dir = test_dir / "output"
    generator = ExaWatcherGraphGenerator(
        db_node_source=str(db_archive),
        cell_node_source=str(cell_archive),
        output_folder=str(output_dir)
    )
    
    generator.progress.connect(lambda msg: print(f"Progress: {msg}"))
    
    try:
        generator.run()
    except Exception as e:
        print(f"CRITICAL ERROR during generator.run(): {e}")
        import traceback
        traceback.print_exc()

    # 4. Verify results
    print("Verifying files...")
    expected_files = ["Exa_Cpu_Idle.png", "Exa_Mem_OS.png", "Exa_IO_Summary.png"]
    found_all = True
    for f in expected_files:
        if (output_dir / f).exists():
            print(f"[PASSED] Found {f}")
        else:
            print(f"[FAILED] MISSING {f}")
            found_all = False
    
    if found_all:
        print("TEST SUCCESSFUL")
    else:
        print("TEST FAILED")

if __name__ == "__main__":
    try:
        test_streaming()
    except Exception as e:
        print(f"Top-level error: {e}")
        import traceback
        traceback.print_exc()
