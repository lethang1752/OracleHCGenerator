import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(r"d:\OneDrive\AntiGravityProject\NewApplication")
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

import traceback

# Mock pyqtSignal for testing without GUI environment
class MockSignal:
    def connect(self, func): self.func = func
    def emit(self, *args): 
        if hasattr(self, 'func'): 
            try:
                self.func(*args)
            except Exception:
                traceback.print_exc()

from src.utils.exawatcher_runner import ExaWatcherGraphGenerator
# Patch the signals
ExaWatcherGraphGenerator.progress = MockSignal()
ExaWatcherGraphGenerator.finished = MockSignal()

def test_extraction():
    db_folder = r"E:\HIEPDH3\HC040326\SYS\ExtractedResults\ExaWatcher_dcexadbadm02.mbbank.com.vn_2026-03-01_00_00_00_71h59m59s\Charts.ExaWatcher.dcexadbadm02.mbbank.com.vn\2026_03_01_00_00_00_71h59m00s_0"
    # For testing, we use the same folder for IO if CELL folder is not handy, 
    # but the runner handles it.
    cell_folder = db_folder 
    output_folder = Path(r"d:\OneDrive\AntiGravityProject\NewApplication\scratch\test_exa_out")
    
    print(f"Testing ExaWatcher Generator...")
    print(f"DB  Source: {db_folder}")
    print(f"Cell Source: {cell_folder}")
    print(f"Output   : {output_folder}")
    
    gen = ExaWatcherGraphGenerator(db_folder, cell_folder, str(output_folder))
    
    # Connect signal
    gen.progress.connect(lambda msg: print(f"Progress: {msg}"))
    gen.finished.connect(lambda ok: print(f"Finished: {'SUCCESS' if ok else 'FAILED'}"))
    
    gen.run()
    
    # Check output
    if output_folder.exists():
        files = list(output_folder.glob("*.png"))
        print(f"\nGenerated Files ({len(files)}):")
        for f in files:
            print(f" - {f.name}")
    else:
        print("\n[ERROR] Output folder not created!")

if __name__ == "__main__":
    test_extraction()
