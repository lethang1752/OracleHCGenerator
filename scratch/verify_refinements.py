import sys
import os
import traceback
from pathlib import Path

# Add project root to path
sys.path.append(str(Path(__file__).parent.parent))

from src.utils.exawatcher_runner import ExaWatcherGraphGenerator

# Mock signals for headless test
class MockSignals:
    def emit(self, msg):
        try:
            print(msg)
        except UnicodeEncodeError:
            print(msg.encode('ascii', 'replace').decode('ascii'))

def run_test():
    import logging
    # Enable logging to capture underlying errors
    logging.basicConfig(level=logging.INFO)
    db_folder = r"D:\OneDrive\AntiGravityProject\log_test\exawatcher_log\ExaWatcher_dcexadbadm02.mbbank.com.vn_2026-03-01_00_00_00_71h59m59s"
    cell_folder = r"D:\OneDrive\AntiGravityProject\log_test\exawatcher_log\ExaWatcher_dcexaceladm02.mbbank.com.vn_2026-03-01_00_00_00_71h59m59s"
    output_dir = Path(__file__).parent / "exa_output"
    
    print(f"Testing ExaWatcher Generator...")
    print(f"DB Node: {db_folder}")
    print(f"Cell Node: {cell_folder}")
    print(f"Output: {output_dir}")
    
    gen = ExaWatcherGraphGenerator(db_folder, cell_folder, str(output_dir))
    
    # Patch signals
    gen.progress = MockSignals()
    gen.finished = MockSignals()
    
    try:
        gen.run()
        print("\nGenerated files:")
        for f in output_dir.glob("*.png"):
            print(f"- {f.name}")
    except Exception:
        traceback.print_exc()

if __name__ == "__main__":
    run_test()
