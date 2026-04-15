import sys
from pathlib import Path

# Add src to path
sys.path.append(str(Path.cwd()))

from src.generators.comprehensive_report_generator import ComprehensiveHealthcareReportGenerator
import logging

logging.basicConfig(level=logging.INFO)

def test_generation(font_option):
    output = f"test_{font_option}.docx"
    print(f"Testing font_option: {font_option} -> {output}")
    gen = ComprehensiveHealthcareReportGenerator(output, font_option=font_option)
    
    # Test data
    data = {
        'db_name': 'TESTDB',
        'nodes': []
    }
    
    # Simulate some sections
    gen.doc.add_heading("Section 1 (Level 1)", level=1)
    gen.doc.add_heading("Sub-section 1.1 (Level 2)", level=2)
    gen.doc.add_heading("Detail 1.1.1 (Level 3)", level=3)
    
    gen.doc.save(output)
    print(f"Saved {output}")

if __name__ == "__main__":
    # Create test dir if needed
    test_dir = Path("test_output")
    test_dir.mkdir(exist_ok=True)
    
    test_generation('times')
    test_generation('calibri')
    test_generation('invalid') # Should fallback
