"""
Unified script to export test cases to Word, Excel, and PDF formats
Install requirements: pip install python-docx openpyxl reportlab
"""

import json
import sys
from pathlib import Path

def load_test_cases_from_file(input_file):
    """Load test cases from various file formats (JSON, C#, Python)"""
    ext = Path(input_file).suffix.lower()
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        if ext == '.json':
            try:
                return json.loads(content)
            except json.JSONDecodeError as e:
                print(f"✗ Error: Invalid JSON format in '{input_file}'\n   {e}")
                return None
        elif ext == '.cs':
            # Enhanced C# test case extraction: supports [Fact], [Test], [TestCase], and method names
            import re
            test_cases = []
            # Find all test methods with [Fact] or [Test] attributes
            # Optionally, extract summary comments above methods
            pattern = r'(?:\s*///\s*(?P<summary>.*))?\s*\[\s*(Fact|Test|TestCase)[^\]]*\]\s*(?:public|private|protected)?\s*async\s*Task\s+(?P<method>\w+)'
            for match in re.finditer(pattern, content):
                tc_name = match.group('method')
                summary = match.group('summary')
                tc_id = f'TC_{len(test_cases)+1:03d}'
                display_name = summary if summary else tc_name.replace('_', ' ')
                test_cases.append({'id': tc_id, 'name': display_name})
            # If no matches, fallback to [TestCase] and comment-based extraction
            if not test_cases:
                for match in re.finditer(r'\[TestCase\s*\(([^)]*)\)\]|//\s*TestCase:\s*(.*)', content):
                    if match.group(1):
                        args = [x.strip(' "') for x in match.group(1).split(',')]
                        tc_id = args[0] if len(args) > 0 else None
                        tc_name = args[1] if len(args) > 1 else None
                    else:
                        tc_id, tc_name = None, match.group(2)
                    test_cases.append({'id': tc_id or f'TC_{len(test_cases)+1:03d}', 'name': tc_name or 'Unnamed Test'})
            return test_cases
        elif ext == '.py':
            # Basic Python test case extraction (looks for docstrings or comments with 'TestCase:')
            import re
            test_cases = []
            # Example: # TestCase: TC_001, Login Test
            for match in re.finditer(r'#\s*TestCase:\s*(.*)', content):
                parts = [x.strip() for x in match.group(1).split(',')]
                tc_id = parts[0] if len(parts) > 0 else None
                tc_name = parts[1] if len(parts) > 1 else None
                test_cases.append({'id': tc_id or f'TC_{len(test_cases)+1:03d}', 'name': tc_name or 'Unnamed Test'})
            return test_cases
        else:
            print(f"✗ Unsupported file format: {ext}")
            return None
    except FileNotFoundError:
        print(f"✗ Error: File '{input_file}' not found!")
        print(f"   Please provide a valid test case file.")
        return None
    except Exception as e:
        print(f"✗ Error loading file: {e}")
        return None

def export_all_formats(test_cases, base_filename="test_cases"):
    """Export test cases to all formats"""
    print("\n🚀 Starting export to all formats...\n")
    
    results = {}
    
    # Export to Word
    try:
        from export_to_word import create_test_case_word
        word_file = f"{base_filename}.docx"
        create_test_case_word(test_cases, word_file)
        results['word'] = word_file
    except ImportError:
        print("✗ Word export failed: export_to_word.py not found")
    except Exception as e:
        print(f"✗ Word export failed: {e}")
    
    # Export to Excel
    try:
        from export_to_excel import create_test_case_excel
        excel_file = f"{base_filename}.xlsx"
        create_test_case_excel(test_cases, excel_file)
        results['excel'] = excel_file
    except ImportError:
        print("✗ Excel export failed: export_to_excel.py not found")
    except Exception as e:
        print(f"✗ Excel export failed: {e}")
    
    # Export to PDF
    try:
        from export_to_pdf import create_test_case_pdf
        pdf_file = f"{base_filename}.pdf"
        create_test_case_pdf(test_cases, pdf_file)
        results['pdf'] = pdf_file
    except ImportError:
        print("✗ PDF export failed: export_to_pdf.py not found")
    except Exception as e:
        print(f"✗ PDF export failed: {e}")
    
    print("\n✅ Export completed!")
    if results:
        print(f"Generated {len(results)} file(s):")
        for format_type, filename in results.items():
            print(f"  • {format_type.upper()}: {filename}")
    else:
        print("⚠️  No files were generated. Check error messages above.")
    
    return results

# Main execution - THIS IS THE PART YOU NEED TO UNDERSTAND
if __name__ == "__main__":
    print("=" * 60)
    print("📄 Test Case Export Tool")
    print("=" * 60)

    # Accept input file as argument or default to copilot_test_cases.json
    if len(sys.argv) > 1:
        input_filename = sys.argv[1]
    else:
        input_filename = 'copilot_test_cases.json'

    print(f"\n📂 Looking for: {input_filename}")

    test_cases = load_test_cases_from_file(input_filename)

    if test_cases:
        print(f"✅ Successfully loaded {len(test_cases)} test case(s)")

        # Show summary of loaded test cases
        print("\n📋 Test Cases Found:")
        for idx, tc in enumerate(test_cases, 1):
            tc_name = tc.get('name', 'Unnamed Test')
            tc_id = tc.get('id', f'TC_{idx:03d}')
            print(f"   {idx}. [{tc_id}] {tc_name}")

        # Export to all formats
        export_all_formats(test_cases, Path(input_filename).stem + "_output")

        print("\n" + "=" * 60)
        print("✨ Done! Check the generated files in this folder.")
        print("=" * 60)
    else:
        print("\n❌ Cannot proceed without test cases.")
        print("\n💡 Make sure you provide a valid test case file (JSON, C#, Python)")

