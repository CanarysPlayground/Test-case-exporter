"""
Unified script to export test cases to Word, Excel, and PDF formats
Install requirements: pip install python-docx openpyxl reportlab
"""

import json
import sys
from pathlib import Path

def load_test_cases_from_json(json_file):
    """Load test cases from a JSON file"""
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"✗ Error: File '{json_file}' not found!")
        print(f"   Please create '{json_file}' in the same folder as this script.")
        return None
    except json.JSONDecodeError as e:
        print(f"✗ Error: Invalid JSON format in '{json_file}'")
        print(f"   {e}")
        return None
    except Exception as e:
        print(f"✗ Error loading JSON file: {e}")
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
    
    # Load test cases from JSON file
    json_filename = 'copilot_test_cases.json'
    print(f"\n📂 Looking for: {json_filename}")
    
    test_cases = load_test_cases_from_json(json_filename)
    
    if test_cases:
        print(f"✅ Successfully loaded {len(test_cases)} test case(s)")
        
        # Show summary of loaded test cases
        print("\n📋 Test Cases Found:")
        for idx, tc in enumerate(test_cases, 1):
            tc_name = tc.get('name', 'Unnamed Test')
            tc_id = tc.get('id', f'TC_{idx:03d}')
            print(f"   {idx}. [{tc_id}] {tc_name}")
        
        # Export to all formats
        export_all_formats(test_cases, "copilot_test_cases_output")
        
        print("\n" + "=" * 60)
        print("✨ Done! Check the generated files in this folder.")
        print("=" * 60)
    else:
        print("\n❌ Cannot proceed without test cases.")
        print("\n💡 Make sure you have a file named 'copilot_test_cases.json'")
        print("   in the same folder with your test case data.")
