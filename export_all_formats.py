# ... (keep all the previous code) ...

# Main execution
if __name__ == "__main__":
    # Load test cases from JSON file
    test_cases = load_test_cases_from_json('copilot_test_cases.json')
    
    if test_cases:
        print(f"\n📦 Loaded {len(test_cases)} test case(s) from JSON file")
        # Export to all formats
        export_all_formats(test_cases, "copilot_test_cases_output")
    else:
        print("❌ No test cases found. Please check your JSON file.")
