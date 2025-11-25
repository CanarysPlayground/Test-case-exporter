"""
Export test cases to PDF document (.pdf)
Requires: reportlab
Install: pip install reportlab
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime

def create_test_case_pdf(test_cases, output_file="test_cases.pdf"):
    """
    Convert test cases to a formatted PDF document
    
    Args:
        test_cases: List of dictionaries containing test case information
        output_file: Output filename for the PDF document
    """
    doc = SimpleDocTemplate(output_file, pagesize=letter,
                           rightMargin=0.5*inch, leftMargin=0.5*inch,
                           topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    # Container for the 'Flowable' objects
    elements = []
    
    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=12,
        spaceBefore=12,
        fontName='Helvetica-Bold'
    )
    
    normal_style = styles['Normal']
    
    # Add title
    title = Paragraph("Test Cases Documentation", title_style)
    elements.append(title)
    
    # Add metadata
    metadata = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Total Test Cases: {len(test_cases)}"
    elements.append(Paragraph(metadata, normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    # Process each test case
    for idx, test_case in enumerate(test_cases, 1):
        # Test case header
        tc_heading = Paragraph(f"Test Case #{idx}: {test_case.get('name', 'Unnamed Test')}", heading_style)
        elements.append(tc_heading)
        
        # Create table data
        data = [
            ['Field', 'Details'],
            ['ID', test_case.get('id', f'TC_{idx:03d}')],
            ['Description', test_case.get('description', 'N/A')],
            ['Preconditions', test_case.get('preconditions', 'N/A')],
            ['Test Steps', test_case.get('steps', 'N/A')],
            ['Expected Result', test_case.get('expected_result', 'N/A')],
            ['Actual Result', test_case.get('actual_result', '')],
            ['Status', test_case.get('status', 'Not Executed')],
            ['Priority', test_case.get('priority', 'Medium')],
            ['Test Type', test_case.get('test_type', 'Functional')]
        ]
        
        # Convert data to Paragraphs for better text wrapping
        formatted_data = []
        for row in data:
            formatted_row = [
                Paragraph(str(cell), normal_style if i == 1 else styles['Heading3'])
                for i, cell in enumerate(row)
            ]
            formatted_data.append(formatted_row)
        
        # Create table
        table = Table(formatted_data, colWidths=[1.5*inch, 5.5*inch])
        
        # Style the table
        table.setStyle(TableStyle([
            # Header row
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            
            # First column (field names)
            ('BACKGROUND', (0, 1), (0, -1), colors.HexColor('#D9E2F3')),
            ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
            
            # All cells
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('PADDING', (0, 0), (-1, -1), 8),
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Add page break after each test case (except the last one)
        if idx < len(test_cases):
            elements.append(PageBreak())
    
    # Build PDF
    doc.build(elements)
    print(f"✓ PDF document saved: {output_file}")
    return output_file


# Example usage
if __name__ == "__main__":
    # Sample test cases structure
    sample_test_cases = [
        {
            "id": "TC_001",
            "name": "User Login with Valid Credentials",
            "description": "Verify that user can login with valid username and password",
            "preconditions": "User account exists in the system",
            "steps": "1. Navigate to login page\n2. Enter valid username\n3. Enter valid password\n4. Click Login button",
            "expected_result": "User is successfully logged in and redirected to dashboard",
            "status": "Pass",
            "priority": "High",
            "test_type": "Functional"
        },
        {
            "id": "TC_002",
            "name": "User Login with Invalid Password",
            "description": "Verify error message when invalid password is entered",
            "preconditions": "User account exists in the system",
            "steps": "1. Navigate to login page\n2. Enter valid username\n3. Enter invalid password\n4. Click Login button",
            "expected_result": "Error message 'Invalid credentials' is displayed",
            "status": "Fail",
            "priority": "High",
            "test_type": "Negative Testing"
        }
    ]
    
    # Create PDF file
    create_test_case_pdf(sample_test_cases, "test_cases_output.pdf")
