import streamlit as st
import pandas as pd
import numpy as np
import random
from faker import Faker
import string
import tempfile
import os
from datetime import datetime, timedelta
import xlsxwriter
from io import BytesIO, StringIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import json
import ollama
import matplotlib.pyplot as plt
import seaborn as sns
from typing import Optional, Dict, Any

# Initialize Faker
fake = Faker()

def set_seed(seed):
    """Set random seed for reproducibility"""
    random.seed(seed)
    np.random.seed(seed)
    return seed

# Default constants
DEFAULT_DEPARTMENTS = ['HR', 'Finance', 'IT', 'Marketing', 'Operations', 'Sales', 'R&D', 'Customer Support']
DEFAULT_ROLES = {
    'HR': ['HR Manager', 'Recruiter', 'HR Generalist', 'Training Specialist', 'Compensation Analyst'],
    'Finance': ['CFO', 'Financial Analyst', 'Accountant', 'Auditor', 'Financial Planner'],
    'IT': ['CTO', 'Software Engineer', 'Data Scientist', 'DevOps Engineer', 'IT Support'],
    'Marketing': ['CMO', 'Marketing Specialist', 'Content Writer', 'SEO Specialist', 'Social Media Manager'],
    'Operations': ['COO', 'Operations Manager', 'Project Manager', 'Logistics Coordinator'],
    'Sales': ['Sales Director', 'Account Executive', 'Sales Representative', 'Business Development'],
    'R&D': ['Research Scientist', 'Product Manager', 'UX Designer', 'Research Assistant'],
    'Customer Support': ['Support Manager', 'Customer Service Rep', 'Technical Support']
}
DEFAULT_INSURANCE_PROVIDERS = ['Aetna', 'Blue Cross', 'Cigna', 'UnitedHealthcare', 'Kaiser', 'Humana']
DEFAULT_TEAM_NAMES = [f'Team {chr(65 + i)}' for i in range(26)]  # Team A to Team Z

# Generate random employee IDs with some duplicates
def generate_employee_ids(n, duplicate_percentage=5, id_prefix='EMP', id_start=100000, id_end=999999):
    """Generate employee IDs with a specified percentage of duplicates
    
    Args:
        n: Number of employee IDs to generate
        duplicate_percentage: Percentage of IDs that should be duplicates
        id_prefix: Either a string prefix or a dictionary of {prefix: weight} for random selection
        id_start: Starting number for the numeric part of the ID
        id_end: Ending number for the numeric part of the ID
    """
    num_duplicates = int(n * (duplicate_percentage / 100))
    num_unique = n - num_duplicates
    
    # Generate unique IDs first
    if id_end - id_start + 1 < num_unique:
        raise ValueError(f"Range too small for {num_unique} unique IDs. Please increase the ID range.")
    
    # Generate unique random numbers within the specified range
    unique_numbers = random.sample(range(id_start, id_end + 1), num_unique)
    
    # Handle both single prefix and weighted random prefixes
    if isinstance(id_prefix, dict) and len(id_prefix) > 0:
        # Extract prefixes and weights from the dictionary
        prefixes = list(id_prefix.keys())
        weights = list(id_prefix.values())
        
        # Normalize weights to sum to 1
        total_weight = sum(weights)
        if total_weight == 0:
            weights = [1/len(weights)] * len(weights)  # Equal weights if all zeros
        else:
            weights = [w/total_weight for w in weights]
        
        # Generate random prefixes based on weights
        random_prefixes = random.choices(prefixes, weights=weights, k=num_unique)
        ids = [f'{prefix}{num:06d}' for prefix, num in zip(random_prefixes, unique_numbers)]
    else:
        # Single prefix mode
        prefix = str(id_prefix) if id_prefix else 'EMP'
        ids = [f'{prefix}{num:06d}' for num in unique_numbers]
    
    # Add duplicates if needed
    if num_duplicates > 0 and ids:  # Make sure we have some IDs to duplicate
        duplicate_ids = random.choices(ids, k=num_duplicates)
        ids.extend(duplicate_ids)
    
    # Shuffle to distribute duplicates
    random.shuffle(ids)
    return ids

def generate_employee_data(num_employees, departments, roles, insurance_providers, team_names, 
                         duplicate_percentage=5, missing_data_percentage=5, 
                         id_prefix='EMP', id_start=100000, id_end=999999, **kwargs):
    """Generate employee data with the specified parameters
    
    Args:
        num_employees: Number of employees to generate
        departments: List of department names
        roles: Dictionary mapping departments to role lists
        insurance_providers: List of insurance provider names
        team_names: List of team names
        duplicate_percentage: Percentage of duplicate employee IDs
        missing_data_percentage: Percentage of missing data
        id_prefix: Prefix for employee IDs (e.g., 'EMP', 'DEV')
        id_start: Starting number for employee IDs
        id_end: Ending number for employee IDs
    """
    # Generate employee IDs with some duplicates
    employee_ids = generate_employee_ids(
        num_employees, 
        duplicate_percentage,
        id_prefix=id_prefix,
        id_start=id_start,
        id_end=id_end
    )
    
    # Generate data with the specified parameters
    data = {
        'Employee_ID': employee_ids,
        'First_Name': [fake.first_name() for _ in range(num_employees)],
        'Middle_Name': [fake.first_name() if random.random() < 0.7 else np.nan 
                       for _ in range(num_employees)],
        'Last_Name': [fake.last_name() for _ in range(num_employees)],
        'Department': [random.choice(departments) for _ in range(num_employees)],
        'Role': [],  # Will be filled based on department
        'Years_In_Company': np.round(np.random.gamma(shape=2, scale=2, size=num_employees), 1),
        'Insurance_Provider': random.choices(insurance_providers, 
                                          k=num_employees),
        'Work_Floor': np.random.randint(1, 11, size=num_employees),
        'Days_In_Office': np.random.choice([1, 2, 3, 4, 5], 
                                         size=num_employees, 
                                         p=[0.1, 0.2, 0.3, 0.25, 0.15]),
        'Team': [random.choice(team_names) for _ in range(num_employees)],
        'Hire_Date': [fake.date_between(start_date='-10y', end_date='today') 
                      for _ in range(num_employees)],
        'Salary': np.round(np.random.lognormal(mean=11, sigma=0.4, size=num_employees), 2)
    }
    
    # Set roles based on department
    data['Role'] = [random.choice(roles.get(dept, ['Not Specified'])) for dept in data['Department']]
    
    # Create some missing values
    if missing_data_percentage > 0:
        for col in data:
            if col not in ['Employee_ID', 'Hire_Date']:  # Don't add missing values to these columns
                mask = np.random.random(size=num_employees) < (missing_data_percentage / 100)
                data[col] = [np.nan if mask[i] else val for i, val in enumerate(data[col])]
    
    # Create a DataFrame
    df = pd.DataFrame(data)
    
    # Ensure Employee_ID is the first column
    cols = ['Employee_ID'] + [col for col in df.columns if col != 'Employee_ID']
    df = df[cols]
    
    return df

def create_excel_with_multiple_sheets(num_employees, sheet_names, departments, roles, insurance_providers, team_names,
                                    duplicate_percentage=5, missing_data_percentage=5, output_format='xlsx',
                                    id_prefix='EMP', id_start=100000, id_end=999999):
    """Create an Excel file with multiple sheets of employee data"""
    # Create a BytesIO object to store the file in memory
    output = BytesIO()
    
    # Set the file extension and engine based on the output format
    if output_format == 'xlsx':
        engine = 'xlsxwriter'
    else:
        engine = 'openpyxl'
    
    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet_name in sheet_names:
            # Generate data for this sheet
            df = generate_employee_data(
                num_employees=num_employees,
                departments=departments,
                roles=roles,
                insurance_providers=insurance_providers,
                team_names=team_names,
                duplicate_percentage=duplicate_percentage,
                missing_data_percentage=missing_data_percentage,
                id_prefix=id_prefix,
                id_start=id_start,
                id_end=id_end
            )
            
            # Write the data to the Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Only apply formatting for xlsx files (xlsxwriter specific)
            if output_format == 'xlsx':
                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                
                # Add a header format
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#4F81BD',
                    'font_color': 'white',
                    'border': 1
                })
                
                # Format the header row
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Set column widths
                for i, col in enumerate(df.columns):
                    # Find the maximum length of the column
                    max_length = max(df[col].astype(str).apply(len).max(), len(str(col))) + 2
                    # Set the column width
                    worksheet.set_column(i, i, min(max_length, 25))
                
                # Add conditional formatting for duplicates in Employee_ID
                if duplicate_percentage > 0:
                    duplicate_format = workbook.add_format({
                        'bg_color': '#FFC7CE',
                        'font_color': '#9C0006'
                    })
                    
                    # Apply conditional formatting to highlight duplicates in Employee_ID
                    worksheet.conditional_format(1, 0, len(df), 0, {
                        'type': 'duplicate',
                        'format': duplicate_format
                    })
                
                # Add a table with autofilter
                worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {
                    'columns': [{'header': col} for col in df.columns],
                    'style': 'Table Style Medium 2',
                    'autofilter': True
                })
    
    # Get the Excel file data
    output.seek(0)
    return output

def generate_ai_insights(df: pd.DataFrame) -> Dict[str, str]:
    """
    Generate AI-powered insights about the employee data using Ollama
    
    Args:
        df: DataFrame containing employee data
        
    Returns:
        Dictionary containing AI-generated insights
    """
    try:
        # Get detailed data summary
        num_employees = len(df)
        departments = df['Department'].value_counts().to_dict()
        avg_salary = df['Salary'].mean()
        
        # Create a detailed prompt that asks for comprehensive analysis
        prompt = f"""Analyze this employee dataset in detail and provide comprehensive insights.
        
        Dataset Overview:
        - Total Employees: {num_employees}
        - Departments: {', '.join(f'{k} ({v} employees)' for k, v in departments.items())}
        - Average Salary: ${avg_salary:,.2f}
        
        Please provide a detailed analysis including:
        1. Workforce Distribution:
           - Departmental breakdown and ratios
           - Team composition
           - Position distribution
        
        2. Compensation Analysis:
           - Salary ranges per department/position
           - Potential pay disparities
           - Overtime patterns
        
        3. Data Quality:
           - Missing or inconsistent data
           - Outliers in salary or other metrics
           - Data entry patterns
        
        4. HR Recommendations:
           - Staffing adjustments
           - Training needs
           - Potential retention risks
        
        Provide the analysis in a clear, structured format with sections and bullet points.
        Do not include any follow-up questions or prompts for elaboration."""
        
        # Get response from Ollama with parameters for detailed response
        response = ollama.chat(
            model='gemma3:4b',
            messages=[
                {'role': 'system', 'content': 'You are an HR analytics assistant providing detailed, structured analysis.'},
                {'role': 'user', 'content': prompt}
            ],
            options={
                'temperature': 0.2,     # More focused and deterministic
                'top_p': 0.9,          # Allow for some creativity in analysis
                'num_ctx': 4096,       # Larger context for detailed response
                'num_predict': 1024,   # Allow longer, more detailed response
                'stop': ['Would you like me to elaborate', 'Let me know if you need', 'Is there anything else']
            }
        )
        
        # Extract and clean the response
        insights = response['message']['content']
        
        # Remove any trailing questions or prompts
        stop_phrases = [
            'Would you like me to elaborate',
            'Let me know if you need',
            'Is there anything else',
            'Would you like me to analyze',
            'Please let me know if',
            'Do you need any other'
        ]
        
        for phrase in stop_phrases:
            if phrase in insights:
                insights = insights.split(phrase)[0].strip()
        
        return {
            'insights': insights,
            'status': 'success'
        }
        
    except Exception as e:
        return {
            'insights': f"Error generating AI insights: {str(e)}\n\nPlease ensure Ollama is running and the model is downloaded.",
            'status': 'error'
        }

def create_pdf_with_summary(df: pd.DataFrame, sheet_name: str) -> BytesIO:
    """
    Create a PDF with summary statistics and AI insights using ReportLab.
    
    Args:
        df: DataFrame containing the employee data
        sheet_name: Name of the sheet being processed
        
    Returns:
        BytesIO: In-memory file object containing the PDF
    """
    # Generate AI insights
    with st.spinner('🤖 Generating AI insights...'):
        ai_response = generate_ai_insights(df)
        
        # Handle the AI response which might be a dictionary
        if isinstance(ai_response, dict):
            if 'insights' in ai_response:
                ai_insights = ai_response['insights']
            else:
                ai_insights = str(ai_response)
        else:
            ai_insights = str(ai_response)
    
    # Create a BytesIO buffer to store the PDF
    buffer = BytesIO()
    
    # Create the PDF document
    doc = SimpleDocTemplate(buffer, pagesize=letter, 
                          rightMargin=72, leftMargin=72,
                          topMargin=72, bottomMargin=72)
    
    # Container for the 'Flowable' objects
    elements = []
    styles = getSampleStyleSheet()
    
    # Get or create custom styles
    def get_or_create_style(name, **kwargs):
        if name not in styles:
            styles.add(ParagraphStyle(name=name, **kwargs))
        return name
    
    # Define or get styles
    center_style = get_or_create_style('MyCenter', alignment=TA_CENTER, fontSize=16, spaceAfter=20)
    h1_style = get_or_create_style('MyHeading1', fontSize=14, spaceAfter=12, leading=16)
    h2_style = get_or_create_style('MyHeading2', fontSize=12, spaceAfter=8, leading=14)
    normal_style = get_or_create_style('MyNormal', fontSize=10, leading=12)
    get_or_create_style('MyTableHeader', fontSize=8, textColor=colors.white, 
                       backColor=colors.HexColor('#4B6EAF'), alignment=TA_CENTER)
    get_or_create_style('MyTableCell', fontSize=7, leading=8)
    
    # Title
    elements.append(Paragraph(f'Employee Data Summary - {sheet_name}', styles[center_style]))
    
    # Add AI Insights section
    elements.append(Paragraph('AI Insights', styles[h1_style]))
    
    # Add AI insights with proper formatting
    if isinstance(ai_insights, str):
        # If it's a string, split by lines
        for line in ai_insights.split('\n'):
            if line.strip().startswith('#'):  # Handle headers
                elements.append(Paragraph(line.lstrip('#').strip(), styles[h2_style]))
            elif line.strip():  # Skip empty lines
                elements.append(Paragraph(line, styles[normal_style]))
    else:
        # If it's not a string, just convert to string and add as a paragraph
        elements.append(Paragraph(str(ai_insights), styles[normal_style]))
    
    elements.append(Spacer(1, 12))
    
    # Add summary statistics
    elements.append(Paragraph('Summary Statistics', styles[h1_style]))
    
    total_employees = len(df)
    departments = df['Department'].nunique()
    teams = df['Team'].nunique()
    
    stats_data = [
        ['Total Employees:', str(total_employees)],
        ['Number of Departments:', str(departments)],
        ['Number of Teams:', str(teams)]
    ]
    
    # Create a table for statistics
    stats_table = Table(stats_data, colWidths=[150, 100])
    stats_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 10),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(stats_table)
    elements.append(Spacer(1, 12))
    
    # Department distribution
    elements.append(Paragraph('Department Distribution', styles[h2_style]))
    
    dept_counts = df['Department'].value_counts().reset_index()
    dept_counts.columns = ['Department', 'Count']
    dept_data = [['Department', 'Employee Count']] + dept_counts.values.tolist()
    
    # Create a table for department distribution
    dept_table = Table(dept_data, colWidths=[200, 100])
    dept_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4B6EAF')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E8EEF4')),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
    ]))
    elements.append(dept_table)
    elements.append(Spacer(1, 12))
    
    # Add data table
    elements.append(Paragraph('Employee Data', styles[h1_style]))
    
    # Convert dataframe to a list of lists for the table
    data = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # Create the table with the data
    table = Table(data, repeatRows=1)
    
    # Style the table
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4B6EAF')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E8EEF4')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    # Add the table to the elements
    elements.append(table)
    
    # Build the PDF
    doc.build(elements)
    
    # Reset buffer position to the beginning
    buffer.seek(0)
    return buffer

def main():
    """Main function to run the Streamlit app"""
    # Set page config
    st.set_page_config(
        page_title="Employee Data Generator",
        page_icon="📊",
        layout="wide"
    )
    
    # Initialize session state for output format if it doesn't exist
    if 'output_format' not in st.session_state:
        st.session_state.output_format = 'Excel'
    
    # Custom CSS
    st.markdown("""
    <style>
    .main-header {font-size: 2rem; color: #1f77b4; text-align: center; margin-bottom: 1rem;}
    .sub-header {font-size: 1.25rem; color: #2c3e50; margin: 1.5rem 0 1rem 0;}
    .stButton>button {width: 100%; margin: 0.5rem 0;}
    .stDownloadButton>button {width: 100%; margin: 0.5rem 0;}
    </style>
    """, unsafe_allow_html=True)
    
    # Title and description
    st.markdown("<h1 class='main-header'>📊 Employee Data Generator</h1>", unsafe_allow_html=True)
    st.markdown("""
    Generate realistic employee data with customizable parameters. 
    Perfect for testing, demos, and development purposes.
    """)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Basic settings
        st.subheader("Basic Settings")
        num_employees = st.slider("Number of employees per sheet", 10, 1000, 200)
        num_sheets = st.slider("Number of sheets", 1, 12, 3)
        
        # Data quality settings
        st.subheader("Data Quality")
        duplicate_percentage = st.slider("Percentage of duplicate Employee IDs", 0, 20, 5)
        missing_data_percentage = st.slider("Percentage of missing data", 0, 20, 5)
        
        # Output settings
        st.subheader("Output Settings")
        output_format = st.radio("Output format", ["Excel", "PDF with Summary"], index=0 if st.session_state.get('output_format', 'Excel') == 'Excel' else 1)
        st.session_state.output_format = output_format
        
        # Custom data settings
        st.subheader("Custom Data")
        custom_departments = st.text_area("Departments (one per line)", "\n".join(DEFAULT_DEPARTMENTS))
        custom_teams = st.text_area("Team names (one per line)", "\n".join(DEFAULT_TEAM_NAMES))
        custom_insurance = st.text_area("Insurance providers (comma separated)", 
                                     ", ".join(DEFAULT_INSURANCE_PROVIDERS))
                                      
        # Employee ID settings
        st.subheader("Employee ID Settings")
        
        # Add toggle for random prefixes
        use_random_prefixes = st.checkbox("Use random prefixes", value=False)
        
        if use_random_prefixes:
            # Default prefixes with weights
            default_prefixes = ["EMP", "DEV", "SAL", "MGR", "TMP", "INT"]
            
            # Allow users to add/remove prefixes and set weights
            st.write("Configure Prefixes and Weights:")
            
            # Start with default prefixes if none exist in session state
            if 'prefix_weights' not in st.session_state:
                st.session_state.prefix_weights = {p: 1 for p in default_prefixes}
            
            # Add new prefix
            col1, col2 = st.columns([3, 1])
            with col1:
                new_prefix = st.text_input("Add new prefix", "")
            with col2:
                st.write("")
                if st.button("Add Prefix"):
                    if new_prefix and new_prefix not in st.session_state.prefix_weights:
                        st.session_state.prefix_weights[new_prefix] = 1
                    st.experimental_rerun()
            
            # Display current prefixes with weights
            prefixes = list(st.session_state.prefix_weights.keys())
            weights = list(st.session_state.prefix_weights.values())
            
            # Create columns for prefix inputs
            cols = st.columns(4)
            
            for i, (prefix, weight) in enumerate(zip(prefixes, weights)):
                with cols[i % 4]:
                    new_weight = st.number_input(
                        f"{prefix} weight",
                        min_value=1,
                        max_value=100,
                        value=weight,
                        step=1,
                        key=f"weight_{prefix}"
                    )
                    if new_weight != weight:
                        st.session_state.prefix_weights[prefix] = new_weight
                    
                    if st.button(f"❌ {prefix}", key=f"del_{prefix}"):
                        if len(st.session_state.prefix_weights) > 1:
                            del st.session_state.prefix_weights[prefix]
                            st.experimental_rerun()
            
            # Store the prefix selection function
            id_prefix = st.session_state.prefix_weights
        else:
            # Single prefix mode
            id_prefix = st.text_input("ID Prefix", "EMP")
        
        # ID range settings
        col1, col2 = st.columns(2)
        with col1:
            id_start = st.number_input("Starting ID", min_value=1, value=100000, step=1)
        with col2:
            id_end = st.number_input("Ending ID", min_value=id_start+1, value=999999, step=1)
    
    # Process custom data
    departments = [d.strip() for d in custom_departments.split('\n') if d.strip()]
    team_names = [t.strip() for t in custom_teams.split('\n') if t.strip()]
    insurance_providers = [i.strip() for i in custom_insurance.split(',') if i.strip()]
    
    # Use default roles for now (could be made customizable)
    roles = DEFAULT_ROLES
    
    # Generate sheet names based on quarters if multiple sheets
    if num_sheets == 1:
        sheet_names = ["Employee_Data"]
    elif num_sheets == 4:
        sheet_names = [f"Q{i+1}_2024" for i in range(4)]
    elif num_sheets == 12:
        sheet_names = [f"{month}_2024" for month in [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]]
    else:
        sheet_names = [f"Sheet_{i+1}" for i in range(num_sheets)]
    
    # Generate data button
    if st.button("🚀 Generate Employee Data"):
        with st.spinner("Generating data... This may take a moment..."):
            try:
                # Set seed for reproducibility
                set_seed(42)
                
                # Generate the employee data
                if output_format == 'PDF with Summary':
                    # For PDF, we'll use a single sheet
                    df = generate_employee_data(
                        num_employees=num_employees,
                        departments=departments,
                        roles=roles,
                        insurance_providers=insurance_providers,
                        team_names=team_names,
                        duplicate_percentage=duplicate_percentage,
                        missing_data_percentage=missing_data_percentage,
                        id_prefix=id_prefix,
                        id_start=id_start,
                        id_end=id_end
                    )
                    output_file = create_pdf_with_summary(df, sheet_name=sheet_names[0])
                    file_ext = 'pdf'
                    mime_type = 'application/pdf'
                    download_label = "💾 Download PDF File"
                else:
                    # For Excel, generate multiple sheets
                    output_file = create_excel_with_multiple_sheets(
                        num_employees=num_employees,
                        sheet_names=sheet_names,
                        departments=departments,
                        roles=roles,
                        insurance_providers=insurance_providers,
                        team_names=team_names,
                        duplicate_percentage=duplicate_percentage,
                        missing_data_percentage=missing_data_percentage,
                        output_format='xlsx',
                        id_prefix=id_prefix,
                        id_start=id_start,
                        id_end=id_end
                    )
                    file_ext = 'xlsx'
                    mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    download_label = "💾 Download Excel File"
                
                # Create a download button with shorter filename
                timestamp = datetime.now().strftime("%y%m%d_%H%M")
                filename = f"EmpData_{timestamp}.{file_ext}"
                
                st.success("✅ Data generated successfully!")
                
                # Show a preview of the first sheet
                st.subheader("📊 Data Preview (First 5 Rows)")
                if output_format == 'PDF with Summary':
                    st.dataframe(df.head())
                else:
                    preview_df = pd.read_excel(output_file, sheet_name=sheet_names[0])
                    st.dataframe(preview_df.head())
                
                # Show statistics
                st.subheader("📈 Data Statistics")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Employees", f"{num_employees:,}")
                with col2:
                    st.metric("Number of Sheets", len(sheet_names) if output_format == 'Excel' else 1)
                with col3:
                    st.metric("Duplicate IDs", f"{int(num_employees * (duplicate_percentage / 100)):,}")
                
                # Download button
                st.download_button(
                    label=download_label,
                    data=output_file,
                    file_name=filename,
                    mime=mime_type
                )
                
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
    
    # Add some documentation
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        ### How to use this tool:
        1. **Configure** the data settings in the sidebar
        2. Select your preferred **Output Format** (Excel or PDF with Summary)
        3. Click the **Generate Employee Data** button
        4. **Download** the generated file
        
        ### Features:
        - **Multiple Output Formats**:
          - **Excel**: Standard spreadsheet format with multiple sheets
          - **PDF with Summary**: Professional report with summary statistics and data tables
        - **Customizable Employee IDs**:
          - Single prefix mode (e.g., EMP1001, EMP1002)
          - Random prefix mode with custom weights (e.g., EMP1001, DEV1002, SAL1003)
          - Custom ID number ranges
        - **Customizable Data**: Adjust departments, teams, and insurance providers
        - **Control Data Quality**: Set the percentage of duplicates and missing data
        - **Multiple Sheets (Excel only)**: Generate data across multiple sheets (e.g., by quarter or month)
        - **Realistic Data**: Names, roles, and other fields are generated realistically
        
        ### Output Format Details:
        - **Excel**:
          - Multiple sheets supported
          - Auto-formatted tables with filters
          - Conditional formatting for duplicates
          - Auto-adjusted column widths
          
        - **PDF with Summary**:
          - Professional cover page
          - Summary statistics and charts
          - Data tables with proper pagination
          - Clean, print-ready formatting
        
        ### Employee ID Configuration:
        - **Single Prefix Mode**:
          1. Uncheck "Use random prefixes"
          2. Enter your desired prefix (e.g., EMP, DEV, SAL)
          
        - **Random Prefix Mode**:
          1. Check "Use random prefixes"
          2. Adjust weights to control frequency (higher = more common)
          3. Add/remove prefixes as needed
          
        - **ID Number Range**:
          - Set the starting and ending numbers for the numeric part of IDs
          - Ensure the range is large enough for your number of employees
          
        ### Tips:
        - For large datasets, be patient as generation may take some time
        - The Excel file includes formatting and filtering for easy exploration
        - The PDF format is ideal for sharing reports with management
        - You can customize the data by editing the text areas in the sidebar
        - Save your favorite prefix combinations by noting down the weights
        - The system ensures unique IDs even with random prefixes
        
        ### Note:
        - When using PDF output, only the first sheet will be included in the summary
        - For multi-sheet exports, use the Excel format
        """)

if __name__ == "__main__":
    main()
