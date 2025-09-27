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
from typing import Optional, Dict, Any, Tuple
from functools import lru_cache

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

@st.cache_data(ttl=3600)  # Cache for 1 hour
def generate_employee_data(num_employees, departments, roles, insurance_providers, team_names, 
                         duplicate_percentage=5, missing_data_percentage=5, 
                         id_prefix='EMP', id_start=100000, id_end=999999, _cache_key=None, **kwargs):
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

@st.cache_resource(ttl=3600)  # Cache for 1 hour
def get_ollama_model():
    """Initialize and cache the Ollama model"""
    try:
        # Check if Ollama is running
        try:
            # Add debug information
            st.sidebar.info("Checking Ollama installation...")
            
            # Get list of models with more detailed error handling
            try:
                models = ollama.list()
                st.sidebar.json(models)  # Debug output
            except Exception as e:
                st.sidebar.error(f"Error calling ollama.list(): {str(e)}")
                st.sidebar.warning("Please ensure Ollama is running. Try running 'ollama serve' in your terminal.")
                return None
            
            # Debug output
            st.sidebar.info(f"Ollama response: {models}")
            
            # Handle different response formats
            if isinstance(models, dict) and 'models' in models:
                model_list = models['models']
            elif isinstance(models, list):
                model_list = models
            else:
                st.sidebar.warning(f"Unexpected Ollama response format: {models}")
                return None
            
            if not model_list:
                st.sidebar.warning("No models found in Ollama. Please download a model first.")
                st.sidebar.info("Try running: ollama pull gemma3:4b")
                return None
            
            # Extract model names with better error handling
            model_names = []
            for model in model_list:
                try:
                    if isinstance(model, dict) and 'name' in model:
                        model_names.append(model['name'])
                except Exception as e:
                    st.sidebar.warning(f"Error processing model: {model}. Error: {str(e)}")
            
            st.sidebar.success(f"Found models: {', '.join(model_names) if model_names else 'None'}")
            
            if not model_names:
                st.sidebar.warning("No valid models found in Ollama. Please download a model first.")
                st.sidebar.info("Try running: ollama pull gemma3:4b")
                return None
            
            # Try to use preferred models in order
            preferred_models = ['gemma3:4b', 'llama3', 'mistral', 'gpt-oss:120b', 'gpt-oss:20b']
            
            for model in preferred_models:
                if model in model_names:
                    st.sidebar.success(f"Using model: {model}")
                    return model
            
            # If no preferred model found, use the first available one
            st.sidebar.info(f"Using first available model: {model_names[0]}")
            return model_names[0]
            
        except Exception as e:
            st.error(f"Error communicating with Ollama: {str(e)}")
            st.warning("""
                Please ensure Ollama is installed and running. You can start it by running:
                ```
                ollama serve
                ```
                In a separate terminal window.
            """)
            return None
            
    except Exception as e:
        st.error(f"Unexpected error initializing Ollama: {str(e)}")
        return None

def generate_ai_insights(df: pd.DataFrame) -> Dict[str, str]:
    """
    Generate detailed AI-powered insights about the employee data using Ollama
    
    Args:
        df: DataFrame containing employee data
        
    Returns:
        Dictionary containing AI-generated insights or error message
    """
    try:
        # Get key metrics
        num_employees = len(df)
        dept_counts = df['Department'].value_counts().to_dict()
        avg_salary = df['Salary'].mean()
        
        # Get model name with caching
        model_name = get_ollama_model()
        if not model_name:
            # Generate a basic report without AI if Ollama is not available
            basic_report = f"""
            # Employee Data Analysis Report
            
            ## Executive Summary
            This is a basic analysis of your employee dataset. For AI-powered insights, 
            please ensure Ollama is installed and running with at least one model.
            
            ## Key Metrics
            - Total Employees: {num_employees}
            - Departments: {', '.join(f'{k} ({v})' for k, v in dept_counts.items())}
            - Average Salary: ${avg_salary:,.2f}
            
            ## How to Enable AI Insights
            1. Install Ollama from https://ollama.ai/
            2. Download a model: `ollama pull gemma3:4b`
            3. Run Ollama in the background: `ollama serve`
            4. Refresh this page
            """
            return {
                'insights': basic_report,
                'status': 'warning',
                'message': 'Ollama not available. Using basic report.'
            }
        
        # Detailed report prompt
        prompt = f"""Generate a comprehensive one-page HR report for this employee dataset.
        
        Report Structure:
        1. Executive Summary (2-3 sentences)
        2. Workforce Composition
           - Department distribution
           - Position analysis
           - Key demographic insights
        
        3. Compensation Analysis
           - Salary distribution by department/level
           - Pay equity highlights
           - Overtime analysis if available
        
        4. Key Findings & Recommendations
           - 3-5 prioritized recommendations
           - Potential risks/opportunities
        
        Key Metrics:
        - Total Employees: {num_employees}
        - Department Breakdown: {', '.join(f'{k} ({v})' for k, v in dept_counts.items())}
        - Average Salary: ${avg_salary:,.2f}
        
        Make it professional, data-driven, and action-oriented.
        Use clear section headers and bullet points for readability."""
        
        # Parameters for detailed response
        params = {
            'temperature': 0.3,
            'top_p': 0.9,
            'num_ctx': 2048,  # Reduced context window for faster response
            'num_predict': 768,  # Reduced max tokens for faster response
            'stop': ['End of Report', 'Page 2', 'Would you like']
        }
        
        # Get response from Ollama with progress
        progress_text = "Generating AI insights..."
        progress_bar = st.progress(0, text=progress_text)
        
        try:
            response = ollama.chat(
                model=model_name,
                messages=[
                    {'role': 'system', 'content': 'You are a senior HR analyst creating detailed, professional reports.'},
                    {'role': 'user', 'content': prompt}
                ],
                options=params,
                stream=True
            )
            
            # Stream the response
            insights = []
            for chunk in response:
                if 'message' in chunk and 'content' in chunk['message']:
                    content = chunk['message']['content']
                    insights.append(content)
                    # Update progress based on content length (simple heuristic)
                    progress = min(0.9, len(''.join(insights)) / 2000)  # Assume ~2000 chars for full response
                    progress_bar.progress(progress, text=f"Generating AI insights... {int(progress*100)}%")
            
            insights = ''.join(insights).strip()
            progress_bar.progress(1.0, text="AI insights generated!")
            
            return {
                'insights': insights,
                'status': 'success'
            }
            
        except Exception as e:
            progress_bar.error(f"Error generating AI insights: {str(e)}")
            raise
            
    except Exception as e:
        return {
            'insights': f"Error generating AI insights: {str(e)}\n\nPlease ensure Ollama is running and the model is downloaded.",
            'status': 'error'
        }
    finally:
        # Ensure progress bar is cleared
        if 'progress_bar' in locals():
            progress_bar.empty()

@st.cache_data(ttl=3600)  # Cache for 1 hour
def create_pdf_with_summary(df: pd.DataFrame, sheet_name: str) -> BytesIO:
    """
    Create a PDF with summary statistics and AI insights using ReportLab.
    
    Args:
        df: DataFrame containing the employee data
        sheet_name: Name of the sheet being processed
        
    Returns:
        BytesIO: In-memory file object containing the PDF
    """
    # Create a progress bar
    progress_bar = st.progress(0, text="Starting PDF generation...")
    
    try:
        # Define colors
        PRIMARY_COLOR = colors.HexColor('#2c3e50')  # Dark blue
        SECONDARY_COLOR = colors.HexColor('#3498db')  # Light blue
        ACCENT_COLOR = colors.HexColor('#e74c3c')  # Red
        LIGHT_GRAY = colors.HexColor('#f5f5f5')
        
        # Create a fresh stylesheet without any predefined styles
        styles = type('Styles', (), {})()
        
        # Define our base style as a ParagraphStyle
        base_style = ParagraphStyle(
            name='BaseStyle',
            fontName='Helvetica',
            leading=12,
            spaceAfter=6,
            spaceBefore=0,
            alignment=0,  # Left
            textColor=colors.black
        )
        
        # Title style
        styles.title = ParagraphStyle(
            name='MyTitle',
            parent=base_style,
            fontSize=24,
            textColor=PRIMARY_COLOR,
            spaceAfter=12,
            alignment=1,  # Center
            fontName='Helvetica-Bold'
        )
        
        # Subtitle style
        styles.subtitle = ParagraphStyle(
            name='MySubtitle',
            parent=base_style,
            fontSize=10,
            textColor=colors.grey,
            spaceAfter=20,
            alignment=1,  # Center
            fontName='Helvetica-Oblique'
        )
        
        # Header1 style
        styles.header1 = ParagraphStyle(
            name='MyHeader1',
            parent=base_style,
            fontSize=16,
            textColor=PRIMARY_COLOR,
            spaceAfter=6,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        )
        
        # Header2 style
        styles.header2 = ParagraphStyle(
            name='MyHeader2',
            parent=base_style,
            fontSize=14,
            textColor=SECONDARY_COLOR,
            spaceAfter=6,
            spaceBefore=10,
            fontName='Helvetica-Bold'
        )
        
        # Normal style
        styles.normal = ParagraphStyle(
            name='MyNormal',
            parent=base_style,
            fontSize=10,
            leading=12,
            spaceAfter=6,
            fontName='Helvetica'
        )
        
        # Bullet style
        styles.bullet = ParagraphStyle(
            name='MyBullet',
            parent=base_style,
            bulletIndent=0,
            leftIndent=20,
            spaceAfter=4,
            fontSize=10,
            leading=12,
            fontName='Helvetica'
        )
        
        # Highlight style
        styles.highlight = ParagraphStyle(
            name='MyHighlight',
            parent=base_style,
            backColor=LIGHT_GRAY,
            borderWidth=1,
            borderColor=colors.lightgrey,
            borderPadding=5,
            spaceAfter=10,
            fontName='Helvetica'
        )
        
        # Generate AI insights with progress update
        progress_bar.progress(10, text="Generating AI insights (this may take a minute)...")
        ai_response = generate_ai_insights(df)
        
        # Handle the AI response
        ai_available = False
        if isinstance(ai_response, dict):
            if 'insights' in ai_response:
                ai_insights = ai_response['insights']
                ai_available = ai_response.get('status') != 'warning'
            else:
                ai_insights = "AI insights could not be generated at this time."
        else:
            ai_insights = str(ai_response)
            
        progress_bar.progress(30, text="AI insights generated. Creating PDF...")
        
        # Create a BytesIO buffer to store the PDF
        buffer = BytesIO()
        
        # Create the PDF document with professional margins
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=72,    # Extra space for header
            bottomMargin=36
        )
        
        # Container for the 'Flowable' objects
        elements = []
        
        # Add header with logo and title
        header_table = Table([
            [Paragraph("EMPLOYEE DATA ANALYSIS REPORT", styles.title)],
            [Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')} | Sheet: {sheet_name}", styles.subtitle)]
        ], colWidths=[500])
        
        header_table.setStyle(TableStyle([
            ('BOTTOMPADDING', (0, 0), (0, 0), 6),
            ('BOTTOMPADDING', (0, 1), (0, 1), 12),
            ('LINEBELOW', (0, 1), (0, 1), 1, PRIMARY_COLOR),
            ('PADDING', (0, 0), (-1, -1), 0)
        ]))
        
        elements.append(header_table)
        elements.append(Spacer(1, 15))
        
        # Key Metrics Section
        elements.append(Paragraph("KEY METRICS", styles.header1))
        
        # Calculate metrics
        num_employees = len(df)
        departments = df['Department'].value_counts().to_dict()
        avg_salary = df['Salary'].mean()
        max_salary = df['Salary'].max()
        min_salary = df['Salary'].min()
        
        # Create metrics boxes
        metrics_data = [
            ["Total Employees", f"{num_employees:,}", "Total headcount"],
            ["Avg Salary", f"${avg_salary:,.2f}", "Average across all departments"],
            ["Salary Range", f"${min_salary:,.0f} - ${max_salary:,.0f}", "Min to max salary"],
            ["Departments", f"{len(departments)}", "Unique departments"]
        ]
        
        # Create metrics table
        metrics_table = Table(metrics_data, colWidths=[120, 100, 280])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), PRIMARY_COLOR),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 1), (-1, -1), LIGHT_GRAY),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.white),
            ('BOX', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('PADDING', (0, 0), (-1, -1), 8),
        ]))
        
        elements.append(metrics_table)
        elements.append(Spacer(1, 20))
        
        # Department Distribution Section
        elements.append(Paragraph("DEPARTMENT OVERVIEW", styles.header1))
        
        # Create department distribution table
        dept_data = [["Department", "Employees", "% of Total", "Avg Salary"]]
        for dept, count in sorted(departments.items()):
            dept_avg = df[df['Department'] == dept]['Salary'].mean()
            dept_pct = (count / num_employees) * 100
            dept_data.append([
                dept,
                str(count),
                f"{dept_pct:.1f}%",
                f"${dept_avg:,.2f}"
            ])
        
        # Add department distribution table
        dept_table = Table(dept_data, colWidths=[180, 80, 80, 100])
        dept_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), PRIMARY_COLOR),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('BOX', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('PADDING', (0, 0), (-1, -1), 5),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, LIGHT_GRAY])
        ]))
        
        elements.append(dept_table)
        elements.append(Spacer(1, 20))
        
        # Add AI insights section
        progress_bar.progress(60, text="Adding AI insights to PDF...")
        
        if ai_available:
            elements.append(Paragraph("AI-POWERED INSIGHTS", styles.header1))
            
            # Add a nice header for AI section
            ai_header = Table([
                ["INTELLIGENT ANALYSIS"],
                ["Powered by AI - Generated on " + datetime.now().strftime('%B %d, %Y')]
            ], colWidths=[500])
            
            ai_header.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, 0), PRIMARY_COLOR),
                ('TEXTCOLOR', (0, 0), (0, 0), colors.whitesmoke),
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (0, 0), 12),
                ('BOTTOMPADDING', (0, 0), (0, 0), 4),
                ('TEXTCOLOR', (0, 1), (0, 1), colors.grey),
                ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Oblique'),
                ('FONTSIZE', (0, 1), (0, 1), 8),
                ('BOTTOMPADDING', (0, 1), (0, 1), 8),
            ]))
            
            elements.append(ai_header)
            
            # Clean and format AI insights
            paragraphs = [p.strip() for p in ai_insights.split('\n') if p.strip()]
            
            for para in paragraphs[:12]:  # Limit to 12 paragraphs
                if para.startswith('#'):  # Handle headers
                    level = min(len(para.split(' ')[0]), 3)  # Limit to H1-H3
                    header_text = para.lstrip('#').strip()
                    if level == 1:
                        elements.append(Paragraph(header_text, styles.header1))
                    elif level == 2:
                        elements.append(Paragraph(header_text, styles.header2))
                    else:
                        elements.append(Paragraph(header_text, styles.normal))
                elif para.startswith(('•', '-', '*')):  # Handle bullet points
                    bullet_text = para[1:].strip()
                    elements.append(Paragraph(f"• {bullet_text}", styles.bullet))
                elif para.strip():
                    elements.append(Paragraph(para, styles.normal))
                    elements.append(Spacer(1, 4))
        else:
            # Add a helpful message when AI isn't available
            elements.append(Paragraph("ENHANCE YOUR REPORT WITH AI", styles.header1))
            
            ai_notice = [
                "Unlock powerful AI insights with Ollama:",
                "• Install Ollama from https://ollama.ai/",
                "• Download a model: `ollama pull gemma3:4b`",
                "• Run Ollama: `ollama serve` in a terminal",
                "• Refresh this page to see AI-powered analysis"
            ]
            
            for line in ai_notice:
                if line.startswith('•'):
                    elements.append(Paragraph(line, styles.bullet))
                else:
                    elements.append(Paragraph(line, styles.normal))
                    elements.append(Spacer(1, 4))
        
        # Add footer with page numbers
        def add_page_numbers(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            page_num = canvas.getPageNumber()
            text = f"Page {page_num}"
            canvas.drawRightString(doc.width + doc.leftMargin, 20, text)
            canvas.drawString(doc.leftMargin, 20, f"{sheet_name} | {datetime.now().strftime('%Y-%m-%d')}")
            canvas.restoreState()
        
        # Build the PDF with footer
        progress_bar.progress(80, text="Finalizing PDF...")
        doc.build(elements, onFirstPage=add_page_numbers, onLaterPages=add_page_numbers)
        
        # Move to the beginning of the BytesIO buffer
        buffer.seek(0)
        
        progress_bar.progress(100, text="PDF generation complete!")
        
        return buffer
        
    except Exception as e:
        progress_bar.error(f"Error generating PDF: {str(e)}")
        raise
    finally:
        # Ensure progress bar is cleared
        if 'progress_bar' in locals():
            progress_bar.empty()

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
