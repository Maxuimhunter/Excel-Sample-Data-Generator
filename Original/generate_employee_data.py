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

def analyze_employee_data(df: pd.DataFrame) -> Dict[str, str]:
    """
    Generate detailed insights about the employee data using local analysis
    
    Args:
        df: DataFrame containing employee data
        
    Returns:
        Dictionary containing generated insights and analysis data
    """
    try:
        # Create a copy of the dataframe to avoid modifying the original
        df_analysis = df.copy()
        
        # Basic metrics
        num_employees = len(df_analysis)
        dept_counts = df_analysis['Department'].value_counts().to_dict()
        avg_salary = df_analysis['Salary'].mean()
        max_salary = df_analysis['Salary'].max()
        min_salary = df_analysis['Salary'].min()
        salary_range = max_salary - min_salary
        median_salary = df_analysis['Salary'].median()
        salary_std = df_analysis['Salary'].std()
        
        # Calculate salary distribution
        salary_quartiles = df_analysis['Salary'].quantile([0.1, 0.25, 0.5, 0.75, 0.9]).to_dict()
        
        # Calculate salary bands
        salary_bands = {
            'Low (Bottom 25%)': (min_salary, salary_quartiles[0.25]),
            'Lower Middle (25-50%)': (salary_quartiles[0.25], salary_quartiles[0.5]),
            'Upper Middle (50-75%)': (salary_quartiles[0.5], salary_quartiles[0.75]),
            'High (Top 25%)': (salary_quartiles[0.75], max_salary)
        }
        
        # Calculate employees in each salary band
        band_counts = {}
        for band, (low, high) in salary_bands.items():
            band_counts[band] = len(df_analysis[(df_analysis['Salary'] >= low) & (df_analysis['Salary'] <= high)])
        
        # Enhanced department analysis
        dept_analysis = []
        for dept, count in dept_counts.items():
            dept_df = df_analysis[df_analysis['Department'] == dept]
            dept_avg = dept_df['Salary'].mean()
            dept_median = dept_df['Salary'].median()
            dept_min = dept_df['Salary'].min()
            dept_max = dept_df['Salary'].max()
            dept_std = dept_df['Salary'].std()
            dept_pct = (count / num_employees) * 100
            
            # Calculate salary distribution within department
            dept_quartiles = dept_df['Salary'].quantile([0.25, 0.5, 0.75]).to_dict()
            
            # Calculate gender distribution if available
            gender_dist = {}
            if 'Gender' in df_analysis.columns:
                gender_dist = dept_df['Gender'].value_counts().to_dict()
            
            dept_analysis.append({
                'name': dept,
                'count': count,
                'avg_salary': dept_avg,
                'median_salary': dept_median,
                'min_salary': dept_min,
                'max_salary': dept_max,
                'std_salary': dept_std,
                'quartiles': dept_quartiles,
                'gender_dist': gender_dist,
                'pct': dept_pct
            })
        
        # Sort departments by average salary (highest first)
        dept_analysis.sort(key=lambda x: x['avg_salary'], reverse=True)
        
        # Generate insights
        insights = ["# Employee Data Analysis Report\n"]
        
        # Executive Summary
        insights.append("## Executive Summary")
        insights.append(
            f"This comprehensive analysis covers {num_employees} employees across {len(dept_counts)} distinct departments. "
            f"The organization's compensation structure shows an average salary of ${avg_salary:,.2f} with significant variation, "
            f"ranging from ${min_salary:,.2f} to ${max_salary:,.2f}. The median salary of ${median_salary:,.2f} "
            f"indicates a {'' if median_salary > avg_salary else 'slightly '}right-skewed distribution "
            f"of compensation across the workforce."
        )
        
        # Add compensation distribution overview
        insights.append("\n### Compensation Distribution Overview")
        insights.append(f"- **Salary Range**: ${min_salary:,.2f} - ${max_salary:,.2f} (${salary_range:,.2f} range)")
        insights.append(f"- **Average (Mean) Salary**: ${avg_salary:,.2f}")
        insights.append(f"- **Median Salary**: ${median_salary:,.2f}")
        insights.append(f"- **Standard Deviation**: ${salary_std:,.2f} ({(salary_std/avg_salary)*100:.1f}% of mean)")
        
        # Add salary band analysis
        insights.append("\n### Salary Band Distribution")
        for band, count in band_counts.items():
            pct = (count / num_employees) * 100
            insights.append(f"- **{band}**: {count} employees ({pct:.1f}% of workforce)")
            
            # Add quartile information for each band
            if '25%' in band:
                low, high = salary_bands[band]
                insights.append(f"  - Salary range: ${low:,.2f} - ${high:,.2f}")
                
        # Add quartile analysis
        insights.append("\n### Salary Quartile Analysis")
        insights.append(f"- **25th Percentile (Q1)**: ${salary_quartiles[0.25]:,.2f}")
        insights.append(f"- **50th Percentile (Median)**: ${salary_quartiles[0.5]:,.2f}")
        insights.append(f"- **75th Percentile (Q3)**: ${salary_quartiles[0.75]:,.2f}")
        iqr = salary_quartiles[0.75] - salary_quartiles[0.25]
        insights.append(f"- **Interquartile Range (IQR)**: ${iqr:,.2f}")
        
        # Add outlier detection
        lower_bound = salary_quartiles[0.25] - (1.5 * iqr)
        upper_bound = salary_quartiles[0.75] + (1.5 * iqr)
        outliers = df_analysis[(df_analysis['Salary'] < lower_bound) | (df_analysis['Salary'] > upper_bound)]
        insights.append(f"- **Potential Outliers**: {len(outliers)} employees ({(len(outliers)/num_employees)*100:.1f}% of workforce) "
                      f"with salaries outside ${lower_bound:,.2f} - ${upper_bound:,.2f}")
        
        # Department Analysis
        insights.append("\n## Workforce Composition")
        insights.append("### Department Distribution")
        
        # Add overall department statistics
        insights.append(f"- **Total Departments**: {len(dept_analysis)}")
        insights.append(f"- **Average Department Size**: {num_employees / len(dept_analysis):.1f} employees")
        
        # Add detailed department analysis
        for i, dept in enumerate(dept_analysis, 1):
            comp_ratio = (dept['avg_salary'] / avg_salary) * 100
            insights.append(
                f"\n#### {i}. {dept['name']} Department\n"
                f"- **Employees**: {dept['count']} ({dept['pct']:.1f}% of workforce)\n"
                f"- **Salary Metrics**:\n"
                f"  - Average: ${dept['avg_salary']:,.2f} ({(dept['avg_salary']/avg_salary - 1)*100:+.1f}% vs company average)\n"
                f"  - Median: ${dept['median_salary']:,.2f}\n"
                f"  - Range: ${dept['min_salary']:,.2f} - ${dept['max_salary']:,.2f}\n"
                f"  - Standard Deviation: ${dept['std_salary']:,.2f} ({(dept['std_salary']/dept['avg_salary'])*100:.1f}% of mean)\n"
                f"  - Compensation Ratio: {comp_ratio:.1f}%\n"
            )
            
            # Add quartile information
            insights.append("  - **Salary Quartiles**:")
            insights.append(f"    - Q1 (25th %ile): ${dept['quartiles'].get(0.25, 0):,.2f}")
            insights.append(f"    - Q2 (Median): ${dept['quartiles'].get(0.5, 0):,.2f}")
            insights.append(f"    - Q3 (75th %ile): ${dept['quartiles'].get(0.75, 0):,.2f}")
            
            # Add gender distribution if available
            if dept['gender_dist']:
                insights.append("  - **Gender Distribution**:")
                for gender, count in dept['gender_dist'].items():
                    pct = (count / dept['count']) * 100
                    insights.append(f"    - {gender}: {count} employees ({pct:.1f}% of department)")
                    
                    # Add gender pay gap analysis
                    if 'Gender' in df_analysis.columns and len(df_analysis['Gender'].unique()) > 1:
                        gender_avg = df_analysis[(df_analysis['Department'] == dept['name']) & 
                                              (df_analysis['Gender'] == gender)]['Salary'].mean()
                        gap = ((gender_avg / dept['avg_salary']) - 1) * 100
                        insights.append(f"      - Average Salary: ${gender_avg:,.2f} ({gap:+.1f}% vs department average)")
        
        # Enhanced Position Analysis
        if 'Position' in df_analysis.columns:
            position_analysis = []
            position_counts = df_analysis['Position'].value_counts().to_dict()
            
            insights.append("\n## Position Analysis")
            insights.append("### Position Distribution")
            
            # Calculate position statistics
            for position, count in position_counts.items():
                pos_df = df_analysis[df_analysis['Position'] == position]
                pos_avg = pos_df['Salary'].mean()
                pos_median = pos_df['Salary'].median()
                pos_min = pos_df['Salary'].min()
                pos_max = pos_df['Salary'].max()
                pos_std = pos_df['Salary'].std()
                pos_pct = (count / num_employees) * 100
                
                position_analysis.append({
                    'name': position,
                    'count': count,
                    'pct': pos_pct,
                    'avg_salary': pos_avg,
                    'median_salary': pos_median,
                    'min_salary': pos_min,
                    'max_salary': pos_max,
                    'std_salary': pos_std
                })
            
            # Sort positions by average salary (highest first)
            position_analysis.sort(key=lambda x: x['avg_salary'], reverse=True)
            
            # Add detailed position analysis
            for i, pos in enumerate(position_analysis, 1):
                comp_ratio = (pos['avg_salary'] / avg_salary) * 100
                insights.append(
                    f"\n#### {i}. {pos['name']} Position\n"
                    f"- **Employees**: {pos['count']} ({pos['pct']:.1f}% of workforce)\n"
                    f"- **Salary Metrics**:\n"
                    f"  - Average: ${pos['avg_salary']:,.2f} ({(pos['avg_salary']/avg_salary - 1)*100:+.1f}% vs company average)\n"
                    f"  - Median: ${pos['median_salary']:,.2f}\n"
                    f"  - Range: ${pos['min_salary']:,.2f} - ${pos['max_salary']:,.2f}\n"
                    f"  - Standard Deviation: ${pos['std_salary']:,.2f}\n"
                    f"  - Compensation Ratio: {comp_ratio:.1f}%"
                )
        
        # Enhanced Compensation Analysis
        insights.append("\n## In-Depth Compensation Analysis")
        insights.append("### Compensation Structure")
        
        # Compensation ratios and metrics
        compa_ratio = (median_salary / avg_salary) * 100
        range_penetration = ((avg_salary - min_salary) / (max_salary - min_salary)) * 100
        
        insights.append(f"- **Compa-Ratio (Median/Mean)**: {compa_ratio:.1f}%")
        insights.append(f"- **Range Penetration**: {range_penetration:.1f}%")
        insights.append(f"- **Salary Range Spread**: {((max_salary - min_salary) / min_salary) * 100:.1f}%")
        
        # Add market position analysis
        market_position = ""
        if avg_salary > 100000:
            market_position = "Above Market"
        elif avg_salary > 70000:
            market_position = "Market Competitive"
        else:
            market_position = "Below Market"
            
        insights.append(f"- **Market Position**: {market_position} (based on average salary of ${avg_salary:,.2f})")
        
        # Add pay equity analysis
        if 'Gender' in df_analysis.columns and len(df_analysis['Gender'].unique()) > 1:
            insights.append("\n### Pay Equity Analysis")
            gender_pay = df_analysis.groupby('Gender')['Salary'].agg(['count', 'mean', 'median', 'std'])
            
            for gender, data in gender_pay.iterrows():
                pct_of_workforce = (data['count'] / num_employees) * 100
                pay_ratio = (data['mean'] / avg_salary) * 100
                
                insights.append(
                    f"- **{gender} Employees**:\n"
                    f"  - Count: {data['count']} ({pct_of_workforce:.1f}% of workforce)\n"
                    f"  - Average Salary: ${data['mean']:,.2f} ({(data['mean']/avg_salary - 1)*100:+.1f}% vs company average)\n"
                    f"  - Median Salary: ${data['median']:,.2f}\n"
                    f"  - Standard Deviation: ${data['std']:,.2f}"
                )
            
            # Calculate overall gender pay gap
            if len(gender_pay) > 1:
                genders = gender_pay.index.tolist()
                gap = ((gender_pay.loc[genders[0], 'mean'] / gender_pay.loc[genders[1], 'mean']) - 1) * 100
                insights.append(f"\n- **Gender Pay Gap**: {abs(gap):.1f}% in favor of {genders[0] if gap > 0 else genders[1]}")
                
                if abs(gap) > 5:  # Significant pay gap
                    insights.append("  - **Note**: This pay gap exceeds the typical threshold of 5% and may require review.")
        
        # Highest and lowest paid departments
        highest_paid = max(dept_analysis, key=lambda x: x['avg_salary'])
        lowest_paid = min(dept_analysis, key=lambda x: x['avg_salary'])
        
        # Add department compensation comparison
        insights.append("### Department Compensation Comparison")
        insights.append(
            f"- **Highest Paid Department**: {highest_paid['name']} (${highest_paid['avg_salary']:,.2f} average, "
            f"{(highest_paid['avg_salary']/avg_salary - 1)*100:+.1f}% above company average)"
        )
        insights.append(
            f"- **Lowest Paid Department**: {lowest_paid['name']} (${lowest_paid['avg_salary']:,.2f} average, "
            f"{(lowest_paid['avg_salary']/avg_salary - 1)*100:+.1f}% vs company average)"
        )
        
        # Calculate department pay disparity
        dept_pay_disparity = ((highest_paid['avg_salary'] - lowest_paid['avg_salary']) / lowest_paid['avg_salary']) * 100
        insights.append(f"- **Department Pay Disparity**: {dept_pay_disparity:.1f}% between highest and lowest paid departments")
        
        # Add detailed recommendations
        insights.append("\n## Strategic Recommendations")
        
        # Compensation strategy recommendations
        insights.append("### 1. Compensation Strategy")
        insights.append("- **Conduct a comprehensive compensation review** to ensure internal equity and market competitiveness")
        insights.append("- **Standardize salary ranges** across similar roles and departments")
        insights.append("- **Implement pay transparency** guidelines to promote fairness and trust")
        
        # Workforce planning recommendations
        insights.append("\n### 2. Workforce Planning")
        insights.append(f"- **Review staffing levels** in {highest_paid['name']} (highest paid) vs {lowest_paid['name']} (lowest paid)")
        insights.append("- **Develop retention strategies** for critical roles with high salary variance")
        insights.append("- **Create career pathways** to support internal mobility and growth")
        
        # Budget and resource allocation
        insights.append("\n### 3. Budget & Resource Allocation")
        insights.append("- **Allocate budget** based on market data and internal equity metrics")
        insights.append("- **Consider merit-based increases** tied to performance metrics")
        insights.append("- **Invest in upskilling** to improve internal mobility and reduce external hiring costs")
        
        # Pay equity and diversity
        if 'Gender' in df_analysis.columns and len(df_analysis['Gender'].unique()) > 1:
            insights.append("\n### 4. Pay Equity & Diversity")
            insights.append("- **Conduct a pay equity audit** to identify and address any unexplained pay gaps")
            insights.append("- **Implement bias training** for managers involved in compensation decisions")
            insights.append("- **Set diversity and inclusion goals** for leadership positions")
        
        # Key findings and action items
        insights.append("\n## Key Findings & Action Items")
        
        # Top 3 key findings
        insights.append("### Top 3 Key Findings")
        insights.append(f"1. **Wide Salary Range**: The organization has a significant salary range from ${min_salary:,.2f} to ${max_salary:,.2f} "
                      f"({((max_salary - min_salary) / min_salary) * 100:.1f}% range)")
        
        insights.append(f"2. **Department Disparity**: {highest_paid['name']} department earns {dept_pay_disparity:.1f}% more on average than {lowest_paid['name']}")
        
        if 'Gender' in df_analysis.columns and len(df_analysis['Gender'].unique()) > 1:
            gender_gap = abs((gender_pay.iloc[0] / gender_pay.iloc[1] - 1) * 100)
            insights.append(f"3. **Gender Pay Gap**: A {gender_gap:.1f}% pay gap exists between {gender_pay.index[0]} and {gender_pay.index[1]} employees")
        else:
            insights.append(f"3. **Compression Risk**: The ratio of highest to lowest paid employees is {max_salary/min_salary:.1f}:1")
        
        # Immediate action items
        insights.append("\n### Immediate Action Items")
        insights.append("1. **Review Outliers**: Examine employees with salaries outside the IQR range")
        insights.append("2. **Department Equity**: Analyze compensation within each department for consistency")
        insights.append("3. **Market Benchmarking**: Compare current salaries with industry standards")
        
        # Long-term considerations
        insights.append("\n### Long-term Considerations")
        insights.append("- **Performance-based Pay**: Link compensation to measurable performance metrics")
        insights.append("- **Total Rewards Strategy**: Consider non-monetary benefits and recognition programs")
        insights.append("- **Succession Planning**: Identify and develop internal talent for key positions")
        
        # Key Metrics Summary
        insights.append("\n## Key Metrics Summary")
        insights.append(f"- **Total Employees**: {num_employees}")
        insights.append("- **Department Breakdown**: " + ", ".join(f"{k} ({v})" for k, v in dept_counts.items()))
        insights.append(f"- **Average Salary**: ${avg_salary:,.2f}")
        insights.append(f"- **Salary Range**: ${min_salary:,.2f} - ${max_salary:,.2f}")
        
        # Join all insights with newlines
        full_insights = "\n".join(insights)
        
        return {
            'insights': full_insights,
            'status': 'success',
            'message': 'Analysis completed successfully'
        }
        
    except Exception as e:
        # Fallback to a basic report if analysis fails
        error_msg = f"Error generating analysis: {str(e)}"
        st.sidebar.error(error_msg)
        
        # Return a basic report with error information
        return {
            'insights': "# Employee Data Analysis\n\n" \
                      "## Error Generating Full Analysis\n" \
                      f"An error occurred during analysis: {error_msg}\n\n" \
                      "## Basic Metrics\n" \
                      f"- Total Employees: {len(df) if 'df' in locals() else 'N/A'}\n" \
                      f"- Average Salary: ${df['Salary'].mean():,.2f if 'df' in locals() and 'Salary' in df.columns else 'N/A'}",
            'status': 'error',
            'message': error_msg
        }
def export_to_excel(df: pd.DataFrame, analysis_data: Dict, sheet_name: str) -> BytesIO:
    """
    Export the employee data and analysis to an Excel file with multiple sheets.
    
    Args:
        df: DataFrame containing the employee data
        analysis_data: Dictionary containing analysis results
        sheet_name: Name of the sheet being processed
        
    Returns:
        BytesIO: In-memory file object containing the Excel file
    """
    # Create a BytesIO buffer for the Excel file
    excel_buffer = BytesIO()
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        # Write the main data to the first sheet
        df.to_excel(writer, sheet_name='Employee Data', index=False)
        
        # Create a summary sheet with key metrics
        summary_data = []
        
        # Add basic metrics
        summary_data.append(['Metric', 'Value'])
        summary_data.append(['Total Employees', len(df)])
        summary_data.append(['Average Salary', df['Salary'].mean()])
        summary_data.append(['Median Salary', df['Salary'].median()])
        summary_data.append(['Minimum Salary', df['Salary'].min()])
        summary_data.append(['Maximum Salary', df['Salary'].max()])
        
        # Add department summary
        if 'Department' in df.columns:
            dept_summary = df.groupby('Department')['Salary'].agg(['count', 'mean', 'min', 'max', 'median'])
            dept_summary.columns = ['Employee Count', 'Average Salary', 'Min Salary', 'Max Salary', 'Median Salary']
            dept_summary = dept_summary.sort_values('Average Salary', ascending=False)
            
            # Add department data to summary
            summary_data.append(['', ''])
            summary_data.append(['Department Summary', ''])
            summary_data.append(['Department', 'Employee Count', 'Avg Salary', 'Min Salary', 'Max Salary', 'Median Salary'])
            
            for dept, row in dept_summary.iterrows():
                summary_data.append([
                    dept,
                    row['Employee Count'],
                    row['Average Salary'],
                    row['Min Salary'],
                    row['Max Salary'],
                    row['Median Salary']
                ])
        
        # Write summary data to a new sheet
        if len(summary_data) > 1:  # Ensure we have data rows
            # Create a DataFrame with the correct number of columns
            max_cols = max(len(row) for row in summary_data)
            
            # Ensure all rows have the same number of columns
            padded_data = []
            for row in summary_data[1:]:  # Skip header
                padded_row = row + [''] * (max_cols - len(row))
                padded_data.append(padded_row)
            
            # Create DataFrame with padded data
            summary_df = pd.DataFrame(padded_data, columns=summary_data[0] + [''] * (max_cols - len(summary_data[0])))
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
            # Add charts to the summary sheet
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            
            # Add a bar chart for department salaries
            if 'Department' in df.columns and len(dept_summary) > 0:
                chart = workbook.add_chart({'type': 'column'})
                
                # Get the department data range
                dept_data_start = len(summary_data) - len(dept_summary)
                dept_data_end = len(summary_data) - 1
                
                chart.add_series({
                    'name': '=Summary!$C$1',
                    'categories': f'=Summary!$A${dept_data_start + 1}:$A${dept_data_end}',
                    'values': f'=Summary!$C${dept_data_start + 1}:$C${dept_data_end}',
                    'data_labels': {'value': True, 'num_format': '$#,##0'}
                })
                
                chart.set_title({'name': 'Average Salary by Department'})
                chart.set_x_axis({'name': 'Department'})
                chart.set_y_axis({'name': 'Salary ($)', 'major_gridlines': {'visible': True}})
                
                # Insert the chart into the worksheet
                worksheet.insert_chart('H2', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Add a sheet for detailed analysis
        analysis_sheet = workbook.add_worksheet('Detailed Analysis')
        
        # Write the analysis text with formatting
        analysis_text = analysis_data.get('insights', 'No analysis available')
        analysis_lines = analysis_text.split('\n')
        
        # Add formatting for headers and sections
        header_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#D9EAD3'})
        section_format = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#E6E6E6'})
        bold_format = workbook.add_format({'bold': True})
        
        row = 0
        for line in analysis_lines:
            if line.startswith('# '):  # Main header
                analysis_sheet.write(row, 0, line[2:], header_format)
                row += 2
            elif line.startswith('## '):  # Section header
                analysis_sheet.write(row, 0, line[3:], section_format)
                row += 1
            elif line.startswith('### '):  # Subsection
                analysis_sheet.write(row, 1, line[4:], bold_format)
                row += 1
            elif line.startswith('- **'):  # Bold items in lists
                # Extract bold text and regular text
                parts = line.split('**')
                col = 2
                for i, part in enumerate(parts):
                    if i % 2 == 1:  # Bold text
                        analysis_sheet.write(row, col, part, bold_format)
                    else:  # Regular text
                        analysis_sheet.write(row, col, part)
                    col += len(part)  # Adjust column position
                row += 1
            elif line.strip():  # Regular text
                analysis_sheet.write(row, 0, line)
                row += 1
            else:  # Empty line
                row += 1
        
        # Auto-adjust column widths
        for i, width in enumerate([40, 20, 30, 30, 30, 30]):
            analysis_sheet.set_column(i, i, width)
    
    # Reset the buffer position to the beginning
    excel_buffer.seek(0)
    return excel_buffer

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
        
        # Generate analysis with progress update
        progress_bar.progress(10, text="Analyzing employee data...")
        analysis_response = analyze_employee_data(df)
        
        # Handle the analysis response
        if isinstance(analysis_response, dict) and 'insights' in analysis_response:
            ai_insights = analysis_response['insights']
            ai_available = analysis_response.get('status') == 'success'
        else:
            ai_insights = "Analysis could not be generated at this time."
            ai_available = False
            
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
    
    # Generate data button with a unique key
    if st.button("🚀 Generate Employee Data", key="generate_button"):
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
                
                # Create a base filename with timestamp
                timestamp = datetime.now().strftime("%y%m%d_%H%M")
                base_filename = f"EmpData_{timestamp}"
                filename = f"{base_filename}.{file_ext}"
                
                st.success("✅ Data generated successfully!")
                
                # Ensure we have a preview DataFrame
                preview_df = None
                try:
                    if output_format == 'PDF with Summary':
                        preview_df = df.copy()  # Use the generated DataFrame for preview
                        st.subheader("📊 Data Preview (First 5 Rows)")
                        st.dataframe(preview_df.head())
                    else:
                        # For Excel, read the first sheet
                        if hasattr(output_file, 'seek'):
                            output_file.seek(0)  # Ensure we're at the start of the file
                            
                        with pd.ExcelFile(output_file) as xls:
                            preview_df = pd.read_excel(xls, sheet_name=sheet_names[0])
                            st.subheader("📊 Data Preview (First 5 Rows)")
                            st.dataframe(preview_df.head())
                except Exception as e:
                    st.error(f"Error preparing data preview: {str(e)}")
                    st.stop()  # Stop execution if we can't create a preview
                
                # Show statistics
                st.subheader("📈 Data Statistics")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Employees", f"{num_employees:,}")
                with col2:
                    st.metric("Number of Sheets", len(sheet_names) if output_format != 'PDF with Summary' else 1)
                with col3:
                    st.metric("Duplicate IDs", f"{int(num_employees * (duplicate_percentage / 100)):,}")
                
                # Export options section
                st.subheader("📤 Export Options")
                
                # Create columns for the download buttons
                col1, col2 = st.columns(2)
                
                # Excel Export
                with col1:
                    try:
                        if output_format == 'PDF with Summary':
                            # Generate Excel version if in PDF mode
                            with st.spinner("Preparing Excel export..."):
                                analysis_data = analyze_employee_data(df)
                                excel_buffer = export_to_excel(df, analysis_data, sheet_names[0])
                                
                                # Ensure we have a file-like object
                                if hasattr(excel_buffer, 'getvalue'):
                                    excel_data = excel_buffer.getvalue()
                                elif hasattr(excel_buffer, 'read'):
                                    excel_data = excel_buffer.read()
                                else:
                                    raise ValueError("Invalid Excel buffer format")
                                
                                st.download_button(
                                    label="📊 Download Excel Version",
                                    data=excel_data,
                                    file_name=f"employee_analysis_{sheet_names[0]}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="excel_download"
                                )
                        else:
                            # Already in Excel mode, use the generated file
                            if hasattr(output_file, 'getvalue'):
                                file_data = output_file.getvalue()
                            elif hasattr(output_file, 'read'):
                                file_data = output_file.read()
                            else:
                                file_data = output_file
                                
                            st.download_button(
                                label="📊 Download Excel File",
                                data=file_data,
                                file_name=filename,
                                mime=mime_type,
                                use_container_width=True,
                                key=f"excel_direct_download_{timestamp}",
                                on_click=None  # This prevents the app from rerunning
                            )
                    except Exception as e:
                        st.error(f"Error preparing Excel download: {str(e)}")
                        st.exception(e)  # Show full traceback for debugging
                
                # PDF Export
                with col2:
                    try:
                        if output_format != 'PDF with Summary':
                            if preview_df is None:
                                st.warning("No preview data available for PDF generation")
                            else:
                                with st.spinner("Generating PDF report..."):
                                    pdf_buffer = create_pdf_with_summary(preview_df, sheet_names[0])
                                    
                                    # Ensure we have the PDF data in the correct format
                                    if hasattr(pdf_buffer, 'getvalue'):
                                        pdf_data = pdf_buffer.getvalue()
                                    elif hasattr(pdf_buffer, 'read'):
                                        pdf_data = pdf_buffer.read()
                                    else:
                                        pdf_data = pdf_buffer
                                    
                                    st.download_button(
                                        label="📄 Download PDF Report",
                                        data=pdf_data,
                                        file_name=f"employee_analysis_{sheet_names[0]}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True,
                                        key="pdf_export"
                                    )
                        else:
                            # Already in PDF mode, use the generated file
                            if hasattr(output_file, 'getvalue'):
                                file_data = output_file.getvalue()
                            elif hasattr(output_file, 'read'):
                                file_data = output_file.read()
                            else:
                                file_data = output_file
                                
                            st.download_button(
                                label="📄 Download PDF File",
                                data=file_data,
                                file_name=filename,
                                mime=mime_type,
                                use_container_width=True,
                                key=f"pdf_direct_download_{timestamp}",
                                on_click=None  # This prevents the app from rerunning
                            )
                    except Exception as e:
                        st.error(f"Error preparing PDF download: {str(e)}")
                        st.exception(e)  # Show full traceback for debugging
                
                # End of the PDF download section
                
                # Add Detailed Summary Section
                st.subheader("📋 Detailed Summary")
                
                if preview_df is not None:
                    # Department Distribution
                    st.markdown("#### Department Distribution")
                    dept_counts = preview_df['Department'].value_counts()
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.dataframe(dept_counts)
                    with col2:
                        fig, ax = plt.subplots()
                        dept_counts.plot(kind='bar', ax=ax, color='skyblue')
                        plt.xticks(rotation=45, ha='right')
                        plt.tight_layout()
                        st.pyplot(fig)
                    
                    # Salary Statistics
                    st.markdown("#### Salary Statistics")
                    if 'Salary' in preview_df.columns:
                        salary_stats = preview_df['Salary'].describe()
                        st.dataframe(salary_stats)
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown("##### Salary Distribution")
                            fig, ax = plt.subplots()
                            sns.histplot(preview_df['Salary'].dropna(), kde=True, ax=ax)
                            st.pyplot(fig)
                        with col2:
                            st.markdown("##### Salary by Department")
                            if 'Department' in preview_df.columns:
                                fig, ax = plt.subplots(figsize=(8, 4))
                                sns.boxplot(data=preview_df, x='Department', y='Salary', ax=ax)
                                plt.xticks(rotation=45, ha='right')
                                plt.tight_layout()
                                st.pyplot(fig)
                    
                    # Missing Data Analysis
                    st.markdown("#### Missing Data Analysis")
                    missing_data = preview_df.isnull().sum()
                    missing_data = missing_data[missing_data > 0]
                    if len(missing_data) > 0:
                        st.dataframe(missing_data.rename('Missing Values'))
                    else:
                        st.info("No missing data found in any column.")
                
                # Documentation section
                st.markdown("---")
                st.markdown("""
### How to use this tool:
1. **Configure** the data settings in the sidebar
2. Select your preferred **Output Format** (Excel or PDF with Summary)
3. Click the **Generate Employee Data** button
4. **Download** the generated file(s) as needed

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

            except Exception as e:
                st.error(f"An error occurred while generating the data: {str(e)}")
                st.exception(e)  # This will show the full traceback in the app

try:
    # ... (rest of the code remains the same)
    st.markdown("### Note:")
    st.markdown("- When using PDF output, only the first sheet will be included in the summary")
    st.markdown("- For multi-sheet exports, use the Excel format")
except Exception as e:
    st.error(f"An error occurred while generating the data: {str(e)}")
    st.exception(e)  # This will show the full traceback in the app

if __name__ == "__main__":
    main()
