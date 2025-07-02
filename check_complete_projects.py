import pandas as pd
from datetime import datetime, timedelta
import os

def load_data(project_path: str, allocation_path: str):
    """Load project and allocation Excel files into DataFrames."""
    project_df = pd.read_excel(project_path)
    allocation_df = pd.read_excel(allocation_path)
    return project_df, allocation_df

def preprocess_dates(project_df, allocation_df):
    """Convert date columns to datetime."""
    project_df['created_date'] = pd.to_datetime(project_df['created_date'])
    allocation_df['allocation_date'] = pd.to_datetime(allocation_df['allocation_date'])
    return project_df, allocation_df

def find_completed_projects(project_df, allocation_df, days=30):
    """Return projects with no allocations in the last 'days' days."""
    today = pd.to_datetime(datetime.today().date())
    print(today)
    last_n_days = today - timedelta(days=days)
    
    recent_allocations = allocation_df[allocation_df['allocation_date'] >= last_n_days]
    active_workids = set(recent_allocations['workid'].unique())
    
    completed_projects = project_df[~project_df['workid'].isin(active_workids)].copy()
    completed_projects['status'] = 'Completed'
    
    return completed_projects

def save_completed_projects(completed_df, output_path='completed_projects.xlsx'):
    """Save the completed projects DataFrame to an Excel file."""
    completed_df.to_excel(output_path, index=False)
    print(f"Completed projects report saved to {output_path}")

def main():
    # Define file paths (update if needed)
    project_file = 'project_info.xlsx'
    allocation_file = 'allocation.xlsx'
    output_file = 'completed_projects.xlsx'

    # Check if files exist
    if not os.path.exists(project_file) or not os.path.exists(allocation_file):
        print("Required input files not found. Please ensure Excel files are in the current directory.")
        return

    # Load and process data
    project_df, allocation_df = load_data(project_file, allocation_file)
    project_df, allocation_df = preprocess_dates(project_df, allocation_df)
    completed_df = find_completed_projects(project_df, allocation_df)
    
    # Save result
    save_completed_projects(completed_df, output_file)

if __name__ == "__main__":
    main()




