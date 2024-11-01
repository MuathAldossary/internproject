import pandas as pd
from datetime import datetime

# Define the file where data will be saved
excel_file = "internship_applications.xlsx"

# Function to add a new application
def add_application(company_name, duo_date):
    today = datetime.today().strftime('%Y-%m-%d')
    status = "Waiting"
    
    # Calculate days left until the due date
    due_date = datetime.strptime(duo_date, '%Y-%m-%d')
    days_left = (due_date - datetime.today()).days

    new_entry = pd.DataFrame([{
        "Company Name": company_name,
        "Application Date": today,
        "Status": status,
        "Assessment Due Date": duo_date,
        "Days Left": days_left
    }])
    
    # Append to existing data or create a new DataFrame if file doesn't exist
    try:
        df = pd.read_excel(excel_file)
        df = pd.concat([df, new_entry], ignore_index=True)
    except FileNotFoundError:
        df = new_entry
    
    # Save to Excel
    df.to_excel(excel_file, index=False)
    print(f"Added application for {company_name} with assessment due on {duo_date}.")

# Function to update the status
def update_status(company_name, new_status):
    try:
        df = pd.read_excel(excel_file)
        # Check if the company name exists in the DataFrame
        if company_name in df["Company Name"].values:
            # Update the status where the company name matches
            df.loc[df["Company Name"] == company_name, "Status"] = new_status
            df.to_excel(excel_file, index=False)
            print(f"Updated status for {company_name} to {new_status}.")
        else:
            print(f"Company name '{company_name}' is not listed.")
    except FileNotFoundError:
        print("No applications found. Add an application first.")


# Function to display all applications
def view_applications():
    try:
        df = pd.read_excel(excel_file)
        
        # Update days left based on the current date
        today = datetime.today()
        
        # Ensure we only process valid date strings
        df["Days Left"] = df["Assessment Due Date"].apply(
            lambda x: (datetime.strptime(x, '%Y-%m-%d') - today).days if isinstance(x, str) else None
        )
        
        # Save updated days left back to Excel
        df.to_excel(excel_file, index=False)
        print(df)
    except FileNotFoundError:
        print("No applications found. Add an application first.")


# Main menu to interact with the program
def main_menu():
    while True:
        print("\n--- Internship Application Tracker ---")
        print("1. Add new application")
        print("2. View all applications")
        print("3. Update application status")
        print("4. Exit")
        
        choice = input("Choose an option: ")
        
        if choice == '1':
            company_name = input("Enter the company name: ")
            duo_date = input("Enter assessment due date (YYYY-MM-DD): ")
            add_application(company_name, duo_date)
        elif choice == '2':
            view_applications()
        elif choice == '3':
            company_name = input("Enter the company name to update: ")
            new_status = input("Enter the new status (e.g., Interview Scheduled, Rejected): ")
            update_status(company_name, new_status)
        elif choice == '4':
            print("Exiting the program.")
            break
        else:
            print("Invalid choice. Please try again.")

# Run the main menu
main_menu()
