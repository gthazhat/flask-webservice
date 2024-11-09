import tkinter as tk
from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time

# Function to browse and select the Excel file
def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select CP360 Performance Issues File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

# Function to run the analysis
def run_analysis():
    # Access the global variables
    global file_entry
    global sheet_name_entry

    file_path = file_entry.get()
    sheet_name = sheet_name_entry.get()

    if not file_path:
        print("Please select a file before running the analysis.")
        return

    try:
        # Read the Excel file
        data = pd.read_excel(file_path, sheet_name=sheet_name)

        # Set up Chrome options
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")

        # Set up the Chrome driver service
        service = Service(executable_path=r'C:\chromedriver\chromedriver.exe')  # Update the path to chromedriver
        driver = webdriver.Chrome(service=service, options=options)

        # Authenticate once at the start
        url = data.loc[0, 'Lumberjack Link']
        driver.get(url)

        # Wait for 2 minutes (120 seconds) for the user to authenticate
        print("Waiting for authentication... Please complete the Yubikey authentication.")
        time.sleep(120)  # Wait for 2 minutes

        # Process all URLs in the "Lumberjack Link" column
        for index in range(len(data)):
            url = data.loc[index, 'Lumberjack Link']
            driver.get(url)
            time.sleep(5)  # Wait for the page to load

            # Get page content and parse with BeautifulSoup
            page_content = driver.page_source
            soup = BeautifulSoup(page_content, 'html.parser')
            page_text = soup.get_text()

            # Identify the starting indicators for SQL and Logical SQL
            sql_query_start = "-------------------- Sending query to database named Oracle_Data_Warehouse"
            logical_sql_start = "-------------------- SQL Request, logical request hash:"
            queries = []
            logical_queries = []

            # Extract SQL queries
            start_pos = 0
            while True:
                start_pos = page_text.find(sql_query_start, start_pos)
                if start_pos == -1:
                    break
                start_pos += len(sql_query_start)
                end_pos = page_text.find("--------------------", start_pos)
                if end_pos == -1:
                    end_pos = len(page_text)
                sql_query = page_text[start_pos:end_pos].strip()
                if "WITH" in sql_query:
                    with_index = sql_query.find("WITH")
                    if with_index != -1:
                        cleaned_sql_query = sql_query[with_index:].strip()
                        queries.append(cleaned_sql_query)
                start_pos = end_pos

            # Select and save the longest SQL query
            if queries:
                longest_sql_query = max(queries, key=len)
                data.loc[index, 'PSQL'] = longest_sql_query
            else:
                data.loc[index, 'PSQL'] = None

            # Extract Logical SQL
            start_pos = 0
            while True:
                start_pos = page_text.find(logical_sql_start, start_pos)
                if start_pos == -1:
                    break
                start_pos += len(logical_sql_start)
                end_pos = page_text.find("--------------------", start_pos)
                if end_pos == -1:
                    end_pos = len(page_text)
                logical_sql_query = page_text[start_pos:end_pos].strip()
                logical_queries.append(logical_sql_query)
                start_pos = end_pos

            # Select and save the longest Logical SQL query
            if logical_queries:
                longest_logical_sql_query = max(logical_queries, key=len)
                data.loc[index, 'LSQL'] = longest_logical_sql_query
            else:
                data.loc[index, 'LSQL'] = None

        # Save the updated DataFrame back to the Excel file
        data.to_excel(file_path, sheet_name=sheet_name, index=False)
        print("Analysis completed and data saved.")

        # Close the browser
        driver.quit()

        # Close the main tkinter window (moved here)
        root.destroy()

    except Exception as e:
        print(f"Error: {e}")
        # Optionally keep the window open if an error occurs
        input("Press Enter to exit...")
        root.destroy()

# Create the main tkinter window
root = tk.Tk()
root.title("CP360 FDI Performance Issues Automation Tool")

# Set window size and make it non-resizable
root.geometry("500x300")
root.resizable(False, False)

# Header label
header_label = tk.Label(root, text="CP360 Performance Issues Automation", font=("Helvetica", 16, "bold"), fg="red")
header_label.pack(pady=10)

# File selection frame
file_frame = tk.Frame(root)
file_frame.pack(pady=5)

file_entry = tk.Entry(file_frame, width=40)
file_entry.grid(row=0, column=0, padx=5, pady=5)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file, bg="green", fg="white")
browse_button.grid(row=0, column=1, padx=5, pady=5)

file_label = tk.Label(file_frame, text="Select CP360 Performance Issues File:")
file_label.grid(row=1, column=0, columnspan=2, pady=5)

# Sheet name frame
sheet_name_frame = tk.Frame(root)
sheet_name_frame.pack(pady=5)

sheet_name_label = tk.Label(sheet_name_frame, text="Sheet Name:")
sheet_name_label.grid(row=0, column=0, padx=5, pady=5)

sheet_name_entry = tk.Entry(sheet_name_frame, width=40)
sheet_name_entry.insert(0, "CP360 Performance Outliers Anal")  # Default sheet name
sheet_name_entry.grid(row=0, column=1, padx=5, pady=5)

# Run Analysis button
run_button = tk.Button(root, text="Run Analysis", command=run_analysis, bg="blue", fg="white",
                       font=("Helvetica", 10, "bold"))
run_button.pack(pady=20)

# Footer label
footer_label = tk.Label(root, text="© 2024 Oracle Corporation Analytics Team", font=("Helvetica", 8))
footer_label.pack(side="bottom", pady=10)

# Run the tkinter main loop
root.mainloop()
