import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd
import os
from datetime import datetime

# SerpAPI key (replace with your API key)
SERPAPI_API_KEY = "528d97b7959e3b477af9a4dd42c520270ce8b942d2d460df9e9ac4fa03cb9dc3"

def search_linkedin_profiles():
    technology = entry_technology.get()
    location = entry_location.get()
    
    if not technology or not location:
        messagebox.showwarning("Input Error", "Please enter both technology and location.")
        return
    
    profiles = []
    
    for page in range(0, 5):  # Limit to 5 pages (50 results total)
        query = f'site:linkedin.com/in "{technology}" "{location}"'
        
        # Build the SerpAPI request URL
        url = f"https://serpapi.com/search.json?q={query.replace(' ', '+')}&location={location}&engine=google&start={page * 10}&api_key={SERPAPI_API_KEY}"
        
        try:
            response = requests.get(url)
            response.raise_for_status()
            
            data = response.json()
            
            if 'organic_results' not in data or len(data['organic_results']) == 0:
                # If no results are found, continue searching on other pages if available
                continue
            
            for result in data['organic_results']:
                profile = {}
                if "linkedin.com/in" in result['link']:
                    profile['Name'] = result['title']
                    profile['LinkedIn URL'] = result['link']
                    profile['Snippet'] = result.get('snippet', 'No description available')
                    
                    # Extract location if available in the snippet
                    location_in_snippet = result.get('rich_snippet', {}).get('top', {}).get('extensions', [])
                    profile['Location'] = ', '.join(location_in_snippet) if location_in_snippet else 'Location not available'
                    
                    # Add any additional details if present
                    if 'rich_snippet' in result and 'top' in result['rich_snippet']:
                        additional_details = result['rich_snippet']['top'].get('detected_extensions', {})
                        if 'experience' in additional_details:
                            profile['Experience'] = additional_details['experience']
                        if 'company' in additional_details:
                            profile['Company'] = additional_details['company']
                        if 'title' in additional_details:
                            profile['Title'] = additional_details['title']
                    
                    profiles.append(profile)
        
        except requests.exceptions.HTTPError as http_err:
            messagebox.showerror("HTTP Error", f"HTTP error occurred: {http_err}")
            continue  # Try the next page if one page fails
        except requests.exceptions.ConnectionError as conn_err:
            messagebox.showerror("Connection Error", f"Error connecting: {conn_err}")
            continue
        except requests.exceptions.Timeout as timeout_err:
            messagebox.showerror("Timeout Error", f"Timeout error occurred: {timeout_err}")
            continue
        except requests.exceptions.RequestException as req_err:
            messagebox.showerror("Request Error", f"Error: {req_err}")
            continue
    
    if profiles:
        folder_path = os.path.join(os.path.expanduser("~"), "Desktop", "Profiles Data")
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        
        # Define the file name
        file_name = f"{technology}_profiles.xlsx"
        file_path = os.path.join(folder_path, file_name)
        
        # If the file already exists, append the current date and time to the new file
        if os.path.exists(file_path):
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"{technology}_profiles_{current_time}.xlsx"
            file_path = os.path.join(folder_path, file_name)
        
        # Save the profiles to the Excel file
        df = pd.DataFrame(profiles)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Profiles saved to {file_path}")
    else:
        messagebox.showinfo("No Results", "No LinkedIn profiles found.")
    
    root.destroy()

# Create the GUI
root = tk.Tk()
root.title("LinkedIn Profile Search")

# Set the GUI to full screen
root.attributes('-fullscreen', True)

# Make the background black
root.configure(bg='black')

# Minimize the black screen (if necessary, it can be done with window minimize)
def minimize_window():
    root.state('iconic')

# Trigger minimize on pressing 'Escape' key
root.bind("<Escape>", lambda e: minimize_window())

# Create the input fields
tk.Label(root, text="Technology:", bg='black', fg='white').grid(row=0, column=0, padx=10, pady=10)
entry_technology = tk.Entry(root)
entry_technology.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Location:", bg='black', fg='white').grid(row=1, column=0, padx=10, pady=10)
entry_location = tk.Entry(root)
entry_location.grid(row=1, column=1, padx=10, pady=10)

# Search button
tk.Button(root, text="Search", command=search_linkedin_profiles).grid(row=2, columnspan=2, pady=20)

root.mainloop()
