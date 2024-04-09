import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import tkinter.ttk as ttk
import requests
import random
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
import datetime
import os
import time
import threading

# Desktop user agent strings
desktop_agent = [
    # Updated Chrome 110 on Windows
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
    
    # Updated Chrome 110 on MacOS
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
    
    # Updated Firefox 105 on MacOS
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:105.0) Gecko/20100101 Firefox/105.0',

    # Updated Safari 15 on MacOS
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:15.0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Safari/605.1.15',
]

def clean_url(url):
    # Find the start of 'https://'
    start = url.find('https://')
    if start == -1:
        return None  # 'https://' not found in the URL

    # Find the end position, which is the start of '&ved'
    end = url.find('&ved', start)
    if end == -1:
        # If '&ved' is not found, return the URL from 'https://' onwards
        return url[start:]
    else:
        # Return the URL from 'https://' up to '&ved'
        return url[start:end]
    
def rank_check(site_names, serp_df, keyword):
    counter = 0
    d = []
    for url in serp_df['URLs']:
        counter += 1
        url_str = str(url)  # Convert to string
        for site_name in site_names:
            site_name_str = str(site_name)  # Convert site_name to string
            if url_str.find(site_name_str) != -1:  # Check if site_name_str is a substring of url_str
                rank = counter
                now = datetime.date.today().strftime("%d-%m-%Y")
                if rank <= 3:
                    page = 0  # Set page to 0 for top 3 rankings
                else:
                    page = (rank - 1) // 10 + 1  # Calculate the page number for ranks beyond 3
                d.append([keyword, now, rank, page])
                # Stop checking for lower results once a match is found
                break
        else:
            # If no site_name was found in the URL, continue to the next URL
            continue
        # Break out of the outer loop once a match is found
        break
    else:
        # If no site_name was found in any URL, add a single row with rank 100 and page 100
        now = datetime.date.today().strftime("%d-%m-%Y")
        d.append([keyword, now, 100, 100])

    df = pd.DataFrame(d, columns=['Keyword', 'Date', 'Rank', 'Page'])
    return df


def get_data(keyword, site_names):
    # Google Search URL
    search_number = 100
    google_url = f'https://www.google.com/search?num={search_number}&q={keyword}&gl=hk&hl=zh-HK'

    useragent = random.choice(desktop_agent)
    headers = {'User-Agent': useragent}

    # Make the request
    response = requests.get(google_url + keyword, headers=headers)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the content
        soup = BeautifulSoup(response.text, 'html.parser')

        urls = soup.find_all('div', class_="yuRUbf")

        data = []
        for div in urls:
            soup = BeautifulSoup(str(div), 'html.parser')

            # Extracting the URL
            url_anchor = soup.find('a')
            if url_anchor:
                url = url_anchor.get('href', "No URL")
            else:
                url = "No URL"

            url = clean_url(url)

            data.append(url)

        serp_df = pd.DataFrame(data, columns=['URLs'])
        serp_df = serp_df.dropna(subset=['URLs'])

        # Convert site_names to strings
        site_names_str = [str(site_name) for site_name in site_names]

        results = rank_check(site_names_str, serp_df, keyword)

        print(f"Ranking results for {', '.join(site_names_str)} with keyword '{keyword}':")
        print(results)

        return results

    elif response.status_code == 429:
        # Handle rate limiting
        print(f"Rate limit hit, status code 429 for keyword '{keyword}'. Skipping this keyword.")
        return pd.DataFrame()  # Return an empty DataFrame instead of an error message
    else:
        # Handle other status codes
        error_message = f'Failed to retrieve data, status code: {response.status_code}'
        print(error_message)
        return pd.DataFrame({'status': [error_message]})  # Return a DataFrame with the error message


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Keyword Ranking Checker")
        self.geometry("600x900")
        self.configure_gui()
        self.create_widgets()
        

    def configure_gui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("H1.TLabel", font=("Helvetica", 16, "bold"), padding=10)
        style.configure("default.TFrame", background="#ffffff")
        style.configure("default.TButton", padding=6)

    def create_widgets(self):
        main_frame = ttk.Frame(self, style="default.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        left_frame = ttk.Frame(main_frame, style="default.TFrame")
        left_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 10))

        right_frame = ttk.Frame(main_frame, style="default.TFrame")
        right_frame.grid(row=0, column=1, sticky=tk.NSEW)

        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Create a label to display the selected company
        self.company_label = ttk.Label(left_frame, text="Select a company:", style="H2.TLabel")
        self.company_label.pack(anchor=tk.W)

        # Create a dropdown menu to select the company
        self.company_var = tk.StringVar()
        self.company_dropdown = ttk.Combobox(left_frame, textvariable=self.company_var, values=self.get_company_names(), state="readonly")
        self.company_dropdown.bind("<<ComboboxSelected>>", self.update_keywords)
        self.company_dropdown.pack(fill=tk.X, pady=(0, 10))

        # Pre-select the first item in the combobox
        if self.company_dropdown['values']:
            self.company_dropdown.current(0)

        # Create a listbox to display the keywords
        self.keyword_listbox = tk.Listbox(left_frame, width=40, height=30)
        self.keyword_listbox.pack(fill=tk.BOTH, expand=True)

        keyword_button_frame = ttk.Frame(left_frame, style="default.TFrame")
        keyword_button_frame.pack(fill=tk.X, pady=(10, 0))

        # Create buttons to add/remove keywords
        self.add_keyword_button = ttk.Button(keyword_button_frame, text="Add Keyword", command=self.add_keyword, style="default.TButton")
        self.add_keyword_button.pack(side=tk.LEFT, padx=(0, 5), anchor="center")

        self.remove_keyword_button = ttk.Button(keyword_button_frame, text="Remove Keyword", command=self.remove_keyword, style="default.TButton")
        self.remove_keyword_button.pack(side=tk.LEFT, anchor="center")

        # Create a listbox to display the target site URLs
        self.site_listbox = tk.Listbox(left_frame, width=40, height=5)
        self.site_listbox.pack(fill=tk.BOTH, expand=True, pady=(20, 0))

        # Create a frame to hold the add/remove site buttons
        site_button_frame = ttk.Frame(left_frame, style="default.TFrame")
        site_button_frame.pack(fill=tk.X, pady=(10, 0))

        # Create buttons to add/remove target sites
        self.add_site_button = ttk.Button(site_button_frame, text="Add Site", command=self.add_site, style="default.TButton")
        self.add_site_button.pack(side=tk.LEFT, padx=(0, 5), anchor="center")

        self.remove_site_button = ttk.Button(site_button_frame, text="Remove Site", command=self.remove_site, style="default.TButton")
        self.remove_site_button.pack(side=tk.LEFT, anchor="center")

        # Create a button to start the search
        self.search_button = ttk.Button(left_frame, text="Search", command=self.search_keywords, style="default.TButton")
        self.search_button.pack(anchor=tk.CENTER, expand=True)

        # Create three text areas to display the keyword, rank, and page
        self.keyword_text = tk.Text(right_frame, wrap=tk.WORD, width=20, height=20)
        self.keyword_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.rank_text = tk.Text(right_frame, wrap=tk.WORD, width=5, height=20)
        self.rank_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.page_text = tk.Text(right_frame, wrap=tk.WORD, width=5, height=20)
        self.page_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Load the keywords and target site URLs for the first company
        self.load_keywords(self.company_var.get())
        self.load_urls(self.company_var.get())

    def add_keyword(self):
        keyword = simpledialog.askstring("Add Keyword", "Enter a new keyword:")
        if keyword:
            if keyword == "20191104":
                messagebox.showinfo("Message", "Hugo love you so so much!")
            else:
                company = self.company_var.get()
                keywords_file = f'keywords/{company}.xlsx'
                if os.path.exists(keywords_file):
                    company_df = pd.read_excel(keywords_file)
                    new_row = pd.DataFrame({'Keyword': [keyword]})  # Create a new DataFrame with the keyword
                    company_df = pd.concat([company_df, new_row], ignore_index=True)  # Concatenate the new row to the existing DataFrame
                    company_df.to_excel(keywords_file, index=False)
                    self.load_keywords(company)
                    self.load_urls(company)  # Update the site listbox as well
                else:
                    messagebox.showerror("Error", f"File not found: {keywords_file}")

    def remove_keyword(self):
        selected_index = self.keyword_listbox.curselection()
        if selected_index:
            keyword = self.keyword_listbox.get(selected_index)
            confirm = messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove the keyword '{keyword}'?")
            if confirm:
                company = self.company_var.get()
                keywords_file = f'keywords/{company}.xlsx'
                if os.path.exists(keywords_file):
                    company_df = pd.read_excel(keywords_file)
                    company_df = company_df[company_df['Keyword'] != keyword]
                    company_df.to_excel(keywords_file, index=False)
                    self.load_keywords(company)
                    self.load_urls(company)  # Update the site listbox as well
                else:
                    messagebox.showerror("Error", f"File not found: {keywords_file}")

    def add_site(self):
        site = simpledialog.askstring("Add Site", "Enter a new target site:")
        if site:
            company = self.company_var.get()
            urls_file = f'URLs/{company}.xlsx'
            if os.path.exists(urls_file):
                company_df = pd.read_excel(urls_file)
                new_row = pd.DataFrame({'Name': [site]})  # Create a new DataFrame with the site name
                company_df = pd.concat([company_df, new_row], ignore_index=True)  # Concatenate the new row to the existing DataFrame
                company_df.to_excel(urls_file, index=False)
                self.load_keywords(company)
                self.load_urls(company)  # Update the site listbox as well
            else:
                messagebox.showerror("Error", f"File not found: {urls_file}")

    def remove_site(self):
        selected_index = self.site_listbox.curselection()
        if selected_index:
            site = self.site_listbox.get(selected_index)
            confirm = messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove the site '{site}'?")
            if confirm:
                company = self.company_var.get()
                urls_file = f'URLs/{company}.xlsx'
                if os.path.exists(urls_file):
                    company_df = pd.read_excel(urls_file)
                    company_df = company_df[company_df['Name'] != site]
                    company_df.to_excel(urls_file, index=False)
                    self.load_keywords(company)
                    self.load_urls(company)  # Update the site listbox as well
                else:
                    messagebox.showerror("Error", f"File not found: {urls_file}")
                    
    def update_keywords(self, event=None):
        company = self.company_var.get()
        self.load_urls(company)  # Load the target site URLs first
        self.load_keywords(company)  # Then load the keywords

    def load_urls(self, company):
        # Clear the site listbox
        self.site_listbox.delete(0, tk.END)

        # Check if the company is not an empty string
        if company:
            # Read the site names from the Excel file
            urls_file = f'URLs/{company}.xlsx'
            if os.path.exists(urls_file):
                company_df = pd.read_excel(urls_file)
                site_names = company_df['Name'].tolist()
                
                # Filter out NaN values from site_names
                site_names = [site_name for site_name in site_names if pd.notnull(site_name)]
                
                # Insert the site names into the site listbox
                for site_name in site_names:
                    self.site_listbox.insert(tk.END, site_name)
            else:
                messagebox.showerror("Error", f"File not found: {urls_file}")
        else:
            messagebox.showerror("Error", "Please select a company first.")
            
    def get_company_names(self):
        keywords_folder = 'keywords'
        company_names = []
        for file_name in os.listdir(keywords_folder):
            if file_name.endswith('.xlsx'):
                company_names.append(os.path.splitext(file_name)[0])

        print(company_names)
        return company_names

    def load_keywords(self, company):
        # Clear the keyword listbox
        self.keyword_listbox.delete(0, tk.END)

        # Check if the company is not an empty string
        if company:
            # Read the keywords from the Excel file
            keywords_file = f'keywords/{company}.xlsx'
            if os.path.exists(keywords_file):
                company_df = pd.read_excel(keywords_file)
                keywords = company_df['Keyword'].tolist()
                
                # Insert the keywords into the keyword listbox
                for keyword in keywords:
                    self.keyword_listbox.insert(tk.END, keyword)
            else:
                messagebox.showerror("Error", f"File not found: {keywords_file}")
        else:
            messagebox.showerror("Error", "Please select a company first.")

    def search_keywords(self):
        # Disable the search button to prevent multiple clicks
        self.search_button.config(state=tk.DISABLED)

        # Start the search operation in a separate thread
        search_thread = threading.Thread(target=self.search_keywords_thread)
        search_thread.start()

    def search_keywords_thread(self):
        # Get the selected company
        company = self.company_var.get()

        # Read the site names and keywords from the Excel file
        keywords_file = f'keywords/{company}.xlsx'
        company_df = pd.read_excel(keywords_file)
        keywords = company_df['Keyword'].tolist()

        urls_file = f'URLs/{company}.xlsx'
        company_df = pd.read_excel(urls_file)
        site_names = company_df['Name'].tolist()

        # Create a single DataFrame to store all results
        all_results = pd.DataFrame()

        # Clear the text areas
        self.keyword_text.delete("1.0", tk.END)
        self.rank_text.delete("1.0", tk.END)
        self.page_text.delete("1.0", tk.END)

        # Loop through keywords and search
        for keyword in keywords:
            retry = True
            while retry:
                desktop = get_data(keyword, site_names)
                if not desktop.empty:
                    if desktop.columns.tolist() == ['status']:
                        # Check if the DataFrame contains the 'status' column (indicating an error)
                        error_message = desktop.iloc[0]['status']
                        if 'status code: 429' in error_message:
                            # Rate limit hit, prompt user to change VPN
                            change_vpn = messagebox.askyesno("Rate Limit Hit", f"Rate limit hit for keyword '{keyword}'. Change VPN and try again?")
                            if change_vpn:
                                # User chose to change VPN, continue retrying the same keyword
                                continue
                            else:
                                # User chose not to change VPN, break out of both loops and skip remaining keywords
                                retry = False
                                break
                        else:
                            # Other error, display message and stop retrying
                            messagebox.showerror("Error", error_message)
                            retry = False
                    else:
                        all_results = pd.concat([all_results, desktop], ignore_index=True)
                        
                        # Insert the keyword, rank, and page into their respective text areas
                        self.keyword_text.insert(tk.END, desktop['Keyword'].iloc[0] + "\n")
                        self.rank_text.insert(tk.END, str(desktop['Rank'].iloc[0]) + "\n")
                        self.page_text.insert(tk.END, str(desktop['Page'].iloc[0]) + "\n")

                        # Auto-scroll to the bottom of each text area
                        self.keyword_text.see(tk.END)
                        self.rank_text.see(tk.END)
                        self.page_text.see(tk.END)

                        retry = False
                else:
                    # Handle the case where an empty DataFrame is returned (429 error)
                    change_vpn = messagebox.askyesno("Rate Limit Hit", f"Rate limit hit for keyword '{keyword}'. Change VPN and try again?")
                    if change_vpn:
                        # User chose to change VPN, continue retrying the same keyword
                        continue
                    else:
                        # User chose not to change VPN, break out of both loops and skip remaining keywords
                        retry = False
                        break
            else:
                # Continue to the next keyword if the inner loop completed normally
                continue
            # Break out of the outer loop if the inner loop was broken out of
            break

        # Remove duplicates and keep the first occurrence
        all_results = all_results.drop_duplicates(['Keyword'], keep='first')

        # Create a new Excel file for the results
        output_folder = 'results'
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        output_file = os.path.join(output_folder, f'{company}_rankings.xlsx')
        writer = pd.ExcelWriter(output_file, engine='openpyxl')

        # Write the DataFrame to the Excel file
        all_results.to_excel(writer, sheet_name='Rankings', index=False)

        # Save the Excel file
        writer.book.save(output_file)
        messagebox.showinfo("Success", f"Results saved to {output_file}")

        # Enable the search button after the search is complete
        self.search_button.config(state=tk.NORMAL)

    def save_results(self, all_results, company):
        # Remove duplicates and keep the first occurrence
        all_results = all_results.drop_duplicates(['Keyword'], keep='first')

        # Create a new Excel file for the results
        output_folder = 'results'
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        output_file = os.path.join(output_folder, f'{company}_rankings.xlsx')
        writer = pd.ExcelWriter(output_file, engine='openpyxl')

        # Write the DataFrame to the Excel file
        all_results.to_excel(writer, sheet_name='Rankings', index=False)

        # Save the Excel file
        writer.book.save(output_file)
        messagebox.showinfo("Success", f"Results saved to {output_file}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
