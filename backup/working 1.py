import tkinter as tk
from tkinter import filedialog, messagebox
import requests
import random
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
import datetime
import os

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
                d.append([keyword, now, rank, page, site_name_str])
                # Stop checking for lower results once a match is found
                break
        else:
            # If no site_name was found in the URL, continue to the next URL
            continue
        # Break out of the outer loop once a match is found
        break
    else:
        # If no site_name was found in any URL, add a row with rank 100 and page 11
        now = datetime.date.today().strftime("%d-%m-%Y")
        for site_name in site_names:
            site_name_str = str(site_name)  # Convert site_name to string
            d.append([keyword, now, 100, 100, site_name_str])

    df = pd.DataFrame(d, columns=['Keyword', 'Date', 'Rank', 'Page', 'Site'])
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
        self.geometry("400x400")

        # Create a label to display the selected company
        self.company_label = tk.Label(self, text="Select a company:")
        self.company_label.pack()

        # Create a dropdown menu to select the company
        self.company_var = tk.StringVar()
        self.company_dropdown = tk.OptionMenu(self, self.company_var, *self.get_company_names(), command=self.update_keywords)
        self.company_dropdown.pack()

        # Create a listbox to display the keywords
        self.keyword_listbox = tk.Listbox(self, width=40, height=10)
        self.keyword_listbox.pack()

        # Create a button to start the search
        self.search_button = tk.Button(self, text="Search", command=self.search_keywords)
        self.search_button.pack()

        # Create a text area to display the results
        self.results_text = tk.Text(self, wrap=tk.WORD, width=50, height=10)
        self.results_text.pack()

        # Load the keywords for the first company
        self.load_keywords(self.company_var.get())

    def get_company_names(self):
        company_folder = 'company'
        company_names = []
        for file_name in os.listdir(company_folder):
            if file_name.endswith('.xlsx'):
                company_names.append(os.path.splitext(file_name)[0])

        print(company_names)
        return company_names

    def update_keywords(self, event=None):
        company = self.company_var.get()
        self.load_keywords(company)

    def load_keywords(self, company):
        # Clear the listbox
        self.keyword_listbox.delete(0, tk.END)

        # Check if the company is not an empty string
        if company:
            # Read the keywords from the Excel file
            keywords_file = f'company/{company}.xlsx'
            if os.path.exists(keywords_file):
                company_df = pd.read_excel(keywords_file)
                site_names = company_df['Name'].tolist()
                keywords = company_df['Keyword'].tolist()

                # Insert the keywords into the listbox
                for keyword in keywords:
                    self.keyword_listbox.insert(tk.END, keyword)
            else:
                messagebox.showerror("Error", f"File not found: {keywords_file}")
        else:
            messagebox.showerror("Error", "Please select a company first.")

    def search_keywords(self):
        # Get the selected company
        company = self.company_var.get()

        # Read the site names and keywords from the Excel file
        keywords_file = f'company/{company}.xlsx'
        company_df = pd.read_excel(keywords_file)
        site_names = company_df['Name'].tolist()
        keywords = company_df['Keyword'].tolist()

        # Create a single DataFrame to store all results
        all_results = pd.DataFrame()

        # Clear the results text area
        self.results_text.delete("1.0", tk.END)

        # Flag to track if a 429 error has occurred
        rate_limit_hit = False

        # Search for each keyword and append the results to the DataFrame
        for keyword in keywords:
            if not rate_limit_hit:
                desktop = get_data(keyword, site_names)
                if not desktop.empty:
                    if desktop.columns.tolist() == ['status']:
                        # Check if the DataFrame contains the 'status' column (indicating an error)
                        error_message = desktop.iloc[0]['status']
                        if 'status code: 429' in error_message:
                            rate_limit_hit = True
                            print(f"Rate limit hit, status code 429 for keyword '{keyword}'. Skipping remaining keywords.")
                            break
                    else:
                        all_results = pd.concat([all_results, desktop], ignore_index=True)
                        # Display the results in the text area
                        self.results_text.insert(tk.END, str(desktop) + "\n\n")
                else:
                    # Handle the case where an empty DataFrame is returned (429 error)
                    rate_limit_hit = True
                    print(f"Rate limit hit, status code 429 for keyword '{keyword}'. Skipping remaining keywords.")
                    break

        if not rate_limit_hit:
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
        else:
            messagebox.showwarning("Rate Limit Hit", "Rate limit reached, skipping remaining keywords.")

if __name__ == "__main__":
    app = App()
    app.mainloop()