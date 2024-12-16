# import all the neccessary library and packages which are need for this code
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from dotenv import load_dotenv
import os
from datetime import datetime, date,timedelta
# Simulate today's date for testing
today = datetime.now()

if today.day == 15:
    # For the 15th, look back to the 30th of the previous month
    first_day_of_month = today.replace(day=1)
    previous_date = first_day_of_month - timedelta(days=1)  # Last day of the previous month
    previous_date = previous_date.replace(day=min(30, previous_date.day))  # Set to 30th if possible
elif today.day == 30 or (today.day >= 28 and today.month == 2):  # Special case for February
    # For the 30th or late February, look back to the 15th of the same month
    previous_date = today.replace(day=15)
else:
    raise ValueError("This script should only run on the 15th or 30th (or 28th in February).")

# Format the date for the directory path
previous_date_str = previous_date.strftime("%Y-%m-%d")
print(previous_date_str,"date previous format")

try:
    # get data from .env file
    load_dotenv()
    username = os.getenv('USERNAME_11')
    password = os.getenv('PASSWORD')
    link = os.getenv('LINK')
    email_user = os.getenv('EMAIL_USER')  # Sender email
    email_password = os.getenv('EMAIL_PASSWORD')  # Sender email password
    recipient_email = os.getenv('RECIPIENT_EMAIL')  # Recipient email
    # Automatically install the ChromeDriver using ChromeDriverManager
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    #Open the login page and link
    driver.get(link)

    def window_size():
        try:
            #Maximize the browser window
            driver.maximize_window()
            driver.maximize_window()

            #Zoom out the window (e.g., 25% zoom)
            zoom_out_percentage = 25  # Adjust zoom percentage (25 means zoom out to 25%)
            driver.execute_script(f"document.body.style.zoom='{zoom_out_percentage}%'")
        except Exception as e:
            print(f"Error Adjusting window size :{e}")

    def Log_in(username,password):
        try:
            window_size()
            # Step 4: Locate the username and password fields and enter credentials
            username_field = driver.find_element(By.ID, "input-13")  # Update if necessary
            password_field = driver.find_element(By.ID, "input-16")  # Update if necessary

            # Step 5: Enter your username and password
            username_field.send_keys(username)
            password_field.send_keys(password)

            # Step 6: Locate and click the login button
            login_button = driver.find_element(By.CSS_SELECTOR, "button span")  # Update if necessary
            login_button.click()

            # Step 7: Wait for the next page to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "__layout"))) 
            print("Logged in successfull")
        except Exception as e:
            print(f"Login failed: {e}")
            driver.quit()
            raise
    Log_in(username,password)
    time.sleep(20)

    def Get_excel(Newlink):
        try:
            driver.get(Newlink)
            window_size()
            time.sleep(20)
            table_1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            headers_1 = [th.text for th in table_1.find_elements(By.TAG_NAME, "th")]
            rows_1 = table_1.find_elements(By.TAG_NAME, "tr")
            table_1_data = []

            for row in rows_1:
                cells = [cell.text for cell in row.find_elements(By.TAG_NAME, "td")]
                if cells:
                    table_1_data.append(cells)

            # Click to load the second table
            try:
                element = driver.find_element(By.CSS_SELECTOR, "#period-select button:nth-of-type(2)")
                element.click()
                is_present = True

                # Wait for the second table and scrape data
                time.sleep(20)
                table_2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
                time.sleep(20)

                headers_2 = [th.text for th in table_2.find_elements(By.TAG_NAME, "th")]
                rows_2 = table_2.find_elements(By.TAG_NAME, "tr")
                table_2_data = []
                for row in rows_2:
                    cells = [cell.text for cell in row.find_elements(By.TAG_NAME, "td")]
                    if cells:
                        table_2_data.append(cells)
            except Exception as e:
                print("Second table not found or could not be loaded.")
                table_2_data = []  # Empty data if second table not found
                headers_2 = []
                is_present = False

            # Retrieve company name and date
            div = driver.find_element(By.CLASS_NAME, "row")
            spans = div.find_elements(By.TAG_NAME, "span")
            Company_name = spans[0]
            Date_Span = spans[12]

            # Create DataFrames from both tables
            df1 = pd.DataFrame(table_1_data, columns=headers_1)
            if is_present:
                df2 = pd.DataFrame(table_2_data, columns=headers_2)

            # Create Excel files
            save_dir = "E:/PYTHON/STOCK_MARKET/ANALYST_ESTIMATES/" f"{date.today()}"
            os.makedirs(save_dir, exist_ok=True)
                # Save as .xlsx using pandas
            excel_file_path_xlsx = os.path.join(save_dir, f"{Company_name.text}.xlsx")
            with pd.ExcelWriter(excel_file_path_xlsx, engine='openpyxl') as writer:
                df1.to_excel(writer, sheet_name='Annually', index=False)
                if is_present:
                    df2.to_excel(writer, sheet_name='Quarterly', index=False)

            print(f"Saved .xlsx file: {excel_file_path_xlsx}")
            
            xlsm_file_path = os.path.join(save_dir, f"{Company_name.text}_{Date_Span.text}.xlsm")
            
            
            
            today=date.today()
            
            # previous_file_path = f"E:/PYTHON/STOCK_MARKET/ANALYST_ESTIMATES/{today - timedelta(days=15)}/{Company_name.text}.xlsx"  # Update this path as needed
            previous_file_path = f"E:/PYTHON/STOCK_MARKET/ANALYST_ESTIMATES/{previous_date_str}/{Company_name.text}.xlsx"  # Update this path as needed
            
            print(previous_file_path)
            
            # Fetch and compare the first row
            new_row_values = df1.iloc[:4, :-1].values.tolist() if not df1.empty else []
            print("New Row Values:", new_row_values)
            if is_present:
                new_row_values2 = df2.iloc[:4, :-1].values.tolist() if not df2.empty else []
                print("New Row Values:", new_row_values2)

            if os.path.exists(previous_file_path):
                # Load the previous file and get the first row
                sheet_name = 'Annually'
                df = pd.read_excel(previous_file_path, sheet_name=sheet_name, dtype=str, na_filter=False)

                rows_5_to_10 = df.iloc[0:4, :-1]

                # Convert the selected rows into a list of lists
                rows_array = rows_5_to_10.values.tolist()
                print("Old row values ",rows_array)
                
                if is_present:
                    sheet_name = 'Quarterly'
                    df = pd.read_excel(previous_file_path, sheet_name=sheet_name, dtype=str, na_filter=False)

                    rows_5_to_10 = df.iloc[0:4, :-1]

                    # Convert the selected rows into a list of lists
                    rows_array2 = rows_5_to_10.values.tolist()
                    print("Old row values 2 ",rows_array2)
                else:
                    rows_array2=[]
                    new_row_values2=[]


                # Compare the rows
                if rows_array == new_row_values and rows_array2 == new_row_values2:
                    print("No changes in the first row. Skipping save.")
                    # app.quit()
                    return
                else:
                    print("Data in the first row changed. Saving new Excel file.")
            else:
                print("Previous file not found. Saving new Excel file.")
            
            
            # Initialize an xlwings app
            app = xw.App(visible=False)
            wb = app.books.add()
            
            
            

            # Add DataFrames to separate sheets

            if is_present:
                sheet2 = wb.sheets.add('Quarterly')
                sheet2.range('A5').options(index=False).value = df2
            sheet1 = wb.sheets.add('Annually')
            sheet1.range('A5').options(index=False).value = df1

            # Merge and format data on sheets
            sheet1.range('B2:E3').merge()
            sheet1.range('B2').value = Company_name.text
            sheet1.range('F2:I3').merge()
            sheet1.range('F2').value = Date_Span.text
            cell_annually = sheet1.range('B2')
            cell_annually.api.Font.Size = 18
            cell_annually.api.Font.Bold = True
            cell_annually.api.Font.Underline = True

            if is_present:
                sheet2.range('B2:E3').merge()
                sheet2.range('B2').value = Company_name.text
                sheet2.range('F2:I3').merge()
                sheet2.range('F2').value = Date_Span.text
                cell_quarterly = sheet2.range('B2')
                cell_quarterly.api.Font.Size = 18
                cell_quarterly.api.Font.Bold = True
                cell_quarterly.api.Font.Underline = True
            # Define the range where data is placed
            last_row_sheet1 = sheet1.range('A5').expand().last_cell.row
            last_col_sheet1 = sheet1.range('A5').expand().last_cell.column
            if(is_present):
                last_row_sheet2 = sheet2.range('A5').expand().last_cell.row
                last_col_sheet2 = sheet2.range('A5').expand().last_cell.column

            # Apply borders and bold formatting for sheet1 (Annually)
            full_range_sheet1 = sheet1.range(f'A5:{xw.utils.col_name(last_col_sheet1)}{last_row_sheet1}')
            full_range_sheet1.api.Borders.Weight = 2  # Set borders around all cells
            sheet1.range(f'A5:A{last_row_sheet1}').api.Font.Bold = True  # Bold the first column
            # sheet1.range('A:Z').autofit()  # Auto-fit columns based on content in sheet1
            for col in range(1, last_col_sheet1 + 1):
                sheet1.range(sheet1.cells(1, col), sheet1.cells(last_row_sheet1, col)).autofit()


            # Apply borders and bold formatting for sheet2 (Quarterly)
            if(is_present):
                full_range_sheet2 = sheet2.range(f'A5:{xw.utils.col_name(last_col_sheet2)}{last_row_sheet2}')
                full_range_sheet2.api.Borders.Weight = 2  # Set borders around all cells
                sheet2.range(f'A5:A{last_row_sheet2}').api.Font.Bold = True  # Bold the first column
                # sheet2.range('A:Z').autofit()  # Auto-fit columns based on content in sheet2
                for col in range(1, last_col_sheet2 + 1):
                    sheet2.range(sheet2.cells(1, col), sheet2.cells(last_row_sheet2, col)).autofit()

            # Save as macro-enabled file (.xlsm)
            wb.save(xlsm_file_path)
            
            print(f"Saved .xlsm file: {xlsm_file_path}")
            wb.close()
            app.quit()
        except Exception as e:
            print(f"Error in get excel for {Newlink}:{e}")





    Get_excel("https://app.tikr.com/stock/estimates?cid=231651802&tid=704119980&tab=est&ref=g4lonq")#applovin #1
    Get_excel("https://app.tikr.com/stock/estimates?cid=276763615&tid=1673049332&tab=est&ref=g4lonq") #robinhood #2
    Get_excel("https://app.tikr.com/stock/estimates?cid=315460452&tid=541712010&ref=g4lonq&tab=est")#Hellofresh #3
    Get_excel("https://app.tikr.com/stock/estimates?cid=246247134&tid=704022864&ref=g4lonq&tab=est")#Oscar Health #4
    Get_excel("https://app.tikr.com/stock/estimates?cid=533311640&tid=1675971465&ref=g4lonq&tab=est")#Hippo Holdings #5
    Get_excel("https://app.tikr.com/stock/estimates?cid=679030261&tid=697654886&ref=g4lonq&tab=est")#Rush Steet Interactive #6
    Get_excel("https://app.tikr.com/stock/estimates?cid=638035157&tid=691055363&ref=g4lonq&tab=est")#Root #7
    Get_excel("https://app.tikr.com/stock/estimates?cid=105810806&tid=707739955&ref=g4lonq&tab=est")#Vimeo #8
    Get_excel("https://app.tikr.com/stock/estimates?cid=1679658907&tid=1680507426&ref=g4lonq&tab=est")#Olaplex #9
    Get_excel("https://app.tikr.com/stock/estimates?cid=572584066&tid=573484889&ref=g4lonq&tab=est")#Opera Ltd #10
    Get_excel("https://app.tikr.com/stock/estimates?cid=280965739&tid=1864865348&ref=g4lonq&tab=est")#R & S Group Holdings # 11
    Get_excel("https://app.tikr.com/stock/estimates?cid=332752&tid=1790506691&ref=g4lonq&tab=est")#Adtran # 12
    Get_excel("https://app.tikr.com/stock/estimates?cid=145722736&tid=717072548&ref=g4lonq&tab=est")#Bark # 13
    Get_excel("https://app.tikr.com/stock/estimates?cid=3044099&tid=207420139&ref=g4lonq&tab=est")#Aspen Aerogels # 14
    # 


    def send_email_with_attachments(subject, body, folder_path, email_user, email_password, recipient_email):
        try:
            # Create the email message
            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['To'] = recipient_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            allowed_extensions=['.xlsm']
            # Attach all files in the specified folder
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                
                # Only attach files, not directories
                if os.path.isfile(file_path) and any(filename.lower().endswith(ext) for ext in allowed_extensions):
                    with open(file_path, 'rb') as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename= {filename}'
                        )
                        msg.attach(part)

            # Send the email via SMTP server
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(email_user, email_password)
                server.sendmail(email_user, recipient_email, msg.as_string())
                print(f"Email with attachments from {folder_path} sent successfully to {recipient_email}!")
        except Exception as e:
            print(f"Error sending email :{e}")
            

    # Usage example
    send_email_with_attachments(
        subject=f"Updated Analyst Estimates report {date.today()}",
        body="Please find the attached Analyst estimates report.",
        folder_path="E:/PYTHON/STOCK_MARKET/ANALYST_ESTIMATES/"f"{date.today()}",
        email_user=email_user,
        email_password=email_password,
        recipient_email=recipient_email
    )
    send_email_with_attachments(
        subject=f"Updated Analyst Estimates report {date.today()}",
        body="Please find the attached Analyst estimates report.",
        folder_path="E:/PYTHON/STOCK_MARKET/ANALYST_ESTIMATES/"f"{date.today()}",
        email_user=email_user,
        email_password=email_password,
        recipient_email="SM.Isengard@protonmail.com"
    )
    send_email_with_attachments(
        subject=f"Updated Analyst Estimates report {date.today()}:",
        body="Please find the attached Analyst estimates report.",
        folder_path="E:/PYTHON/STOCK_MARKET/ANALYST_ESTIMATES/"f"{date.today()}",
        email_user=email_user,
        email_password=email_password,
        recipient_email="siddarthmehta426@gmail.com"
    )

    print("Mail Send successfully")
except Exception as e:
    print(f"Unexpected error :{e}")
finally:
    driver.quit()
    print("Completed")