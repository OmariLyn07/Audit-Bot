from tkinter import filedialog
from IPython.display import display
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import tkinter as tk
from tkinter import *
from tkinter.messagebox import showinfo


# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC


def launch_firefox_and_login_with_credentials(sheetname, filename, folderpath, username, password, count=0):
    df = pd.read_excel(filename, sheet_name=sheetname)
    display(df)

    mime_types = "application/pdf,application/vnd.adobe.xfdf,application/vnd.fdf,application/vnd.adobe.xdp+xml"
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.download.dir",
                        fr"{folderpath}")
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
    fp.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)
    fp.set_preference("pdfjs.disabled", True)

    # for increment in range(26):

    # Initialize the Firefox webdriver
    baseUrl = "https://degreeworks.cuny.edu/Dashboard_lc/"
    driver = webdriver.Firefox(firefox_profile=fp)

    try:
        print(folderpath)
        # Load DegreeWorks webpage
        driver.get(baseUrl)

        driver.maximize_window()
        print("Window maximize done")

        time.sleep(2)

        driver.find_element(By.ID, "CUNYfirstUsernameH").send_keys(username)
        time.sleep(1)

        driver.find_element(By.ID, "CUNYfirstPassword").send_keys(password)
        time.sleep(1)
        driver.find_element(By.ID, "CUNYfirstPassword").send_keys(Keys.ENTER)
        time.sleep(2)

        # for row in range(500):
        for row in range(len(df.index)):
            print("Start Row is: " + str(count))
            driver.refresh()
            emplText = str(int(df.iat[count, 0]))
            time.sleep(3)
            degEmplText = driver.find_element(By.ID, "student")
            degEmplText.send_keys(emplText)
            degEmplText.send_keys(Keys.ENTER)
            time.sleep(6)

            # Process New Degree Audit
            # Adding a nested while loop to wait for Process New Button to become visible
            proces_new_button_visible = False
            start_time = time.time()
            while not proces_new_button_visible and time.time() - start_time < 20:  # Set a 20-second timeout:
                try:
                    driver.find_element(
                        By.XPATH,
                        "/html/body/div/div/div[2]/div/main/div/div[3]/div/div[2]/div/div/div[1]/div[2]/div[3]/button") \
                        .click()
                    proces_new_button_visible = True
                except Exception as e:
                    print(f"An error occurred while waiting for PROCESS NEW button to appear at "
                          f"Row Number {str(count)}: {e}")

            # Print the page document - REDO (Omari's Version)
            print_sleep = False
            start_time = time.time()

            # Adding a nested while loop for new audit to actually be processed
            check = False
            while not check and not proces_new_button_visible:
                try:
                    driver.find_element(By.XPATH,
                                        "/html/body/div/div/div[2]/div/main/div/div[3]/div/div[2]/div/div/div[1]/div[2]/div[3]/button")
                    check = True
                except:
                    pass

            while proces_new_button_visible and not print_sleep and time.time() - start_time < 20:  # Set a 20-second timeout
                try:
                    driver.find_element(By.ID, "print-icon").click()
                    driver.find_element(By.XPATH, "/html/body/div[2]/div[3]/div/div[3]/button[2]").click()
                    # Adding 10s wait time for file to actually download
                    time.sleep(10)
                    print_sleep = True
                    print(f"Successfully printed Row Number {str(count)}")
                    count = count + 1
                    print(f"Moving on to Row Number {str(count)}")

                except Exception as e:
                    print(f"An error occurred while trying to print Row Number {str(count)}: {e}")

                # If print_sleep is still False--or Process New Button is still not visible--
                # after 20-second timeout, move on to the "except" block...
            if not print_sleep or not proces_new_button_visible:
                print("Printing timed out. Moving to except block.")
                print(f"Error occurred on Row Number {str(count)}")

    except Exception as e:
        print(f"An error occurred in the EXCEPT BLock for Row Number {str(count)}: {e}")
        # Close Firefox Browser
        driver.quit()

        # Reload webpage with same credentials
        launch_firefox_and_login_with_credentials(sheetname, filename, folderpath, username, password, count=count)
        print("Relaunch has been executed in EXCEPTION Block at Row Number " + str(count))
    finally:
        if count == len(df.index):
            driver.quit()
            root.iconify()
            root.deiconify()
            showinfo("Complete", "All Students Completed!")

        else:
            driver.quit()
            launch_firefox_and_login_with_credentials(sheetname, filename, folderpath, username, password, count=count+1)
            print("Loop has been initiated in FINALLY Block at Row Number " + str(count))

# Call function with initial credentials


def bot_start_op():
    x = text3.get()
    x = x.replace("/" , "\\")
    root.withdraw()
    launch_firefox_and_login_with_credentials(text.get(), text2.get(), x, text4.get(), text5.get())

def folder_button_click():
    folclick = filedialog.askdirectory(initialdir="C:/Users/")
    text3.insert(0, str(folclick))

def file_button_click():
    fileclick = filedialog.askopenfilename(initialdir="C:/Users/")
    text2.insert(0, str(fileclick))


root = tk.Tk()
root.title("Audit Bot")
root.geometry('774x317')

# THE ENTRY PANELS AND LABELS
Infldr = Label(root, text="Excel Sheet Name:")
Infldr.place(x=20, y=35)
text = Entry(root, width=50)
text.place(x=130, y=35)

Path = Label(root, text="Excel File Name:")
Path.place(x=20, y=95)
text2 = Entry(root, width=50)
text2.place(x=130, y=95)

Outfldr = Label(root, text="Output Folder Path:")
Outfldr.place(x=20, y=155)
text3 = Entry(root, width=50)
text3.place(x=130, y=155)

USER = Label(root, text="USERID:")
USER.place(x=450, y=55)
text4 = Entry(root, width=30)
text4.place(x=520, y=55)

PASS = Label(root, text="Password:")
PASS.place(x=450, y=115)
text5 = Entry(root, width=30, show="*")
text5.place(x=520, y=115)


# THE BUTTONS
StartBtn = Button(root, text="START",
                               fg="blue",
                               command=lambda: bot_start_op())
StartBtn.place(x=330, y=265)

filebtn = Button(root, text="Browse", command=lambda: file_button_click())
filebtn.place(x=385, y=115)

foldrbtn = Button(root, text="Browse", command=lambda: folder_button_click())
foldrbtn.place(x=385, y=175)

QuitBtn = Button(root, text="Quit",
                              fg="red", command=lambda: root.quit())
QuitBtn.place(x=400, y=265)


root.mainloop()

