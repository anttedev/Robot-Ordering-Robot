from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import time

def embed_screenshot_to_receipt(screenshot, pdf_file):
     """Embeds a picture of the robot to the receipt pdf file"""
     pdf = PDF()

     image_to_add = [
          pdf_file,
          screenshot
     ]
     pdf.add_files_to_pdf(
          files=image_to_add,
          target_document=pdf_file
     )


def take_screenshot_of_robot(order_number):
     """Takes a screenshot of the robot and only the robot"""
     page = browser.page()

     robot_pic = page.locator("//div[@id='robot-preview-image']")
     time.sleep(0.5) 
     """Gives the browser time to fully draw the robot before taking the screenshot"""
     robot_pic.screenshot(path="output/screenshots/" + order_number + ".png")
     

def zip_files():
     """Zips the complete receipts into a zip file"""
     archive = Archive()
     archive.archive_folder_with_zip("output/receipts", "output/receipts.zip")

def open_robot_order_website():
    """Opens the target website"""
    browser.goto("https://robotsparebinindustries.com/?firstname=gf&lastname=DFGDH&salesresult=1235#/robot-order")

def close_constitutional_rights_popup():
     """Closes the annoying pop-up, smh my head"""
     page = browser.page()
     page.click("//button[text()='I guess so...']")

def download_excel():
     """Downloads the excel file with the orders in it"""
     http = HTTP()
     http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
     """Reads the orders.csv file and returns the datatable"""
     table = Tables()
     worksheet = table.read_table_from_csv("orders.csv", header=True)
     return worksheet

def fill_the_form(worksheet):
     """Fills the ordering form"""
     page = browser.page()

     pdf = "output/receipts/receipt_" + worksheet["Order number"] + ".pdf"
     screenshot = "output/screenshots/" + worksheet["Order number"] + ".png"

     page.select_option("//select[@id='head']", worksheet["Head"])
     page.click("//input[@name='body' and @value='" + worksheet["Body"] + "']")
     page.fill("//input[contains(@placeholder, 'legs')]", worksheet["Legs"])
     page.fill("//input[@id='address']", worksheet["Address"])
     
     while True:
          page.click("//button[text()='Order']")
          if page.query_selector("//button[text()='Order another robot']"):
               take_screenshot_of_robot(worksheet["Order number"])
               receipt_as_pdf(worksheet["Order number"])
               embed_screenshot_to_receipt(screenshot, pdf)
               page.click("//button[text()='Order another robot']")
               break

def place_orders():
     """Loops through all the orders"""
     worksheet = get_orders()
     page = browser.page()
     for row in worksheet:
          close_constitutional_rights_popup()
          fill_the_form(row)

def receipt_as_pdf(order_number):
     """Gets the receipt as a PDF file"""
     page = browser.page()
     pdf = PDF()

     receipt = page.locator("//div[@id='receipt']").inner_html()
     pdf.html_to_pdf(receipt, "output/receipts/receipt_" + order_number + ".pdf")

@task
def order_robots_from_RobotSpareBin():
     """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
open_robot_order_website()
download_excel()
place_orders()
zip_files()
