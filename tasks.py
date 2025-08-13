from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
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
    orders = get_orders()
    for order in orders:
        fill_the_form(order)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads excel file from the given URL and get the orders"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    tables = Tables()
    orders = tables.read_table_from_csv(
        "orders.csv",
        header=True,
        columns=["Order number", "Head", "Body", "Legs", "Address"],
    )
    return orders

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def fill_the_form(order):
    """Fill the form for one order"""
    page = browser.page()
    close_annoying_modal()
    page.select_option("#head", order["Head"])
    page.click(f"//input[@name='body' and @value='{order['Body']}']")
    page.fill("xpath=//input[contains(@placeholder,'part number')]", str(order["Legs"]))
    page.fill("#address", order["Address"])
    page.click("#preview")
    while page.locator("#order").count() > 0:
        page.click("#order")
    pdf_file = store_receipt_as_pdf(str(order["Order number"]))
    scr_robot = screenshot_robot(str(order["Order number"]))
    embed_screenshot_to_receipt(scr_robot, pdf_file)
    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    output_file = "output/receipt_" + order_number + ".pdf"
    pdf.html_to_pdf(receipt_html, output_file)
    return output_file

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    output_scr = "output/robot_" + order_number + ".png"
    page.locator("#robot-preview-image").screenshot(path=output_scr)
    return output_scr

def embed_screenshot_to_receipt(scr_robot, pdf_file):
    pdf = PDF()
    list_of_files = [pdf_file, scr_robot]
    pdf.add_files_to_pdf(files=list_of_files, target_document=pdf_file)

def archive_receipts():
    """Creates a ZIP archive of all receipts and images"""
    lib = Archive()
    lib.archive_folder_with_zip("./output", "./output/robots_pdf.zip", include="*.pdf")
