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
    browser.configure(
        slowmo=100
    )
    open_robot_order_website()
    orders = get_orders()
    for order in orders:
        give_up_rights()
        fill_the_robot_order_form(order)
    
    archive_receipts()


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def give_up_rights():
    page = browser.page()
    page.click("button:text('OK')")


def get_orders():
    """Download the orders file, read it as a table, and return the result"""
    download_orders_file()

    # excel = Files()
    # excel.open_workbook("orders.csv")
    # orders_worksheet = excel.read_worksheet_as_table("orders", header=True)
    # excel.close_workbook()

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )

    return orders


def download_orders_file():
    """Download CSV file"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def fill_the_robot_order_form(order):
    """Fill the robot order form and click 'ORDER'"""
    page = browser.page()
    head = str(order["Head"])
    page.select_option(selector="#head", value=head)
    body = str(order['Body'])
    page.click(f"#id-body-{body}")
    page.fill(selector="input[type='number']", value=str(order["Legs"]))
    page.fill(selector="#address", value=str(order["Address"]))
    while True:
        try:
            page.click("#order", timeout=500)
        except:
            break

    order_number = order["Order number"]
    pdf_file = store_receipt_as_pdf(order_number)
    robot_screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot=robot_screenshot, pdf_file=pdf_file)

    page.click("#order-another")


def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    receipt_pdf = f"output/receipts/receipt_{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, receipt_pdf)
    return receipt_pdf


def screenshot_robot(order_number):
    """Take a screenshot of the robot"""
    page = browser.page()
    robot = page.locator("#robot-preview-image")
    robot_screenshot = f"output/receipts/robots/robot_{order_number}.png"
    robot.screenshot(path=robot_screenshot)
    return robot_screenshot


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embed the robot screenshot to the receipt PDF file"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file
    )


def archive_receipts():
    """Create a ZIP file of receipt PDF files"""
    zip = Archive()
    zip.archive_folder_with_zip(folder="output/receipts", archive_name="output/receipts.zip", recursive=True)
