from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Archive import Archive
from RPA.Tables import Tables
from RPA.PDF import PDF

@task
def order_robots_from_RobotSpareBin():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    close_annoying_modal()
    download_csv_file()
    fill_form_with_csv_data()
    archive_receipts()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Ok')")

def fill_and_submit_orders_form(order_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.select_option("#head", str(order_rep["Head"]))
    page.fill("#address", order_rep["Address"])
    page.set_checked(f'#id-body-{order_rep["Body"]}', checked=True)
    page.get_by_placeholder("Enter the part number for the legs").fill(str(order_rep["Legs"]))
    page.click("#order")
    page.click("text=Order")
    is_error = page.is_visible(".alert-danger")
    if is_error:
        page.click("#order")
        if page.is_visible("#order"):
            page.click("#order")
        if page.is_visible("#receipt"):
            store_receipt_as_pdf(order_rep["Order number"])
        if page.is_visible("#robot-preview-image"):
            screenshot_robot(order_rep["Order number"])
        page.click("#order-another")
        close_annoying_modal()
    else:
        if page.is_visible("#receipt"):
            store_receipt_as_pdf(order_rep["Order number"])
        if page.is_visible("#robot-preview-image"):
            screenshot_robot(order_rep["Order number"])
        page.click("#order-another")
        close_annoying_modal()

def download_csv_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def fill_form_with_csv_data():
    """Read data from csv and fill in the sales form"""
    orders = get_orders()
    for row in orders:
        fill_and_submit_orders_form(row)

def store_receipt_as_pdf(order_number):
    page = browser.page()
    order_results_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(order_results_html, f'output/receipt_{order_number}.pdf')

def screenshot_robot(order_number):
    page = browser.page()
    order_results_html = page.locator("#robot-preview-image")
    order_results_html.screenshot(path=f'output/robot-preview-image_{order_number}.png')
    embed_screenshot_to_receipt(order_number)

def embed_screenshot_to_receipt(order_number):
    pdf = PDF()
    list_of_files = [
        f'output/robot-preview-image_{order_number}.png',
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=f'output/receipt_{order_number}.pdf',
        append=True
    )
    pdf.open_pdf(f'output/receipt_{order_number}.pdf')
    pdf.save_pdf(output_path=f'receipts/receipt_{order_number}.pdf')

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('./receipts', 'receipts.zip')