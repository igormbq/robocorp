from robocorp import browser
from robocorp.tasks import task
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import time
import os

@task
def order_robots_from_robotsparebin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Takes a screenshot of the ordered robot.
    Embeds the robot screenshot to the PDF receipt.
    Creates ZIP archive of the receipts.
    """
    browser.configure(
        slowmo=100,
    )
    
    # Create output directory if it doesn't exist
    if not os.path.exists("output"):
        os.makedirs("output")
    
    # Download the orders file
    orders = get_orders()
    
    # Open the robot order website
    open_robot_order_website()
    
    # Process orders
    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
        
        # Submit order and verify success
        if not submit_order():
            print(f"Failed to submit order {row['Order number']}, skipping to next order")
            continue
            
        # Take screenshot with retries
        if not screenshot_robot(row["Order number"]):
            print(f"Failed to take screenshot for order {row['Order number']}, skipping PDF creation")
            continue
            
        pdf_file = store_receipt_as_pdf(row["Order number"])
        
        # Retry if PDF creation failed
        retries = 3
        while not os.path.exists(pdf_file) and retries > 0:
            if not submit_order():
                print(f"Failed to submit order during PDF retry for order {row['Order number']}")
                break
            pdf_file = store_receipt_as_pdf(row["Order number"])
            retries -= 1
            
        order_another_robot()
    
    # Create a ZIP file of the receipts
    archive_receipts()

def get_orders():
    """Downloads orders file and returns it as a table"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv")
    return orders

def open_robot_order_website():
    """Opens the robot order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    """Closes the annoying modal that appears before ordering"""
    page = browser.page()
    try:
        page.click("button:text('OK')", timeout=3000)
    except:
        pass

def fill_the_form(order):
    """Fills in the form for one order"""
    page = browser.page()
    
    # Select the head
    page.select_option("#head", str(order["Head"]))
    
    # Select the body
    page.click(f"#id-body-{order['Body']}")
    
    # Input legs
    page.fill("input[placeholder='Enter the part number for the legs']", str(order["Legs"]))
    
    # Input address
    page.fill("#address", order["Address"])

def submit_order():
    """Submits the order and retries on failure"""
    page = browser.page()
    max_retries = 5
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # Click the order button
            page.click("#order")
            
            # Wait for either error or receipt
            success = False
            try:
                # First check if error appears
                error = page.wait_for_selector(".alert-danger", timeout=2000)
                if error:
                    print(f"Order failed, retrying... (attempt {retry_count + 1})")
                    retry_count += 1
                    time.sleep(1)
                    continue
            except:
                # No error found, check for receipt
                receipt = page.wait_for_selector("#receipt", timeout=2000)
                if receipt:
                    print("Order submitted successfully!")
                    return True
        
        except Exception as e:
            print(f"Error during order submission: {str(e)}")
            retry_count += 1
            time.sleep(1)
    
    print("Failed to submit order after maximum retries")
    return False

def screenshot_robot(order_number):
    """Takes a screenshot of the robot"""
    page = browser.page()
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            print(f"Attempting to take screenshot for order {order_number} (attempt {retry_count + 1})")
            
            # Wait for preview to be visible
            preview = page.wait_for_selector("#robot-preview-image", timeout=5000)
            if not preview:
                print("Preview element not found, retrying...")
                retry_count += 1
                time.sleep(2)
                continue
                
            # Check if preview is actually visible
            if not preview.is_visible():
                print("Preview element found but not visible, retrying...")
                retry_count += 1
                time.sleep(2)
                continue
            
            # Take the screenshot
            preview.screenshot(path=f"output/robot_{order_number}.png")
            print(f"Screenshot taken successfully for order {order_number}")
            return True
            
        except Exception as e:
            print(f"Error taking screenshot: {str(e)}")
            retry_count += 1
            time.sleep(2)
    
    print(f"Failed to take screenshot for order {order_number} after {max_retries} attempts")
    return False

def store_receipt_as_pdf(order_number):
    """Stores the receipt as a PDF file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    
    pdf_path = f"output/receipt_{order_number}.pdf"
    
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, pdf_path)
    
    # Embed the robot screenshot
    pdf.add_files_to_pdf(
        files=[f"output/robot_{order_number}.png"],
        target_document=pdf_path,
        append=True
    )
    
    return pdf_path

def order_another_robot():
    """Clicks the 'Order another robot' button"""
    page = browser.page()
    page.click("#order-another")

def archive_receipts():
    """Creates a ZIP file of all receipts"""
    lib = Archive()
    lib.archive_folder_with_zip("output", "output/receipts.zip", include="*.pdf", recursive=True) 