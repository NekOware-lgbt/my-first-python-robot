from time import sleep as time_sleep
import os

from Browser import Browser
from Browser.utils.data_types import SelectAttribute
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Dialogs import Dialogs
from RPA.Robocorp.Vault import Vault, FileSecrets as Vault_FileSecrets


OUT = f'{os.getcwd()}/output/'
OUT_ORDERS = OUT + 'Orders/'
OUT_ORDER = OUT_ORDERS + '/Order_'
browser = Browser()
pdf = PDF()
archive = Archive()
dialogs = Dialogs()
vault = Vault( default_adapter=Vault_FileSecrets )


def ask_user_for_csv_file_url():
    dialogs.add_text_input('url',label='Orders CSV file URL',rows=1)
    result = dialogs.run_dialog()
    return result.url

def open_the_order_website():
    secret_url = vault.get_secret('Robot Order Python Robot').get('ORDER_URL')
    browser.new_browser() # headless=False
    browser.new_page(secret_url)


def download_the_orders_csv_file():
    url = ask_user_for_csv_file_url()
    http = HTTP()
    http.download(
        url=url, # "https://robotsparebinindustries.com/orders.csv"
        overwrite=True,
        target_file=OUT + "orders.csv")
    if not os.path.exists(OUT + 'Orders'):
        os.makedirs(OUT + 'Orders')


def close_start_modal():
    browser.click("text=Yep")
    
def refresh_page_and_close_modal():
    browser.reload()
    browser.wait_for_elements_state('css=div.alert-buttons')
    browser.click('text="Yep"')


def get_that_csv_data():
    tables = Tables()
    return tables.read_table_from_csv(path=OUT + 'orders.csv')

def screenshot_the_preview(num):
    browser.click('text=Preview')
    browser.wait_for_elements_state('css=#robot-preview-image > img[alt="Head"]')
    browser.wait_for_elements_state('css=#robot-preview-image > img[alt="Body"]')
    browser.wait_for_elements_state('css=#robot-preview-image > img[alt="Legs"]')
    return browser.take_screenshot(
        selector='id=robot-preview-image',
        filename=OUT_ORDER + num + '_Preview')

def click_the_order_button():
    while True:
        browser.click('text="Order"')
        alert_count = browser.get_element_count('css=div.alert-danger[role="alert"]')
        if not alert_count or alert_count<1:
            break

def save_the_receipt_to_a_file(num,screensot_path):
    save_path = OUT_ORDER + num + '_Receipt.pdf'
    receipt = browser.get_property(selector='id=receipt',property='outerHTML').replace('</div><div>','<br>').replace('</p><p>','<br>')
    pdf.html_to_pdf(receipt, save_path)
    pdf.add_files_to_pdf(
        files = [screensot_path + ':align=center'],
        target_document = save_path,
        append = True
    )

def fill_a_single_order(order):
    refresh_page_and_close_modal()
    browser.select_options_by('id=head',SelectAttribute.index,order.get('Head'))
    browser.check_checkbox('input[name="body"][value="%(Body)s"]' % order)
    legs_element_id = browser.get_attribute('text=3. Legs:','for')
    browser.type_text('input[id="%s"]' % legs_element_id, order.get('Legs'))
    browser.type_text('id=address', order.get('Address'))
    screensot_path = screenshot_the_preview(order.get('Order number'))
    click_the_order_button()
    save_the_receipt_to_a_file(order.get('Order number'),screensot_path)


def complete_robot_orders_using_csv_file():
    orders = get_that_csv_data()
    for order in orders:
        fill_a_single_order(order)
    #    print(order)
    #fill_a_single_order(orders.get_row(0))
    #element = browser.get_element('text=3. Legs:')
    #print( element.getAttribute(SelectAttribute.label) )
    #browser.select_options_by('id=head',SelectAttribute.index,orders.get_row(0).get('Head'))

def zip_the_orders_output():
    archive.archive_folder_with_zip(OUT_ORDERS,OUT+'Orders.zip')

def delete_dir(in_path: str):
    if os.path.isdir(in_path):
        for entry in os.scandir(in_path):
            if os.path.isdir(entry.path):
                delete_dir(entry.path)
                os.rmdir(entry.path)
            else:
                os.remove(entry.path)
        os.rmdir(in_path)

def clean_output():
    for entry in os.scandir(OUT):
        if entry.name[0:6].lower()=='orders':
            if os.path.isdir(OUT+entry.name):
                delete_dir(OUT+entry.name)
            else:
                os.remove(OUT+entry.name)
            


def main():
    try:
        clean_output()
        download_the_orders_csv_file()
        open_the_order_website()
        #close_start_modal()
        complete_robot_orders_using_csv_file()
        zip_the_orders_output()
        time_sleep(5)
    finally:
        browser.playwright.close()

def alt_main():
    print(ask_user_for_csv_file_url())
    #print(vault.get_secret('Robot Order Python Robot').get('ORDER_URL'))
    #vault.set_secret()


if __name__ == "__main__":
    #alt_main()
    main()