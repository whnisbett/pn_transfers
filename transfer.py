import sys
import time
from getpass import getpass
import pandas as pd
import numpy as np
import xlrd
from tabulate import tabulate
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import constants

# TODO: Find way to wrap each action in a random wait time using decorators
# TODO: Check that key columns have values
# TODO: Review which columns the transfer values come from


class SettlementCase:
    """
    SettlementCase is a wrapper class that gives a convenient interface for working with transfer data.
    """
    def __init__(self, settlement_series):
        """
        Arguments
        ---------
            - settlement_series (pandas.Series or dict): Single row of the excel spreadsheet corresponding to one settlement.
        """
        self.settlement_series = settlement_series
        self.parse_settlement_series(settlement_series)


    def parse_settlement_series(self, settlement_series):
        """
        Parses a settlement_series object and assigns appropriate values to easy-to-interpret attributes.
        """
        self.client_name = settlement_series["NAME"]
        self.amount_iolta_to_operating = settlement_series["IOLTA to Business"]
        self.amount_operating_to_marketing = settlement_series["MKT ACCT"]
        self.amount_operating_to_cash = settlement_series["CASH'S LOAN"]
        self.amount_operating_to_building = settlement_series["BUILDING BONUS"]


class SettlementData:
    """
    SettlementData provides an interface for parsing the settlement spreadsheet as well as selecting and executing bank transfers on rows of the spreadsheet.
    """
    def __init__(self, file_path):
        """
        Reads and preprocesses settlement spreadsheet, then asks user to specify the rows they would like to perform transfers on.

        Arguments
        ---------
            - file_path (str): Path to settlement spreadsheet (which should be in XLSX format)
        """
        self.file_path = file_path
        self.sheet_name = self.get_settlement_sheet_name(file_path)
        self.settlements_df = self.read_settlements_file(file_path, sheet_name=self.sheet_name)
        self.preprocess_settlements_df()
        self.validate_columns()
        self.settlement_rows = self.select_settlements_by_row()

    def validate_columns(self):
        """
        Validate that required columns exist in the selected spreadsheet
        """
        missing_columns = set(constants.REQUIRED_COLUMNS) - set(self.settlements_df.columns)
        if len(missing_columns) != 0:
            print(f"Columns {missing_columns} are missing from the spreadsheet. Please ensure that these columns exist and are spelled correctly before attempting a transfer.")
            sys.exit()
    
    def get_settlement_sheet_name(self, file_path):
        """
        Get name of desired sheet from excel file to perform transfers on.
        """
        xl_file = pd.ExcelFile(file_path, engine="openpyxl")
        sheet_names = xl_file.sheet_names
        desired_sheet = input("Which sheet would you like to select for transfers?\n")
        while desired_sheet not in sheet_names:
            desired_sheet = input(f"Sheet not found. Please enter a sheet from the following list:\n{sheet_names}\n")
        return desired_sheet

    def read_settlements_file(self, file_path, **kwargs):
        """
        Read XLSX file containing settlement information at file_path
        """
        df = pd.read_excel(file_path, engine="openpyxl", **kwargs)
        return df

    def preprocess_settlements_df(self):
        """
        Preprocess settlements_df using the following steps
            1. Strip leading and lagging whitespaces in column names
            2. Drop empty rows and columns
        """
        self.settlements_df.columns = [
            col.strip() for col in self.settlements_df.columns
        ]
        self.settlements_df = self.settlements_df.dropna(axis=1, how="all")
        self.settlements_df = self.settlements_df.dropna(
            axis=0, subset=constants.REQUIRED_COLUMNS
        )

    def select_settlements_by_row(self):
        """
        Select a settlement from self.settlements_df based on the corresponding row from the raw Excel spreadsheet.
        """
        row_input = input("Which row numbers would you like to perform a transfer for?\n")
        rows = self.parse_row_input(row_input)
        idxs = [row - 2 for row in rows]
        try:
            sub_df = self.settlements_df.loc[idxs]
        except KeyError:
            print(f"Row unavailable for transfer. Please ensure that the following required columns are filled out:\n{constants.REQUIRED_COLUMNS}")
            return self.select_settlements_by_row()
        
        input_confirmed = self.row_input_is_correct(sub_df)
        if input_confirmed:
            return sub_df
        else:
            reinput_prompt = (
                "Would you like to try inputting the rows again again? (yes/no)\n"
            )
            reinput_confirmed = self.response_is_yes(reinput_prompt)
            if reinput_confirmed:
                return self.select_settlements_by_row()
            else:
                print("No rows selected")
                return None

    def parse_row_input(self, row_input):
        """
        Converts string input from user into list of rows for subsetting data
        """
        rows = row_input.split(",")
        rows = [int(row.strip()) for row in rows]
        return rows

    def row_input_is_correct(self, df_rows):
        """
        Checks with user that the rows they selected are the desired rows.
        """
        tabulated_rows = tabulate(
            df_rows[constants.DISPLAY_COLUMNS], headers="keys", tablefmt="psql"
        )
        prompt = f"{tabulated_rows}\nAre these the correct rows? (yes/no) \n"
        return self.response_is_yes(prompt)

    def append_more_settlement_rows(self):
        """
        Function for user to add settlements after initializing object.
        """
        df_append = self.select_settlements_by_row()
        tabulated_rows = tabulate(
            df_append[constants.DISPLAY_COLUMNS], headers="keys", tablefmt="psql"
        )

        prompt = f"{tabulated_rows} \n Would you like to add these rows to the transfer operation? (yes/no) \n"
        if self.response_is_yes(prompt):
            self.settlement_rows = pd.concat([self.settlement_rows, df_append])
        else:
            print("No additional settlements added")

    def response_is_yes(self, prompt):
        """
        Prompts response from user based on prompt (str) and returns response (which should be either y(es) or n(o)) as a boolean. 
        """
        response = input(prompt).lower().strip()

        if response[:1] == "y":
            return True
        elif response[:1] == "n":
            return False
        else:
            print('Please respond with "yes" or "no"')
            return self.response_is_yes(prompt)

    def execute_transfers(self):
        """
        Execute bank transfers on the rows specified by user at initialization.
        """
        self.preprocess_settlement_rows_for_transfer()
        tabulated_rows = tabulate(
            self.settlement_rows[constants.DISPLAY_COLUMNS],
            headers="keys",
            tablefmt="psql",
        )
        prompt = f"{tabulated_rows} \n Would you like to execute transfers on these rows? (yes/no)\n"
        transfer_confirmed = self.response_is_yes(prompt)
        if not transfer_confirmed:
            print("Transfer not executed")
            return
        print('Executing transfers...')
        transfer_items = self.convert_settlement_rows_to_transfer_items()
        perform_transfers_via_selenium(transfer_items)

    def preprocess_settlement_rows_for_transfer(self):
        """
        Preprocess the settlement items before executing bank transfer by doing the following things:
            1. Drop any duplicates so that no transfer is performed twice
        """
        self.settlement_rows = self.settlement_rows.drop_duplicates()

    def convert_settlement_rows_to_transfer_items(self):
        """
        Iterate over settlement_rows and create a SettlementCase for each.
        """
        return [SettlementCase(row) for _, row in self.settlement_rows.iterrows()]


def sleep_random_time(max, min=0):
    """
    Sleep for a random time sampled from the interval [min, max] (uniformly distributed).
    """
    t = (max - min) * np.random.rand() + min
    time.sleep(t)


def perform_transfers_via_selenium(transfer_items):
    """
    Perform appropriate transfers for cases in transfer_items using Selenium webdriver.
    """
    browser = initialize_browser()
    wait_for_url_load(browser, url="https://www.frostbank.com/")
    sleep_random_time(5 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
    execute_login(browser)
    wait_for_url_load(browser, url="https://www.frostbank.com/mf/accounts/main")
    sleep_random_time(5 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
    verify_login(browser)
    navigate_to_transfers(browser)
    for settlement_item in transfer_items:
        transfer_dict = initialize_transfer_dict(settlement_item)
        for order in transfer_dict:
            wait_for_url_load(browser, url="https://www.frostbank.com/mf/transfers/main")
            sleep_random_time(5 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
            verify_navigation_to_transfers(browser)
            complete_transfer_form(browser, order)
            click_next(browser)
            wait_for_url_load(browser, url="https://www.frostbank.com/mf/transfers/verify")
            sleep_random_time(5 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
            submit_transfer(browser)
            wait_for_url_load(browser, url="https://www.frostbank.com/mf/transfers/confirm")
            sleep_random_time(5 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
            click_make_another_transfer(browser)

    browser.close()


def initialize_browser():
    """
    Initializes and returns a chromedriver browser
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-extensions")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--start-maximized")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"
    )
    browser = webdriver.Chrome(executable_path=constants.DRIVERPATH, options=options)
    browser.delete_all_cookies()
    browser.set_window_size(800, 800)
    browser.set_window_position(0, 0)
    browser.get(constants.HOMEPAGE)
    return browser


def execute_login(browser):
    """
    Performs login by prompting by receiving username and password from user
    """
    i = 0
    try:
        user_field = browser.find_element_by_id("username-field")
    except:
        i = i + 1
        if i < 5:
            sleep_random_time(3 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
        else:
            raise Exception('Unable to load main page. Please try again.')
    user_field.clear()
    user = input("Enter Username: ")
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)
    user_field.send_keys(user)
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)
    pass_field = browser.find_element_by_id("password-field")
    pass_field.clear()
    pwd = getpass()
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)
    pass_field.send_keys(pwd)
    pass_field.send_keys(Keys.RETURN)


def verify_login(browser):
    """
    Verify that login was successful.
    """
    if not browser.current_url == "https://www.frostbank.com/mf/accounts/main":
        print(
            "Failed to login. Check that username and password were entered correctly"
        )
        sys.exit()


def navigate_to_transfers(browser):
    """
    Navigate from main accounts page to transfer page by:
        1. Tab x 3
        2. Down arrow
        3. Enter
    """
    try:
        transfers_tab = browser.find_element_by_id("tabTransfers")
    except:
        i = i + 1
        if i < 5:
            sleep_random_time(3 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
        else:
            raise Exception('Unable to load main accounts page. Please try again.')
    actions = ActionChains(browser)
    for _ in range(3):
        # actions.key_down(Keys.SHIFT)
        actions.send_keys(Keys.TAB)
        # actions.key_up(Keys.SHIFT)
        wait_time = 0.5 * np.random.rand() + 0.5
        actions.pause(wait_time)
    actions.send_keys(Keys.ARROW_DOWN)
    wait_time = 0.5 * np.random.rand() + 0.5
    actions.pause(wait_time)
    actions.send_keys(Keys.ENTER)
    wait_time = 0.5 * np.random.rand() + 0.5
    actions.pause(wait_time)
    actions.perform()


def initialize_transfer_dict(settlement_item):
    """
    Extract transfer info from settlement_item and store in dict. This could be replaced with a BankTransfer object and SettlementCase could just be called SettlementCase in the future...
    """
    transfer_dict = [
    dict(
        from_acct="Iolta",
        to_acct="Operating",
        from_acct_num=str(constants.IOLTA_ACCT_NUM),
        to_acct_num=str(constants.OPERATING_ACCT_NUM),
        amount=settlement_item.amount_iolta_to_operating,
        name=settlement_item.client_name
    ),
    dict(
        from_acct="Operating",
        to_acct="Marketing",
        from_acct_num=str(constants.OPERATING_ACCT_NUM),
        to_acct_num=str(constants.MARKETING_ACCT_NUM),
        amount=settlement_item.amount_operating_to_marketing,
        name=settlement_item.client_name
    ),
    dict(
        from_acct="Operating",
        to_acct="Cash Loan Repayment Fund",
        from_acct_num=str(constants.OPERATING_ACCT_NUM),
        to_acct_num=str(constants.CASH_ACCT_NUM),
        amount=settlement_item.amount_operating_to_cash,
        name=settlement_item.client_name
    ),
    dict(
        from_acct="Operating",
        to_acct="Building",
        from_acct_num=str(constants.OPERATING_ACCT_NUM),
        to_acct_num=str(constants.BUILDING_ACCT_NUM),
        amount=settlement_item.amount_operating_to_building,
        name=settlement_item.client_name
    ),
    ]
    return transfer_dict


def verify_navigation_to_transfers(browser):
    """
    Verify that the browser successfully navigated to transfers page.
    """
    if not browser.current_url == "https://www.frostbank.com/mf/transfers/main":
        raise Exception(
            "Failed to navigate to transfers page. Check that the specified actions correctly take you to the transfers page."
        )


def complete_transfer_form(browser, order):
    """
    Complete transfer form based on order.
    """
    select_from_account(browser, order)
    select_to_account(browser, order)
    insert_amount(browser, order)
    insert_memo(browser, order)


def select_from_account(browser, order):
    """
    Select account to transfer funds from based on order.
    """
    try:
        from_acct_dropdown = browser.find_element_by_id("from-account-list")
    except:
        i = i + 1
        if i < 5:
            sleep_random_time(3 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
        else:
            raise Exception('Unable to load transfers page. Please try again.')

    from_acct_dropdown.click()
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)
    from_acct_row = from_acct_dropdown.find_elements_by_xpath(
        f"//*[contains(text(), \"{order['from_acct_num']}\")]"
    )[0]
    from_acct_row.click()
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2* constants.SPEED_FACTOR)
    

def select_to_account(browser, order):
    """
    Select account to transfer funds to based on order.
    """
    to_acct_dropdown = browser.find_element_by_id("to-account-list")
    to_acct_dropdown.click()
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)
    to_acct_row = to_acct_dropdown.find_elements_by_xpath(f"//*[contains(text(), \"{order['to_acct_num']}\")]")[0]
    to_acct_row.click()
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)


def insert_amount(browser, order):
    """
    Insert amount to be transferred based on order.
    """
    amount_field = browser.find_element_by_id("amount")
    amount_field.send_keys(str(order["amount"]))
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)


def insert_memo(browser, order):
    """
    Insert memo (client's name) based on order.
    """
    memo_field = browser.find_element_by_id("memo")
    memo_field.send_keys(order['name'])
    sleep_random_time(5 * constants.SPEED_FACTOR, min=2 * constants.SPEED_FACTOR)


def click_next(browser):
    """
    Click "Next" button after completing transfer form
    """
    next_btn = browser.find_element_by_id("btn-next")
    next_btn.click()


def submit_transfer(browser):
    """
    Submit transfer form by clicking appropriate button.
    """
    try:
        submit_btn = browser.find_element_by_id("btn-submit")
    except:
        i = i + 1
        if i < 5:
            sleep_random_time(3 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
        else:
            raise Exception('Unable to load confirmation page. Please try again.')
    submit_btn.click()


def click_make_another_transfer(browser):
    """
    Click "Make another transfer" button after submitting transfer
    """
    try:
        submit_btn = browser.find_element_by_id("btn-submit")
    except:
        i = i + 1
        if i < 5:
            sleep_random_time(3 * constants.SPEED_FACTOR, 2 * constants.SPEED_FACTOR)
        else:
            raise Exception('Unable to load completion page. Please try again.')
    submit_btn.click()


def wait_for_url_load(browser, url, wait_time=1.5, max_iter=10):
    """
    Sleep until URL has loaded and abort after 10 failed iterations.
    """
    i = 0
    while browser.current_url != url:
        i = i + 1
        if i > max_iter:
            print(f"Failed to reach {url}. Aborting, please try again.")
        time.sleep(wait_time)


if __name__ == "__main__":
    settlement_xlsx_path = sys.argv[1]
    settlement_data = SettlementData(settlement_xlsx_path)
    if settlement_data.settlement_rows is None:
        print("No settlement cases selected, aborting transfer. Please select at least one row before executing a transfer.")
        sys.exit()
    settlement_data.execute_transfers()