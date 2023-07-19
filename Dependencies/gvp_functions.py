from datetime import date
from selenium.webdriver.chrome.webdriver import WebDriver
import os
from pathlib import Path
import sys
from setup import setup

class download_profile:
    def __init__(self, url, file_name, prompt) -> None:
        self.url = url
        self.file_name = file_name
        self.prompt = prompt

def fix_makepy():
    from shutil import rmtree, copytree
    from subprocess import check_call
    from sys import executable 
    """
    There is an issue where when trying to launch using EnsureDispatch, 
    the win32com.client is unable to do so and runs into the error to run makepy manually.
    In order to correct for that, this except block catches the raised errors, 
    deletes the corrupted gen_py folder, then runs the makepy file manually before launching again.
    """
    try:
        gen_py_path = os.path.join(
            Path.home(), 'AppData', 'Local', 'Temp', 'gen_py')
        if os.path.exists(gen_py_path):
            rmtree(gen_py_path)
        check_call([executable, '-m', 'win32com.client.makepy', 'Excel.Application'])
    except:
        backup_path = r"" # Network Share drive
        gen_py_path = os.path.join(
            Path.home(), 'AppData', 'Local', 'Temp', 'gen_py')
        
        if os.path.exists(gen_py_path):
            rmtree(gen_py_path)
        
        copytree(backup_path, gen_py_path)

def column_cleaner(df):
    # delcaring what is not wanted in the column names and an empty list to hold the new names
    not_wanted = ['ID', 'DESC']
    new_column_names = []

    # checking each header for the unwanted parts
    for column in df.columns:
        name_list = column.split()
        last_word = name_list[-1]
        if (last_word in not_wanted) and (name_list[-2] != 'Management'):
            name_list.pop(-1)

        name = ' '.join(name_list)
        new_column_names.append(name)
    # print('New Column Names:')
    # [print(x) for x in new_column_names]

    # renaming the headers
    df.columns = new_column_names
    return df

def correct_export_options(driver: WebDriver):
    # import dependencies
    from time import sleep
    from .setup import setup

    # import dependencies
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
    except ImportError: 
        setup()
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By


    options_url = "" # Microstrategy URL
    export_url = '' # Microstrategy URL
    driver.get(options_url)
    sleep(3)
    driver.get(export_url)

    export_option_xpath = '//*[@id="exportShowOptions"]'
    export_element = driver.find_element(By.XPATH, export_option_xpath)
    ActionChains(driver).move_to_element(export_element).perform() # type: ignore

    checked_export = export_element.is_selected()

    if checked_export is False:
        print('Correcting Export options')
        export_element.click()
        driver.find_element(By.XPATH, '//*[@id="25003"]').click()

def restore_export_options(driver: WebDriver):
    # import dependencies
    from time import sleep
    from .setup import setup

    # import dependencies
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
    except ImportError: 
        setup()
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By


    options_url = "" # Microstrategy URL
    export_url = '' # Microstrategy URL
    driver.get(options_url)
    sleep(3)
    driver.get(export_url)

    export_option_xpath = '//*[@id="exportShowOptions"]'
    export_element = driver.find_element(By.XPATH, export_option_xpath)
    ActionChains(driver).move_to_element(export_element).perform() # type: ignore

    checked_export = export_element.is_selected()

    if checked_export:
        print('Restoring Export options')
        export_element.click()
        driver.find_element(By.XPATH, '//*[@id="25003"]').click()

def decide_we_sat(date_obj: date) -> date:
    from dateutil.relativedelta import relativedelta, SA
    return date_obj + relativedelta(weekday=SA)

# declaring fiscal month calculator functions
# takes in a date, and decides which fiscal month it is.
def decide_fm(date_obj: date, return_date='month') -> date:
    """
    Takes in a date object and depending on the return_date argument, it will return
    the fiscal month, FM start date, or FM end date as a date object.
    """
    def isleap(year: int) -> bool:
        # Return True for leap years, False for non-leap years.
        return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
    
    accepted_return_dates = ('month', 'end', 'beginning')

    year = date_obj.year
    month = date_obj.month
    
    if return_date =='month':
        day = 1

        if (month == 12) and (date_obj.day > 28):
            year = date_obj.year + 1
            month = 1
        elif date_obj.day > 28:
            month = date_obj.month + 1
    
    elif return_date == 'beginning':
        month = month - 1
        day = 29

        if date_obj.month == 1:
            year = year - 1
        elif (isleap(date_obj.year) == False) and (date_obj.month == 3):
            month = date_obj.month
            day = 1

    elif return_date == 'end':
        year = date_obj.year
        month = date_obj.month
        day = 28

        if (month == 12) and (date_obj.day > 28):
            year = date_obj.year + 1
            month = 1
        elif date_obj.day > 28:
            month = date_obj.month + 1

    else:
        raise TypeError(f'{return_date} is not an accepted argument for return_date. Accepted arguments are {", ".join(accepted_return_dates)}')
    
    date_obj = date_obj.replace(year=year, month=month, day=day)
    return date_obj
    

# takes in a date and decides what the beginning day of the fiscal month is
def decide_fm_beginning(date_obj: date) -> date:
    try: 
        from dateutil.relativedelta import relativedelta
    except ImportError:
        setup()
        from dateutil.relativedelta import relativedelta


    def isleap(year: int) -> bool:
        # Return True for leap years, False for non-leap years.
        return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)

    year = date_obj.year
    month = (date_obj + relativedelta(months=-1)).month
    day = 29

    if date_obj.month == 1:
        year = year - 1
    elif (isleap(date_obj.year) == False) and (date_obj.month == 3):
        month = date_obj.month
        day = 1

    date_obj = date_obj.replace(year=year, month=month, day=day)
    return date_obj

# takes in a date and decides what the end of the fiscal month is
def decide_fm_end(date_obj: date) -> date:

    year = date_obj.year
    month = date_obj.month
    day = 28

    if (month == 12) and (date_obj.day > 28):
        year = date_obj.year + 1
        month = 1
    elif date_obj.day > 28:
        month = date_obj.month + 1

    date_obj = date_obj.replace(year=year, month=month, day=day)
    return date_obj


def download_reports(driver: WebDriver, mstr_url: str, download_file: str, export_type='excel', download_folder=os.path.join(Path.home(), 'Downloads'), timeout=120, **kwargs) -> None:
    import os
    from time import sleep

    # import dependencies
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.select import Select
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.common.exceptions import TimeoutException
        from dateutil.relativedelta import relativedelta
    except ImportError or ModuleNotFoundError:
        setup()
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.select import Select
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.common.exceptions import TimeoutException
        from dateutil.relativedelta import relativedelta
    
    
    def answer_prompts_fm(fiscal_month:date, months:int) -> None:
        """
        This function takes in the fiscal month kwarg and months kwargs in order to answer the mictrostrategy prompts
        for fiscal months
        """
        today = date.today()
        yesterday = today - relativedelta(days=1)

        current_fm = decide_fm(yesterday)
        month_diff = relativedelta(current_fm, fiscal_month)
        print(month_diff)

        xpath_month = 1 + month_diff.months + (month_diff.years * 12)

        # answers the fiscal months prompts and click throughs in order to get to the export screen
        selected_month = f'//*[@id="id_mstr47ListContainer"]/div[{xpath_month}]/div'
        selector = '//*[@id="id_mstr49"]/img'
        remove_all_xpath = '//*[@id="id_mstr52"]/img'
        
        # wait until the remove all button is available before removing previous selected months move to the element if not on screen
        WebDriverWait(driver=driver, timeout=10).until(
            lambda x: x.find_element(By.XPATH, remove_all_xpath))
        ActionChains(driver).move_to_element(
            driver.find_element(By.XPATH, remove_all_xpath)).perform()
        driver.find_element(By.XPATH, remove_all_xpath).click()

        # loop to scroll through until needed option is available.
        for stop_gap in range(xpath_month):
            stop_gap_xpath = f'//*[@id="id_mstr47ListContainer"]/div[{stop_gap+1}]/div'
            WebDriverWait(driver=driver, timeout=10).until(
                lambda x: x.find_element(By.XPATH, stop_gap_xpath))
            ActionChains(driver).scroll_to_element(
                driver.find_element(By.XPATH, stop_gap_xpath)).perform()
        for i in range(months):
            ActionChains(driver).scroll_to_element(
            driver.find_element(By.XPATH, selected_month)).perform()
        
            # month_text = driver.find_element(By.XPATH, selected_month).text
            driver.find_element(By.XPATH, selected_month).click()

            ActionChains(driver).move_to_element(
                        driver.find_element(By.XPATH, selector)).perform()
            driver.find_element(By.XPATH, selector).click()
            sleep(1)

    def answer_prompts_we(week_end: date, weeks:int) -> None:
        
        today = date.today()
        yesterday = today - relativedelta(days=1)

        current_we = decide_we_sat(yesterday)
        week_diff = relativedelta(current_we, week_end)

        xpath_week = 1 + week_diff.weeks + (week_diff.years * 52)

        selected_week = f'/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[2]/span/div[2]/div[1]/table/tbody/tr[3]/td[1]/table[2]/tbody/tr/td/span/div[3]/div[1]/div[{xpath_week}]/div'
        selector = '/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[2]/span/div[2]/div[1]/table/tbody/tr[3]/td[2]/div[1]/div/img'
        remove_all_xpath = '/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[2]/span/div[2]/div[1]/table/tbody/tr[3]/td[2]/div[5]/div/img'
        
        
        
        # wait until the remove all button is available before removing previous selected months move to the element if not on screen
        WebDriverWait(driver=driver, timeout=10).until(
            lambda x: x.find_element(By.XPATH, remove_all_xpath))
        ActionChains(driver).move_to_element(
            driver.find_element(By.XPATH, remove_all_xpath)).perform()
        driver.find_element(By.XPATH, remove_all_xpath).click()

        # loop to scroll through until needed option is available.
        for stop_gap in range(xpath_week):
            stop_gap_xpath = f'/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[2]/span/div[2]/div[1]/table/tbody/tr[3]/td[1]/table[2]/tbody/tr/td/span/div[3]/div[1]/div[{stop_gap+1}]/div'
            WebDriverWait(driver=driver, timeout=10).until(
                lambda x: x.find_element(By.XPATH, stop_gap_xpath))
            ActionChains(driver).scroll_to_element(
                driver.find_element(By.XPATH, stop_gap_xpath)).perform()
        for i in range(weeks):
            ActionChains(driver).scroll_to_element(
            driver.find_element(By.XPATH, selected_week)).perform()
        
            # month_text = driver.find_element(By.XPATH, selected_month).text
            driver.find_element(By.XPATH, selected_week).click()

            ActionChains(driver).move_to_element(
                        driver.find_element(By.XPATH, selector)).perform()
            driver.find_element(By.XPATH, selector).click()
            sleep(1)
    
    def answer_prompts_year(fiscal_year:int, years:int) -> None:
        """
        This function takes in the fiscal year kwarg and months kwargs in order to answer the mictrostrategy prompts
        for fiscal years
        """
        today = date.today()
        yesterday = today - relativedelta(days=1)

        current_fm = decide_fm(yesterday)
        year_diff = current_fm.year - fiscal_year

        xpath_year = 1 + year_diff

        # answers the fiscal months prompts and click throughs in order to get to the export screen
        first_option = '/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[1]/span/div[2]/div[1]/table/tbody/tr[3]/td[1]/table[2]/tbody/tr/td/span/div[3]/div[1]/div[1]/div'
        
        selector = '/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[1]/span/div[2]/div[1]/table/tbody/tr[3]/td[2]/div[1]/div/img'
        remove_all_xpath = '/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[1]/span/div[2]/div[1]/table/tbody/tr[3]/td[2]/div[5]/div/img'
        
        # wait until the remove all button is available before removing previous selected months move to the element if not on screen
        WebDriverWait(driver=driver, timeout=10).until(
            lambda x: x.find_element(By.XPATH, remove_all_xpath))
        ActionChains(driver).move_to_element(
            driver.find_element(By.XPATH, remove_all_xpath)).perform()
        driver.find_element(By.XPATH, remove_all_xpath).click()

        xpath_year = current_fm.year - int(driver.find_element(By.XPATH, first_option).text) + 1

        selected_month = f'/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[1]/span/div[2]/div[1]/table/tbody/tr[3]/td[1]/table[2]/tbody/tr/td/span/div[3]/div[1]/div[{xpath_year}]/div'

        # loop to scroll through until needed option is available.
        for stop_gap in range(xpath_year):
            stop_gap_xpath = f'/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/div[1]/span/div[2]/div[1]/table/tbody/tr[3]/td[1]/table[2]/tbody/tr/td/span/div[3]/div[1]/div[{stop_gap+1}]/div'
            WebDriverWait(driver=driver, timeout=10).until(
                lambda x: x.find_element(By.XPATH, stop_gap_xpath))
            ActionChains(driver).scroll_to_element(
                driver.find_element(By.XPATH, stop_gap_xpath)).perform()
        for i in range(years):
            ActionChains(driver).scroll_to_element(
            driver.find_element(By.XPATH, selected_month)).perform()
        
            # month_text = driver.find_element(By.XPATH, selected_month).text
            driver.find_element(By.XPATH, selected_month).click()

            ActionChains(driver).move_to_element(
                        driver.find_element(By.XPATH, selector)).perform()
            driver.find_element(By.XPATH, selector).click()
            sleep(1)

    supported_types = ('excel', 'csv')
    if export_type not in supported_types:
        raise TypeError(f'{export_type} is not a supported export type')
    
    prompt = kwargs.get('prompt')
    yesterday = date.today() - relativedelta(days=1)

    if prompt == 'fm':
        if 'fiscal_month' not in kwargs:
            kwargs['fiscal_month'] = decide_fm(yesterday) # setting default fiscal month to yesterday's fiscal
        elif type(kwargs.get('fiscal_month')) != date:
            raise TypeError('fiscal_month argument must be of date type')
        if 'months' not in kwargs:
            kwargs['months'] = 1 # setting default amount of months to pull to one
        elif type(kwargs.get('months')) != int:
            raise TypeError('months argument must be of int type')
    elif prompt == 'we':
        if 'week_end' not in kwargs:
            kwargs['week_end'] = decide_we_sat(yesterday - relativedelta(weeks=1)) # setting default week end to last completed week
        elif type(kwargs.get('week_end')) != date:
            raise TypeError('weekend argument must be of date type')
        if 'weeks' not in kwargs:
            kwargs['weeks'] = 1 # default number of weeks to pull is 1
        elif type(kwargs.get('weeks')) != int:
            raise TypeError('weeks argument must be of int type')
    if prompt == 'fy':
        if 'fiscal_year' not in kwargs:
            kwargs['fiscal_year'] = decide_fm(yesterday).year # setting default fiscal month to yesterday's fiscal
        elif type(kwargs.get('fiscal_year')) != int:
            raise TypeError('fiscal_year argument must be of int type')
        if 'years' not in kwargs:
            kwargs['years'] = 1 # setting default amount of months to pull to one
        elif type(kwargs.get('years')) != int:
            raise TypeError('years argument must be of int type')

    # looping through the dataframe and pulling data from Mstr
    print(f'Starting {download_file}')
    
    # removing old call data if in downloads
    # downloads_folder = os.path.join(Path.home(), 'Downloads')
    
    download_path = os.path.join(download_folder, download_file)
    if os.path.exists(download_path):
        os.remove(download_path)
        print('Old copy of report removed...')

    driver.get(mstr_url)
    sleep(2)



    report_run_xpath = '//input[@value="Export"]'
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
    driver.find_element(By.XPATH, report_run_xpath).click()

    if prompt == 'fm':
        fiscal_month = kwargs['fiscal_month']
        months = kwargs['months']
        answer_prompts_fm(fiscal_month, months)
    elif prompt == 'we':
        week_end = kwargs['week_end']
        weeks = kwargs['weeks']
        answer_prompts_we(week_end, weeks)
    elif prompt == 'fy':
        fiscal_year = kwargs['fiscal_year']
        years = kwargs['years']
        answer_prompts_year(fiscal_year, years)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
    driver.find_element(By.XPATH, report_run_xpath).click()

    sleep(2)
    # checking for export page and answering prompts
    export_corrected = False

    try:
        WebDriverWait(driver, 30).until(EC.title_contains('Export Options'))
    except (TimeoutException):
        correct_export_options(driver)
        export_corrected = True

        if os.path.exists(download_path):
            os.remove(download_path)

        driver.get(mstr_url)
        sleep(2)
        report_run_xpath = '//input[@value="Export"]'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
        driver.find_element(By.XPATH, report_run_xpath).click()

        if prompt == 'fm':
            fiscal_month = kwargs['fiscal_month']
            months = kwargs['months']
            answer_prompts_fm(fiscal_month, months)
        elif prompt == 'we':
            week_end = kwargs['week_end']
            weeks = kwargs['weeks']
            answer_prompts_we(week_end, weeks)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
        driver.find_element(By.XPATH, report_run_xpath).click()
        WebDriverWait(driver, 10).until(EC.title_contains('Export Options'))

    print("Export Options Found")
    excel_button = '//*[@id="exportFormatGrids_excelPlaintextIServer"]'
    if export_type == 'csv':
        excel_button = '//*[@id="exportFormatGrids_csvIServer"]'

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, excel_button)))

    checked_excel = driver.find_element(By.XPATH, excel_button).is_selected()
    if checked_excel is False:
        driver.find_element(By.XPATH, excel_button).click()

    report_title = '//*[@id="exportReportTitle"]'
    checked_title = driver.find_element(By.XPATH, report_title).is_selected()
    if checked_title:

        driver.find_element(By.XPATH, report_title).click()

    filter_details = '//*[@id="exportFilterDetails"]'
    checked_filter = driver.find_element(
        By.XPATH, filter_details).is_selected()
    if checked_filter:
        driver.find_element(By.XPATH, filter_details).click()

    extra_column = '//*[@id="exportOverlapGridTitles"]'
    select = Select(driver.find_element(By.XPATH, extra_column))
    select.select_by_visible_text('Yes')

    export_button = '//*[@id="3131"]'
    ActionChains(driver).move_to_element(
        driver.find_element(By.XPATH, export_button)).perform()
    driver.find_element(By.XPATH, export_button).click()

    count = 0
    # sleeping until the file is downloaded.
    while os.path.exists(download_path) == False:
        if count > timeout:
            print(f'{download_file} failed after {timeout/60} minutes')
            sys.exit()
        count += 1
        if count % 5 == 0:
            print(f'Waiting for {download_file} to finish downloading... (Attempt {count} of {timeout})')
        sleep(3)
    else:
        print(f'Report Downloaded to {download_folder}')
    
    if export_corrected:
        restore_export_options(driver)


def generate_email(explainer_html:str, subject:str, email_type:str, recipients:list, cc=['Network_Email_Address'], embed_images={}) -> None:
    try:
        import win32com.client
    except ImportError:
        setup('pywin32')
        import win32com.client

    recipients_string = '; '.join(recipients)
    cc_string = '; '.join(cc)
    
    src_folder = os.path.join(os.path.dirname(__file__), 'src')
    logo = os.path.join(src_folder, 'logo.png')

    supported_emails = ('comp', 'dept', 'lead', 'leader')

    if email_type not in supported_emails:
        raise TypeError(f'{email_type} is not a correct option. Available options are {", ".join(supported_emails)}')

    if email_type == 'comp':
        vid_repair = os.path.join(src_folder, 'compliance_logo.png')
        subject = f'COMPLIANCE: {subject}'

        header = '<td width=951 style="width:580.0pt;background:#500778;padding:0in 5.4pt 0in 5.4pt";height:45.35pt"><p><span style="color:white"><img src=cid:vid_repair height=51></span></td>'
        footer = f'<tr> <td width=951 valign=top style="width:713.4pt;background:#500778;padding: 0in 5.4pt 0in 5.4pt"> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><b><span style="color:white"><img border=0 width=168 height=53 src=cid:charter_logo></span></b></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><strong><span style="font-size:10.5pt;font-family:"Calibri",sans-serif;color:white">For Internal Use Only</span></strong></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><span style="font-size:8.5pt;color:white">This communication is the property of Charter Communications and is intended for internal use only. Distribution outside of the Company, in whole or part, is not permitted, except with Company permission in the course of your authorized duties. </span></p> <p style="margin-bottom:12.0pt;text-align:center"><b><span style="color:white">Video Reporting &amp; Analytics</span></b></p></td></tr>'
    
    elif email_type == 'dept':
        vid_repair = os.path.join(src_folder, 'Department_Logo.png')
        subject = f'DEPARTMENT: {subject}' 

        header = '<td width=951 style="width:580.0pt;background:#0099D8;padding:0in 5.4pt 0in 5.4pt";height:45.35pt"><p><span style="color:white"><img src=cid:vid_repair height=51></span></td>'
        footer = f'<tr> <td width=951 valign=top style="width:713.4pt;background:#0099D8;padding: 0in 5.4pt 0in 5.4pt"> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><b><span style="color:white"><img border=0 width=168 height=53 src=cid:charter_logo></span></b></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><strong><span style="font-size:10.5pt;color:white">For Internal Use Only</span></strong></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><span style="font-size:8.5pt;color:white">This communication is the property of Charter Communications and is intended for internal use only. Distribution outside of the Company, in whole or part, is not permitted, except with Company permission in the course of your authorized duties. </span></p> <p style="margin-bottom:12.0pt;text-align:center"><b><span style="color:white">Video Reporting &amp; Analytics</span></b></p></td></tr>'

    elif email_type == 'lead':
        vid_repair = os.path.join(src_folder, 'Lead_Logo.png')
        subject = f'LEAD: {subject}'

        header = '<td width=951 style="width:580.0pt;background:#FAA91A;padding:0in 5.4pt 0in 5.4pt";height:45.35pt"><p><span style="color:white"><img src=cid:vid_repair height=51></span></td>'
        footer = f'<tr> <td width=951 valign=top style="width:713.4pt;background:#FAA91A;padding: 0in 5.4pt 0in 5.4pt"> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><b><span style="color:white"><img border=0 width=168 height=53 src=cid:charter_logo></span></b></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><strong><span style="font-size:10.5pt;color:white">For Internal Use Only</span></strong></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><span style="font-size:8.5pt;color:white">This communication is the property of Charter Communications and is intended for internal use only. Distribution outside of the Company, in whole or part, is not permitted, except with Company permission in the course of your authorized duties. </span></p> <p style="margin-bottom:12.0pt;text-align:center"><b><span style="color:white">Video Reporting &amp; Analytics</span></b></p></td></tr>'

    elif email_type == 'leader':
        vid_repair = os.path.join(src_folder, 'Leader_Logo.png')
        subject = f'LEADER: {subject}'

        header = '<td width=951 style="width:580.0pt;background:#787878;padding:0in 5.4pt 0in 5.4pt";height:45.35pt"><p><span style="color:white"><img src=cid:vid_repair height=51></span></td>'
        footer = f'<tr> <td width=951 valign=top style="width:713.4pt;background:#787878;padding: 0in 5.4pt 0in 5.4pt"> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><b><span style="color:white"><img border=0 width=168 height=53 src=cid:charter_logo></span></b></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><strong><span style="font-size:10.5pt;color:white">For Internal Use Only</span></strong></p> <p style="margin-top:6.0pt;margin-right:0in; margin-bottom:6.0pt;margin-left:0in;text-align:center"><span style="font-size:8.5pt;color:white">This communication is the property of Charter Communications and is intended for internal use only. Distribution outside of the Company, in whole or part, is not permitted, except with Company permission in the course of your authorized duties. </span></p> <p style="margin-bottom:12.0pt;text-align:center"><b><span style="color:white">Video Reporting &amp; Analytics</span></b></p></td></tr>'
    

    # declaring html to build email
    
    conclusion = '<br><p><span style="color:black">If you have any questions, please reach out <a href="Network_Email_Address"><span style="font-size: 12.0pt">here</span></a><span style="font-size:12.0pt;color:black">.</span></span></p></br>'
    body = f'<table border=0 cellspacing=0 cellpadding=0 style="border-collapse:collapse"><tr>{header}</tr><tr>{explainer_html}</tr></tr>{conclusion}<p>&nbsp;</p>{footer}</table>'

    # generating email
    print('Launching Outlook')
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.To = recipients_string
    newMail.CC = cc_string
    newMail.Recipients.ResolveAll()

    # attaching banner and footer
    newMail.Attachments.Add(vid_repair)\
        .PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "vid_repair")
    newMail.Attachments.Add(logo)\
        .PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "charter_logo")

    # iterating through embed_items dict to embed all items
    if len(embed_images) > 0:
        for cid, path in embed_images.items():
            if cid is None:
                newMail.Attachments.Add(path)
            else:
                newMail.Attachments.Add(path)\
                    .PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)

    newMail.HTMLBody = body
    newMail.Display()
    # newMail.Send()