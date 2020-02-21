# -*- coding: utf-8 -*-
"""
Created on Thu May 24 13:34:26 2018
@author: Shuai

Forked, updated and changed fundamentally by: @hlra
Early 2020

Aimed to work in early 2020 until Orbis switches to new GUI in spring 2020
    (then this code would have to be adjusted/ changed/ replaced to work with the new GUI).

"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup as soup
from bs4 import SoupStrainer
from lxml import html
import numpy as np
import time, os, re
import win32com.client as win32
import pandas as pd
from statistics import mean
import threading
from win32com.client import GetObject
import pythoncom
import smtplib

## Request Orbis login and store locally
login_url = "https://oldorbis.bvdinfo.com/version-2019126/Login.serv?product=orbisneo&SetLanguage=en"
# user = input("Enter your Orbis username: ")
# pw = getpass.getpass(prompt="Enter your Orbis password: ")
user = "lu@mpifg.de"
pw = "fqmcauGPg!e4hKC"
form_data = {"user": user, "pw": pw}
# report = input("Would you like to receive a report via E-Mail every 2000 pages? If yes, type your email. If no, type NO")
report = "lu@mpifg.de"

def login_orbis():
    ## Open the Chrome test browser with selenium
    browser = webdriver.Chrome()
    browser.get(login_url)
    browser.set_page_load_timeout(900)
    ## Enter username and password
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    ## Click login button
    login_button = browser.find_element_by_id("bnLoginNeo")
    login_button.click()
    try:
        ## If there is already a running session, restart the session
        restart_button = browser.find_element_by_css_selector("#Div1 > div.container_login > div.login_enter > table > tbody > tr:nth-child(2) > td > a:nth-child(1)")
        ## by clicking the restart button
        restart_button.click()
    except:
        ## Otherwise continue in the Orbis GUI
        pass

    ## Check whether the site has loaded successfully
    #while not visible_in_time(browser, '#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_AjaxpanelTab > table > tbody > tr > td.tabFindACriterion', 0.1):
    #    ## Otherwise login again from the start
    #    time.sleep(0.1)

    ## Let Orbis list all companies without any filters
    # all_companies = browser.find_element_by_css_selector(
    #    "#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_Menu2 > li:nth-child(11)")
    # all_companies.click()
    WebDriverWait(browser, 900).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#ContentContainer1_ctl00_Content_QuickSearch1_ctl05_GoToList")))
    submit = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_QuickSearch1_ctl05_GoToList")
    submit.click()

    #sav_ser = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_MySavedSearches1_DataGridResultViewer_ctl04_Linkbutton1")
    #sav_ser.click()
    #time.sleep(900)
    browser.implicitly_wait(30)  # seconds
    return(browser)

def hard_refresh(browser,start_page):
    browser.close()

    browser = webdriver.Chrome()
    browser.get(login_url)
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    while 1:
        try:
            login_button = browser.find_element_by_class_name("ok")
            login_button.click()
            break
        except:
            browser.close()
            browser = webdriver.Chrome()
            browser.get(login_url)

    login_orbis()
    while 1:
        try:
            page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
            page_input.clear()
            page_input.send_keys(str(start_page))
            page_input.send_keys(Keys.RETURN)
            break
        except:
            continue
    refresh_stuck = visible_in_time(browser, '#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_AjaxpanelTab > table > tbody > tr > td.tabFindACriterion', 20)
    if refresh_stuck == False:
        return hard_refresh(browser,start_page)
    else:
        return browser

def visible_in_time(browser, address, time):
    try:
        WebDriverWait(browser, time).until(EC.presence_of_element_located((By.CSS_SELECTOR, address)))
        return True
    except TimeoutException:
        return False

## Function used to select all relevant base info variables. See readme file for a list of variables and more information.
def sel_base_info_vars():
    ## Open a new Chrome Browser and login to Orbis. Show the list of all companies.
    browser = login_orbis()
    time.sleep(2)
    print("New Chrome browser opened.")
    ## Click on the adjust columns button to add or remove columns
    adj_cols = browser.find_element_by_css_selector(
        "#ContentContainer1_ctl00_Content_ListHeader_ListHeaderRightButtons_ColumnBtnTD")
    adj_cols.click()
    allvars = browser.find_element_by_css_selector(
        "#ContentContainer1_ctl00_Content_VariablesSelectionCtrl_UsersSelections_UserSelectionVariableHeader > table > tbody > tr > td:nth-child(1)")
    allvars.click()
    ## Select all relevant variables (which should not exceed the maxinum number of lines in the HTML table)
    contact = browser.find_element_by_css_selector("#GCONTACT_INFO_NodeImg")
    contact.click()
    checkboxes = browser.find_elements_by_css_selector("#SubNodes_GCONTACT_INFO *.CheckboxRadioOver.middle")
    for c in checkboxes[:22]:
        try:
            c.click()
        except ElementClickInterceptedException:
            browser.switch_to.frame("frameFormatOptionDialog")
            if visible_in_time(browser,
                               '#ctl00_OptionSubViews_RepeatableFieldOption_rdFirst',
                               5) is False:
                ## Otherwise login again from the start
                browser.get(login_url)
            submit = browser.find_element_by_css_selector("#ctl00_OptionSubViews_RepeatableFieldOption_rdFirst")
            submit.click()
            submit = browser.find_element_by_css_selector("#ctl00_OptionFooterSubView_OkButton")
            submit.click()
            time.sleep(1)
            browser.switch_to.default_content()
            c.click()
        time.sleep(0.1)
    contact.click()
    print("Contact information columns added.")
    select_id = browser.find_element_by_css_selector("#GIDENT_CODE_NodeImg")
    select_id.click()
    ## Select BvD ID and LEI
    select_bvdid = browser.find_element_by_css_selector("#TreeView1\#IDENT_CODE\.BVDID\*A > img")
    select_bvdid.click()
    select_LEI = browser.find_element_by_css_selector("#TreeView1\#IDENT_CODE\.lei_LEI_Header\*A > img")
    select_LEI.click()
    select_id.click()
    print("ID columns added.")
    ## Select all size & group information variables
    select_size = browser.find_element_by_css_selector("#TreeView1\#GSIZE_GROUP_INFO > img.CheckboxRadioOver.middle")
    select_size.click()
    browser.switch_to.frame("frameFormatOptionDialog")
    if visible_in_time(browser,
                       '#ctl00_OptionFooterSubView_OkButton',
                       5) is False:
        ## Otherwise login again from the start
        browser.get(login_url)
    submit = browser.find_element_by_css_selector("#ctl00_OptionFooterSubView_OkButton")
    submit.click()
    browser.switch_to.default_content()
    goback_size = browser.find_element_by_css_selector("#GSIZE_GROUP_INFO_NodeImg")
    goback_size.click()
    uncheck_dup = browser.find_element_by_css_selector("#TreeView1\#SIZE_GROUP_INFO\.NACE2_MAIN_SECTION\*A > img")
    uncheck_dup.click()
    uncheck_dup = browser.find_element_by_css_selector("#TreeView1\#SIZE_GROUP_INFO\.MAJOR_SECTOR\*A > img")
    uncheck_dup.click()
    print("Group and size information columns added.")
    ## Select all industry and activities variables
    open_tab1 = browser.find_element_by_css_selector("#GACTIVITIES_NodeImg")
    open_tab1.click()
    open_tab2 = browser.find_element_by_css_selector("#ACTIVITIES\*ACTIVITIES\.TITLE01\*A_NodeImg")
    open_tab2.click()
    open_tab3 = browser.find_element_by_css_selector("#GPEERGROUP_NodeImg")
    open_tab3.click()
    open_tab4 = browser.find_element_by_css_selector("#GOVERVIEW_NodeImg")
    open_tab4.click()
    checkboxes = browser.find_elements_by_css_selector("#SubNodes_GACTIVITIES *.CheckboxRadioOver.middle")
    for c in checkboxes[:7]:
        try:
            c.click()
        except ElementClickInterceptedException:
            browser.switch_to.frame("frameFormatOptionDialog")
            if visible_in_time(browser,
                               '#ctl00_OptionSubViews_ACTIVITIES-RepeatableGroupFieldOption_rdFirst',
                               5) is False:
                ## Otherwise login again from the start
                browser.get(login_url)
            submit = browser.find_element_by_css_selector("#ctl00_OptionSubViews_ACTIVITIES-RepeatableGroupFieldOption_rdFirst")
            submit.click()
            submit = browser.find_element_by_css_selector("#ctl00_OptionFooterSubView_OkButton")
            submit.click()
            time.sleep(1)
            browser.switch_to.default_content()
            c.click()
        time.sleep(0.1)
        try:
            c = browser.find_element_by_css_selector("#TreeView1\#ACTIVITIES\*ACTIVITIES\.TITLE01\*A > img.CheckboxRadioOver.middle")
            c.click()
            c = browser.find_element_by_css_selector("#TreeView1\#GPEERGROUP > img.CheckboxRadioOver.middle")
            c.click()
            c = browser.find_element_by_css_selector("#TreeView1\#GOVERVIEW > img.CheckboxRadioOver.middle")
            c.click()
        except ElementClickInterceptedException:
            browser.switch_to.frame("frameFormatOptionDialog")
            if visible_in_time(browser,
                               '#ctl00_OptionSubViews_ACTIVITIES-RepeatableGroupFieldOption_rdFirst',
                               5) is False:
                ## Otherwise login again from the start
                browser.get(login_url)
            submit = browser.find_element_by_css_selector("#ctl00_OptionSubViews_ACTIVITIES-RepeatableGroupFieldOption_rdFirst")
            submit.click()
            submit = browser.find_element_by_css_selector("#ctl00_OptionFooterSubView_OkButton")
            submit.click()
            time.sleep(1)
            browser.switch_to.default_content()
            c.click()
            continue
    open_tab4.click()
    open_tab3.click()
    open_tab2.click()
    open_tab1.click()
    print("Industry and activities columns added.")
    ## Go to the list of all companies with this view
    submit = browser.find_element_by_css_selector(
        "#ContentContainer1_ctl00_Content_SaveFormat_OkButton")
    submit.click()
    print("Base variable columns successfully selected.")
    return(browser)

def scrape_table(browser):
    # html retrieving
    innerHTML = browser.execute_script("return document.body.innerHTML")
    page_soup = soup(innerHTML, "html.parser")
    columns = page_soup.find(id="ContentContainer1_ctl00_Content_ListCtrl1_LB1_VHDRRW").find_next('tr').find_next('tr')
    label_info = columns.find_all('td', class_=re.compile('.*mclbOvH mclbCP'))
    column_names = []
    for x in label_info:
        column_names.append(
            x.getText())
    column_names = ['company_name'] + column_names
    column_num = len(column_names)
    company_data = pd.DataFrame()
    company_names = []
    for x in column_names:
        company_data[x] = []

    ## page_num = Current page number taken from the input field in Orbis
    page_num = page_soup.find(attrs= {'class': 'form_textarea_current_page'})['value']
    page_num = int(page_num)
    ## The total number of companies shown in the search window. This needs to be set to 100 manually in the Orbis settings.
    per_page = 100
    ## total_ companies = The total number of companies taken from the corresponding field in Orbis
    total_companies = int(page_soup.find_all(attrs= {'class': 'label_3 WHR WVT'})[1].text.replace(',', ''))
    ## total_ companies = The total number of pages taken from the corresponding field in Orbis
    grand_total_pages = total_companies // per_page + 1  # Number of pages of data to retrieve
    total_pages = grand_total_pages
    ## Pages scraped per round
    per_round = 20
    ## The total number of pages scraped
    pages = 0

    print("Start retrieving data of Page {1}!")

    strainer_a = SoupStrainer('a', {'data-action': "reporttransfer"})
    strainer_td = SoupStrainer('td', {'class': 'scroll-data'})
    strainer_td = SoupStrainer('td', {'class': 'scroll-data'})
    strainer_input = SoupStrainer('input', {'title': 'Number of page'})
    strainer_tbody = SoupStrainer(id='resultsTable')

    innerHTML = []
    page_done = 0
    while page_num <= total_pages:
        stopwatch = time.time()
        company_data = company_data.drop("company_name", axis=1)
        innerHTML = browser.execute_script("return document.body.innerHTML")
        # tree = html.fromstring(innerHTML)
        page_num = int(page_soup.find(attrs= {'class': 'form_textarea_current_page'})['value'])
        while pages < per_round:
            print("Page {0} opened!".format(page_done+1))
            page_soup = soup(innerHTML, "lxml").select_one('#master_content')

            ## If company names not generated yet (because it is the first page) create empty list and add first 100 companys
            if company_names == []:
                company_names = [x.text for x in page_soup.select(
                    '#ContentContainer1_ctl00_Content_ListCtrl1_LB1_FDTBL * > a[href="#"]')]
            else:
                ## ..otherwise add to existing company names
                company_names += [x.text for x in page_soup.select(
                    '#ContentContainer1_ctl00_Content_ListCtrl1_LB1_FDTBL * > a[href="#"]')]
            ## Scrape table for current page
            data_points = page_soup.select(
                'td[class*="mclbOvH"][class*="resultsItems"]')
            if page_num == total_pages:
                num_on_page = total_companies - (page_done + pages - per_round) * per_page
            else:
                num_on_page = per_page
            data = [x.text.replace("\xa0", "") for x in data_points]
            data = np.array_split(data, per_page)
            company_data = pd.concat([company_data, pd.DataFrame(data, columns=column_names[1:])], sort=False)

            print("Page {0} finished!".format(page_done+1))
            next_page = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_ListHeader_ListNavigation_NextPage")
            next_page.click()

            page_done += 1
            new_num = 1
            while new_num != (page_done*100)+1:
                time.sleep(0.5)
                new_num = int(browser.find_element_by_css_selector(
                    "#ContentContainer1_ctl00_Content_ListCtrl1_LB1_FDTBL > tbody > tr:nth-child(2) > td:nth-child(1)").text.replace(
                    ".", ""))
            pages += 1
            page_num += 1
    # Rolling after saving the data
    company_data.insert(0, "company_name", company_names)
    if pages == per_round:
        company_data.to_csv('All_columns-{0}.txt', mode='a', sep='|', index=False)
    else:
        company_data.to_csv('All_columns-{0}.txt', mode='a', sep='|', index=False, header=False)

    print('{0} to {1} pages output! Time cost:{2:.2f}s'.format(max(pages - per_round + 1,0), page_num,
                                                                   time.time() - stopwatch))
    # Report each 20000 pages
    try:

        if pages % 2000 == 0 & (report!="NO") :
            avg_time = (time.time() - start_time) / (pages - start_page) * 1000
            round_time_spent = (time.time() - big_round_time) / (pages - max(start_page, pages - 2000)) * 1000
            big_round_time = time.time()
            fastest_time = 0

            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = report
            mail.Subject = 'Orbis Data Scraping System update'
            mail.Body = """System is running.
                            Current Time: {2}.
                            Start Time: {3}.
                            Current Page: {1}. 
                            Total Pages: {4}.
                            Average Time per 2000 pages: {5:.2f}s.
                            Time spent on the last 2000 pages: {7:.2f}s
                            Number of Hard Refresh in the last 2000 pages: {8}.
                            Approximately another {6:.2f} hours to finish rest of data.
                            Reported by Orbis Data Scraping System
                """.format(pages, time.ctime(), start_datetime, grand_total_pages, avg_time,
                           (grand_total_pages - pages) / 1000 * avg_time / 3600, round_time_spent,
                           hard_refresh_times)
            mail.Display()
            mail.Save()
            mail.Close(0)
    except:
        pass
    # if round_time_spent > avg_time * 2:
    #    browser = hard_refresh(browser, pages)
    print('Successful output to csv file!')

def export_all(browser):

    exp = browser.find_element_by_css_selector(
        "#ContentContainer1_ctl00_Content_ListHeader_ListHeaderRightButtons_AddRemoveColumns")
    exp.click()
    exp = browser.find_element_by_css_selector("#setAsCurrent_listformatcompanies0006")
    exp.click()

    round = 7
    min_val = range(1, 339053, 2400)

    for x in min_val:
        # export button
        exp = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_ListHeader_ListHeaderRightButtons_ExportButtons_ExportButton")
        exp.click()

        while len(browser.window_handles) <= 1:
            time.sleep(0.1)
        browser.switch_to.window(browser.window_handles[1])

        exp = browser.find_element_by_css_selector("#chRepeatSingleItem")
        exp.click()
        exp = browser.find_element_by_css_selector(
            "#aspnetForm > table:nth-child(39) > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(5) > td > table > tbody > tr > td:nth-child(2) > input[type=text]")
        exp.click()
        min = browser.find_element_by_css_selector("#aspnetForm > table:nth-child(39) > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(5) > td > table > tbody > tr > td:nth-child(2) > input[type=text]")
        min.send_keys(x)
        maxdown = 2400
        max = browser.find_element_by_css_selector("#aspnetForm > table:nth-child(39) > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(5) > td > table > tbody > tr > td.Label > input[type=text]")
        max.send_keys(x+maxdown-1)
        name = browser.find_element_by_css_selector("#ctl00_ContentContainer1_ctl00_LowerContent_Formatexportoptions1_ExportDisplayName")
        name.clear()
        name.send_keys("DM"+str(round))
        ok = browser.find_element_by_css_selector("#ctl00_ContentContainer1_ctl00_ButtonsContent_ExportOptionsBottomButtons_OkLabel")
        ok.click()
        if round == 7:
            ok = browser.find_element_by_css_selector("#CloseLink > img")
            ok.click()
        else:
            if round%10 !=0 or round-1%10 ==0:
                mail = browser.find_element_by_css_selector(
                    "#SendEmailCheck")
                mail.click()
            ok = browser.find_element_by_css_selector("#RegisterTimeoutLink > img")
            ok.click()
        print("Download started for companies "+str(x) + " to " + str(x+2400-1) +".")
        round += 1
        #ok = browser.find_element_by_css_selector("#CloseLink > img")
        #ok.click()
        browser.switch_to.window(browser.window_handles[0])

def download_all(browser):
    # Snippet taken from: https://medium.com/@moungpeter/how-to-automate-downloading-files-using-python-selenium-and-headless-chrome-9014f0cdd196
    # instantiate a chrome options object so you can set the size and headless preference
    # some of these chrome options might be uncessary but I just used a boilerplate
    # change the <path_to_download_default_directory> to whatever your default download folder is located

    download_directory = r"S:\Meine Bibliotheken\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\Ownership\History\\"

    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--verbose')
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
    })
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-software-rasterizer')

    ## Open the Chrome test browser with selenium
    browser = webdriver.Chrome(chrome_options=chrome_options)
    browser.get(login_url)
    browser.set_page_load_timeout(900)
    ## Enter username and password
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    ## Click login button
    login_button = browser.find_element_by_id("bnLoginNeo")
    login_button.click()
    try:
        ## If there is already a running session, restart the session
        restart_button = browser.find_element_by_css_selector("#Div1 > div.container_login > div.login_enter > table > tbody > tr:nth-child(2) > td > a:nth-child(1)")
        ## by clicking the restart button
        restart_button.click()
    except:
        ## Otherwise continue in the Orbis GUI
        pass

    ## Check whether the site has loaded successfully
    #while not visible_in_time(browser, '#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_AjaxpanelTab > table > tbody > tr > td.tabFindACriterion', 0.1):
    #    ## Otherwise login again from the start
    #    time.sleep(0.1)

    ## Let Orbis list all companies without any filters
    # all_companies = browser.find_element_by_css_selector(
    #    "#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_Menu2 > li:nth-child(11)")
    # all_companies.click()
    WebDriverWait(browser, 900).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#ContentContainer1_ctl00_Content_QuickSearch1_ctl05_GoToList")))
    submit = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_QuickSearch1_ctl05_GoToList")
    submit.click()

    #sav_ser = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_QuickSearch1_ctl02_MySavedSearches1_DataGridResultViewer_ctl04_Linkbutton1")
    #sav_ser.click()
    #time.sleep(900)
    browser.implicitly_wait(30)  # seconds

    set = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Header_ctl00_ctl06_TopMenu > li > table > tbody > tr > td:nth-child(2)")
    set.click()
    setexp = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Header_ctl00_ctl06_TopMenu > li > ul > li:nth-child(3)")
    move = ActionChains(browser).move_to_element(setexp)
    time.sleep(0.5)
    move.perform()
    set = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Header_ctl00_ctl06_TopMenu > li > ul > li:nth-child(3) > ul > li.first")
    set.click()

    num = ""
    numdown = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_ctl00_ExportsDataGrid > tbody > tr.label_3 > td > table > tbody > tr > td:nth-child(1)").text
    num = int([num + s for s in numdown.split() if s.isdigit()][0])

    download = 0

    while download < num:

        while browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_ctl00_ExportsDataGrid > tbody > tr:nth-child(3) > td:nth-child(3) > table > tbody > tr > td > table > tbody > tr > td.headeven > div").text != "Done":
            time.sleep(1)
            print("File not downloaded yet... Waiting.")
        print("Starting first download..")
        firstdown = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_ctl00_ExportsDataGrid > tbody > tr:nth-child(3) > td:nth-child(2) > table > tbody > tr > td > a")
        firstdown.click()

        filename = browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_ctl00_ExportsDataGrid > tbody > tr:nth-child(3) > td:nth-child(2) > table > tbody > tr > td > a").text
        filepath = download_directory+filename+".csv"

        x = 1
        while x < 130:
            while not os.path.exists(filepath):
                time.sleep(5)
                x += 1
                continue
            if os.path.exists(filepath):
                break
        if not os.path.exists(filepath):
            print("File " + filename + "not downloaded... Process aborted.")
            download = num
            break
        elif os.path.exists(filepath):
            firstdown = browser.find_element_by_css_selector(
                "# ContentContainer1_ctl00_Content_ctl00_ExportsDataGrid_ctl03_DeleteButton > img")
            firstdown.click()
            download += 1


def pytimeout(browser):

    pythoncom.CoInitialize()

    WMI = GetObject('winmgmts:')
    processes = WMI.InstancesOf('Win32_Process')

    for p in WMI.ExecQuery('select * from Win32_Process where Name="WerFault.exe"'):
        # print("Killing PID:", p.Properties_('ProcessId').Value)
        os.system("taskkill /f /pid " + str(p.Properties_('ProcessId').Value))

    try:
        browser.quit()
    except:
        pass

    WMI = GetObject('winmgmts:')
    processes = WMI.InstancesOf('Win32_Process')

    for p in WMI.ExecQuery('select * from Win32_Process where Name="chromedriver.exe"'):
        # print("Killing PID:", p.Properties_('ProcessId').Value)
        os.system("taskkill /f /pid " + str(p.Properties_('ProcessId').Value))

    WMI = GetObject('winmgmts:')
    processes = WMI.InstancesOf('Win32_Process')

    for p in WMI.ExecQuery('select * from Win32_Process where Name="chrome.exe"'):
        # print("Killing PID:", p.Properties_('ProcessId').Value)
        os.system("taskkill /f /pid " + str(p.Properties_('ProcessId').Value))

    print("The browser was closed because it took over 2 minutes to scrape the page. "
          "If this happens five times in a row with the same company, this company is skipped and its ID printed at the end of the process.")

def scrape_ownership_report(browser, startpage):
    try:
        # navigate to company reports of sample from search result list
        first_comp = browser.find_element_by_css_selector(
            "#ContentContainer1_ctl00_Content_ListCtrl1_LB1_FDTBL > tbody > tr:nth-child(2) > td.cellBlank.nameResultItems.mclbOvH.mclbEl > a")
        first_comp.click()
        while not visible_in_time(browser, "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_AddRemoveSections", 0.1):
            time.sleep(0.1)
        # navigate to selection of sections in company reports
        sel_sections = browser.find_element_by_css_selector(
            "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_AddRemoveSections")
        sel_sections.click()
        # only select historical ownership data
        sel_sections = browser.find_element_by_css_selector(
            "#TreeView1\#GROUPOWNERSHIP > a")
        sel_sections.click()
        time.sleep(0.1)
        sel_sections = browser.find_element_by_css_selector(
            "#TreeView1\#GROUPSHAREHOLDERS > a")
        sel_sections.click()
        time.sleep(0.1)
        sel_sections = browser.find_element_by_css_selector(
            "#checkBoxSelection")
        sel_sections.click()
        time.sleep(0.1)
        sel_sections = browser.find_element_by_css_selector(
            "#TreeView1\#SHAREHOLDERSHISTORY > a")
        sel_sections.click()
        time.sleep(0.1)
        sel_sections = browser.find_element_by_css_selector(
            "#ContentContainer_ctl00_Content_ReportFormatCustomizerControl_SectionOptionFooter_OkButton > div")
        sel_sections.click()
        # Select BvD ID and LEI Identifier as additional columns
        sel_sections = browser.find_element_by_css_selector(
            "#m_ContentControl_ContentContainer1_ctl00_Content_Header_SHAREHOLDERSHISTORY_m_RootContainer > tbody > tr > td > table > tbody > tr > td.jQueryTarget.section_title.WNW.WHC.WVT > span > a")
        sel_sections.click()
        sel_sections = browser.find_element_by_css_selector(
            "#PossibleSuiteId")
        sel_sections.click()
        sel_sections = browser.find_element_by_css_selector(
            "#PossibleLEI")
        sel_sections.click()
        sel_sections = browser.find_element_by_css_selector(
            "#ContentContainer_ctl00_Content_ColumnsEditionFooter_OkButton > div")
        sel_sections.click()

        ## Pages scraped per round
        per_round = 100
        if startpage == 1:
            round = 1
            startpage = 1
        else:
            round = int(startpage/per_round)
            startpage = round*per_round
            ## The total number of pages scraped
        page_done = startpage-1

        if startpage > 1:
            visible_in_time(browser,
                            "#SeqNrlbl",
                            2)
            sp = browser.find_element_by_css_selector(
                "#SeqNrlbl")
            browser.execute_script("arguments[0].click()", sp)
            browser.execute_script("arguments[0].setAttribute('value', '" + str(startpage) +"')", sp)
            sp.send_keys(Keys.ENTER)


        quartals = []
        for year in [2018]:
        #for year in reversed(range(2012,2020)):
            #for mon in ["12", "09", "06", "03"]:
            for mon in["06"]:
                quartals.append(mon+"/"+str(year))

        ## There are cases for which no quartal information is available.. Store and save this data
        quartalmissflag = []

        total_companies = int(browser.find_element_by_css_selector("#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_NavigationTD").text.replace("of ", "").replace(",", ""))


        ## Repeat for the whole sample
        while page_done < total_companies:
            company_data = pd.DataFrame()
            pages = 0
            stopwatch = time.time()
            ## Repeat for per_round pages per round
            while pages < per_round:
                print("Round {0}, retrieving data of company {1} started at {2}.".format(round, page_done+1, time.ctime()))
                case = 0
                ## Check whether the company is in the error list and must be skipped
                if page_done+1 in missing_ids:
                    next_page = browser.find_element_by_css_selector(
                        "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_NavigationTD > a:nth-child(5) > img")
                    next_page.click()
                    page_done += 1
                    pages += 1
                    continue
                ## There are four cases of ownership situations:
                ## Case 1: The company does not have any previous or current shareholders
                try:
                    while not visible_in_time(browser, "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_NavigationTD > a:nth-child(5) > img", 0.1):
                        time.sleep(0.1)
                    browser.find_element_by_css_selector(
                        "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholdersHistoryNoDataContainer > tbody > tr:nth-child(1) > td:nth-child(2)").text == "There is no shareholder information available for this entity."
                    print("Page {0}: Skipped because of case 1: The company does not have any previous or current shareholders!".format(page_done + 1))
                    try:
                        visible_in_time(browser,
                                        "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate",
                                        2)
                        quart = browser.find_element_by_css_selector(
                            "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate")
                        quart.clear()
                    except:
                        pass

                    try:
                        visible_in_time(browser,
                                        "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_NavigationTD > a:nth-child(5) > img",
                                        2)
                        quart = browser.find_element_by_css_selector(
                            "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate")
                        quart.send_keys(quartals[0])
                        time.sleep(0.1)
                        quart = browser.find_element_by_css_selector(
                            "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_NumHistoryLines")
                        quart.send_keys("a")
                    except:
                        ## When there is no field for quartal, the company does not have any shareholders:
                        ## e.g. "The company is a foreign company of ..."

                        ## There are also cases in which there are simply no fields for quarter which means the timing is unclear
                        print(
                            "Page {0}: Skipped because: The company does not have any previous or current shareholders! Probably because it is a foreign firm of another firm.".format(
                                page_done + 1))
                        pass
                    next_page = browser.find_element_by_css_selector(
                        "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_NavigationTD > a:nth-child(5) > img")
                    next_page.click()
                    page_done += 1
                    pages += 1
                    continue
                except:
                    # repeat for all quartals for 10 years (2009-2019)
                    for qu in quartals:
                        # For very large ownership reports (e.g. Norwegian companies with many listed individual shareholders) chrome crashes when scraping the data.
                        # Therefore a timer is built in which skips a company after five tries and saves its ID to a list for later reference.
                        try:
                            ## If the timer process was not terminated in previous round: terminate now
                            t.cancel()
                        except:
                            pass
                        t = threading.Timer(120.0, lambda: pytimeout(browser))
                        t.start()
                        ## If quartal is before the first quartal scrape all historical data
                        if qu == quartals[-1]:
                            try:
                                visible_in_time(browser,
                                                "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate",
                                                5)
                                quart = browser.find_element_by_css_selector(
                                    "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate")
                                quart.clear()
                            except:
                                pass

                            try:
                                visible_in_time(browser,
                                                "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate",
                                                5)
                                quart = browser.find_element_by_css_selector(
                                    "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate")
                                quart.send_keys(qu)
                                time.sleep(0.1)
                                quart = browser.find_element_by_css_selector(
                                    "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_NumHistoryLines")
                                # quart.send_keys("a")
                                quart.send_keys("0")
                            except:
                                quartalmissflag.append(page_done+1)
                                pd.DataFrame(quartalmissflag).to_csv("quartalmissflag.csv", sep=';', index=False, mode="a")
                                pass
                        ## Otherwise only the last entry
                        else:
                            try:
                                visible_in_time(browser,
                                                "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate",
                                                5)
                                quart = browser.find_element_by_css_selector(
                                    "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate")
                                quart.clear()
                            except:
                                pass

                            try:
                                visible_in_time(browser,
                                                "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate",
                                                5)
                            except:
                                pass
                            quart = browser.find_element_by_css_selector(
                                "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_LastHistoryDate")
                            quart.send_keys(qu)
                            time.sleep(0.1)
                            quart = browser.find_element_by_css_selector(
                                "#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholderSettings_NumHistoryLines")
                            quart.send_keys("0")
                        ## Case 2: The company does not have any previous shareholders but has current shareholders:
                        ## Unfold all current shareholders
                        try:
                            # Check if there is the "Current shareholders" table
                            browser.find_element_by_xpath("//td[text()='Current shareholders:']")
                            # print("There is a current shareholder table on this page.")
                            try:
                                ## Try if there are more than 50 current shareholders, and if yes, click the unfold button
                                all_sh = browser.find_element_by_xpath("//a[text()='View all current shareholders']")
                                all_sh.click()
                            except:
                                # print("There are less than 50 current shareholders on this page.")
                                ## If there are less than 50, just pass
                                pass
                            ## Assign case 2: There is a current shareholder table
                            case = 2
                        except:
                            # print("No current shareholder table found on page.")
                            case = 0
                            # If there is no such table, assign case 2 and just pass
                            pass

                        ## Case 3: The company does not have any current shareholders but has previous shareholders:
                        ## Unfold all previous shareholders
                        try:
                            # Check if there is the "Previous shareholders" table
                            browser.find_element_by_xpath("//td[text()='Previous shareholders:']")
                            # print("There is a current shareholder table on this page.")
                            try:
                                ## Try if there are more than 50 current shareholders, and if yes, click the unfold button
                                all_sh = browser.find_element_by_xpath("//a[text()='View all previous shareholders']")
                                all_sh.click()
                            except:
                                # print("Less than 50 previous shareholders found on this page.")
                                ## If there are less than 50, just pass
                                pass
                            if case == 2:
                                # print("Case 4: There are current as well as previous shareholders on this page.")
                                ## If there is also a current shareholder table (case = 2), assign case 4: both tables are present
                                case = 4
                        except:
                            # If there is no such table, assign case 3 and just pass
                            if case == 2:
                                case = 3
                                # print("Case 3: No previous shareholders found on this page but current shareholders.")
                            elif case == 0:
                                print(
                                    "Page {0}: Skipped because of case 1: The company does not have any previous or current shareholders!".format(
                                        page_done + 1))
                                # print("Timer stopped at {0}.".format(time.ctime()))
                                t.cancel()
                                break
                            pass

                        # html retrieving
                        innerHTML = browser.execute_script("return document.body.innerHTML")
                        soups = soup(innerHTML, "html.parser")


                        ## If case is 2 or 4: scrape current shareholder table
                        if case == 4 or case == 2:
                            ## scrape current shareholder table
                            # print("Page {0} opened! Which has a current shareholder table to be scraped.".format(page_done + 1))
                            columns = soups.find("tr", attrs={"class", "Header"})
                            label_info = columns.find_all('td')
                            column_names = []

                            for x in label_info:
                                if x.getText().replace('\xa0', "").replace("\n", "") == "":
                                    continue
                                column_names.append(
                                    x.getText().replace('\xa0', "").replace("\n", ""))
                            column_names = ['no'] + column_names
                            column_num = len(column_names)
                            company_names = []
                            for x in column_names:
                                if company_data.empty: company_data[x] = []

                            innerHTML = []
                            innerHTML = browser.execute_script("return document.body.innerHTML")

                            page_soup = soup(innerHTML, "lxml").select_one(
                                '#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholdersHistoryDataContainer')

                            ## Read in all td nodes from current shareholder table by selecting a heading that is between the two tables
                            data = [x.text.replace('\xa0', "").replace("\n", "") for x in
                                    page_soup.select("td.label_8.WVT")[1].
                                        find_all_previous('td', class_=re.compile('WVT'))]

                            data.reverse()
                            ## Cut all superflous cells scraped before beginning of the table:
                            pattern = re.compile("\d{1,2}[.]$")
                            num = 0
                            for x in data:
                                if not pattern.match(x):
                                    num += 1
                                else:
                                    break
                            data = data[num:]
                            # On the basis of regular expressions, try to structure the data suitible for data analysis (by repeating values per row)
                            length = len(data) - 1
                            for x in range(0, len(data) - 1):
                                if pattern.match(data[x]):
                                    num = 1
                                    pattern2 = re.compile("\d{1,2}[.]")
                                    while not (pattern2.match(data[x + 4]) or data[x + 4] == "-" or data[
                                        x + 4] == "n.a.") and num <= 4:
                                        if data[x + num] == "": del data[x + num]
                                        num += 1
                                    length = len(data)
                                    if length >= x + 24 and data[x + 13] == "" and (data[x + 1] == ""):
                                        if data[x + 13] == "": data[x + 13] = data[x]
                                        if data[x + 14] == "": data[x + 14] = data[x + 2]
                                        if data[x + 15] == "": data[x + 15] = data[x + 3]
                                        if data[x + 16] == "": data[x + 16] = data[x + 4]
                                        if data[x + 21] == "": data[x + 21] = data[x + 9]
                                        if data[x + 22] == "": data[x + 22] = data[x + 10]
                                        if data[x + 23] == "": data[x + 23] = data[x + 11]
                                        if data[x + 24] == "": data[x + 24] = data[x + 12]
                                        del data[x + 1]
                                        length = len(data)
                                    elif length >= x + 23 and data[x + 12] == "" and not (data[x + 1] == ""):
                                        if data[x + 12] == "": data[x + 12] = data[x]
                                        if data[x + 13] == "": data[x + 13] = data[x + 1]
                                        if data[x + 14] == "": data[x + 14] = data[x + 2]
                                        if data[x + 15] == "": data[x + 15] = data[x + 3]
                                        if data[x + 20] == "": data[x + 20] = data[x + 8]
                                        if data[x + 21] == "": data[x + 21] = data[x + 9]
                                        if data[x + 22] == "": data[x + 22] = data[x + 10]
                                        if data[x + 23] == "": data[x + 23] = data[x + 11]
                                if x == length - 1:
                                    break

                            ## Cut superfluous cells after the end of the table
                            data.reverse()
                            pattern = re.compile("\d{1,2}[.]$")
                            num = 0
                            for x in data:
                                if not pattern.match(x):
                                    num += 1
                                else:
                                    break
                            num = num - 11
                            data = data[num:]
                            data.reverse()

                            data = np.array_split(data, len(data) / column_num)

                            ## INSERT BVD AND COMPANY_NAME VARIABLE HERE!
                            try:
                                bvd = browser.find_element_by_xpath(
                                    "//td[text()='BvD ID number']/following-sibling::td[1]").text
                            except:
                                print("BvD ID on page {0} could not be scraped...".format(page_done + 1))
                                pass
                            try:
                                comp_name = browser.find_element_by_css_selector(
                                    "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Section_TITLE_CompanyName > tbody > tr > td.reportTitle.WVM").text
                            except:
                                print(
                                    "Company name on page {0} of company reports could not be scraped...".format(page_done + 1))
                                pass

                        if case == 4 or case == 3:

                            if case == 3:
                                innerHTML = []
                                innerHTML = browser.execute_script("return document.body.innerHTML")

                                page_soup = soup(innerHTML, "lxml").select_one(
                                    '#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholdersHistoryDataContainer')

                                columns = page_soup.find("tr", attrs={"class", "Header"})
                                label_info = columns.find_all('td')
                                column_names = []

                                for x in label_info:
                                    if x.getText().replace('\xa0', "").replace("\n", "") == "":
                                        continue
                                    column_names.append(
                                        x.getText().replace('\xa0', "").replace("\n", ""))
                                column_names = ['no'] + column_names
                                column_num = len(column_names)
                                company_names = []
                                for x in column_names:
                                    if company_data.empty: company_data[x] = []

                                page_soup = soup(innerHTML, "lxml").select_one(
                                    '#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholdersHistoryDataContainer')

                                ## Read in all td nodes from current shareholder table by selecting a heading that is between the two tables
                                data = [x.text.replace('\xa0', "").replace("\n", "") for x in
                                        page_soup.select("td.label_8.WVT")[0].
                                            find_all_next('td', class_=re.compile('WVT'))]

                            elif case == 4:
                                innerHTML = []
                                innerHTML = browser.execute_script("return document.body.innerHTML")

                                page_soup = soup(innerHTML, "lxml").select_one(
                                    '#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholdersHistoryDataContainer')

                                columns = page_soup.find("tr", attrs={"class", "Header"})
                                label_info = columns.find_all('td')
                                column_names = []

                                for x in label_info:
                                    if x.getText().replace('\xa0', "").replace("\n", "") == "":
                                        continue
                                    column_names.append(
                                        x.getText().replace('\xa0', "").replace("\n", ""))
                                column_names = ['no'] + column_names
                                column_num = len(column_names)
                                company_names = []
                                for x in column_names:
                                    if company_data.empty: company_data[x] = []

                                page_soup = soup(innerHTML, "lxml").select_one(
                                    '#m_ContentControl_ContentContainer1_ctl00_Content_Section_SHAREHOLDERSHISTORY_ShareholdersHistoryDataContainer')

                                ## Read in all td nodes from current shareholder table by selecting a heading that is between the two tables
                                data = [x.text.replace('\xa0', "").replace("\n", "") for x in
                                        page_soup.select("td.label_8.WVT")[1].
                                            find_all_next('td', class_=re.compile('WVT'))]

                            ## Cut all superflous cells scraped before beginning of the table:
                            pattern = re.compile("\d{1,2}[.]$")
                            num = 0
                            for x in data:
                                if not pattern.match(x):
                                    num += 1
                                else:
                                    break
                            data = data[num:]
                            # On the basis of regular expressions, try to structure the data suitible for data analysis (by repeating values per row)
                            length = len(data) - 1
                            for x in range(0, len(data) - 1):
                                if pattern.match(data[x]):
                                    num = 1
                                    pattern2 = re.compile("\d{1,2}[.]")
                                    while not (pattern2.match(data[x + 4]) or data[x + 4] == "-" or data[
                                        x + 4] == "n.a.") and num <= 4:
                                        if data[x + num] == "": del data[x + num]
                                        num += 1
                                    length = len(data)
                                    if length >= x + 24 and data[x + 13] == "" and (data[x + 1] == ""):
                                        if data[x + 13] == "": data[x + 13] = data[x]
                                        if data[x + 14] == "": data[x + 14] = data[x + 2]
                                        if data[x + 15] == "": data[x + 15] = data[x + 3]
                                        if data[x + 16] == "": data[x + 16] = data[x + 4]
                                        if data[x + 21] == "": data[x + 21] = data[x + 9]
                                        if data[x + 22] == "": data[x + 22] = data[x + 10]
                                        if data[x + 23] == "": data[x + 23] = data[x + 11]
                                        if data[x + 24] == "": data[x + 24] = data[x + 12]
                                        del data[x + 1]
                                        length = len(data)
                                    elif length >= x + 23 and data[x + 12] == "" and not (data[x + 1] == ""):
                                        if data[x + 12] == "": data[x + 12] = data[x]
                                        if data[x + 13] == "": data[x + 13] = data[x + 1]
                                        if data[x + 14] == "": data[x + 14] = data[x + 2]
                                        if data[x + 15] == "": data[x + 15] = data[x + 3]
                                        if data[x + 20] == "": data[x + 20] = data[x + 8]
                                        if data[x + 21] == "": data[x + 21] = data[x + 9]
                                        if data[x + 22] == "": data[x + 22] = data[x + 10]
                                        if data[x + 23] == "": data[x + 23] = data[x + 11]
                                if x == length - 1:
                                    break

                            ## Cut superfluous cells after the end of the table
                            data.reverse()
                            pattern = re.compile("\d{1,2}[.]$")
                            num = 0
                            for x in data:
                                if not pattern.match(x):
                                    num += 1
                                else:
                                    break
                            num = num - 11
                            data = data[num:]
                            data.reverse()

                            data = np.array_split(data, len(data) / column_num)

                        # print("Timer stopped at {0}.".format(time.ctime()))
                        t.cancel()
                        # Generate dataframe data for new quartal to be added to data from previous quartals
                        newdata = pd.DataFrame(data, columns=column_names)

                        ## INSERT BVD AND COMPANY_NAME VARIABLE HERE!
                        try:
                            bvd = browser.find_element_by_xpath(
                                "//td[text()='BvD ID number']/following-sibling::td[1]").text
                            newdata["comp_bvd"] = bvd
                        except:
                            # print("BvD ID on page {0} could not be scraped...".format(page_done+1))
                            break
                        try:
                            comp_name = browser.find_element_by_css_selector(
                                "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Section_TITLE_CompanyName > tbody > tr > td.reportTitle.WVM").text
                            newdata["comp_name"] = comp_name
                        except:
                            # print(
                            #    "Company name on page {0} of company reports could not be scraped...".format(page_done+1))
                            break
                        newdata["quartal"] = qu
                        company_data = pd.concat([company_data, newdata], sort=False)
                        print("Company {0} successfully scraped for quartal {1}".format(page_done+1, qu))
                try:
                    company_data["no"] = (company_data.groupby(['comp_name', 'Shareholder name']).ngroup()+1).astype(
                                                'category')
                except:
                    # print("Could not add company name for company {0}..".format(page_done + 1))
                    pass
                next_page = browser.find_element_by_css_selector(
                            "#m_ContentControl_ContentContainer1_ctl00_FixedContent_Headerbarreport1_NavigationTD > a:nth-child(5) > img")
                next_page.click()
                page_done += 1
                pages += 1

            # Rolling after saving the data
            company_data.to_csv('S:/Meine Bibliotheken/Meine Bibliothek/Dissertation/Data/ORBIS/Scraping/Scraped_Data/Ownership/ownrep-{0}.csv'.format(round), mode='a', sep='|', index=False)

            round_time_spent = time.time() - stopwatch
            #('{0} to {1} pages output! Time cost:{2:.2f}s'.format(page_done - per_round, round*per_round,
            #                                                                   round_time_spent))

            try:
                time_sample.append(round_time_spent)
                avg_time = mean(time_sample)
            except NameError:
                time_sample = []
                avg_time = round_time_spent
                pass
            #if round > 1:
            #    avg_time = (((round-1)*avg_time)+round_time_spent)/round

            print("""Exported round {7} to csv!
                                                        Current Time: {1}.
                                                        Start Time: {2}.
                                                        Current Page: {0}. 
                                                        Total Pages: {3}.
                                                        Average Time per {8} pages: {4:.2f}s.
                                                        Time spent on the last {8} pages: {6:.2f}s
                                                        Approximately another {5:.2f} days to finish rest of data.
                                                        Reported by Orbis Data Scraping System
                                            """.format(page_done, time.ctime(), start_datetime, total_companies, avg_time,
                                                       (total_companies - page_done) / per_round * avg_time / 86400,
                                                       round_time_spent, round, per_round))
            company_data = company_data.iloc[0:0]
            round += 1

            try:
                if (page_done % 50000 == 0) & (report != "NO"):
                    gmail_user = 'hlr.arndt@gmail.com'
                    gmail_password = open("C:/Users/ad_arndt/Documents/pw.txt", "r").read()

                    # Python code to illustrate Sending mail from
                    # your Gmail account
                    # creates SMTP session
                    s = smtplib.SMTP('smtp.gmail.com', 587)
                    # start TLS for security
                    s.starttls()
                    # Authentication
                    s.login(gmail_user, gmail_password)
                    # message to be sent
                    message =  """\
                    Subject: Orbis Scraping System is running.\n\n           
                                            
                    Current Time: {1}.
                    The system is running smoothly for {2:.2f} days now.
                    Current Page: {0}. 
                    Total Pages: {3}.
                    Average Time per 50 pages: {4:.2f}s.
                    Time spent on the last 50 pages: {6:.2f}s
                    Approximately another {5:.2f} days to finish rest of data.
                    
                    The Orbis Data Scraping System wishes you a great day!
                    """.format(page_done+1, time.ctime(), (time.ctime()-start_datetime)/ 86400, total_companies, avg_time,
                               (total_companies - page_done) / 1000 * avg_time / 86400,
                               round_time_spent)

                    # sending the mail
                    s.sendmail(gmail_user, "lu@mpifg.de", message)
                    # terminating the session
                    s.quit()
            except:
                print('Sending the email did not work..')
                pass
    except Exception as ex:
        print(ex)
        return page_done+1

    print("Finished sucessfully!")
    return -99

def ownership_starter(browser):

    # missing_ids = [5617, 9533, 14830]
    missing_ids = pd.read_csv("missing_ids.csv", sep=';')

    while True:
        same_val = 0
        success = 0
        while same_val < 3:
            try:
                # Open a new session
                success_t_1 = start_page
                success = scrape_ownership_report(browser, startpage=start_page)
                print("The Scrape System is restarted because of company {0} for the {1}. time. After the third time, the company is skipped.".format(success, same_val+1))
                if success == -99:
                    break
                else:
                    if success == success_t_1: same_val += 1
                    start_page = success
                    try:
                        browser.close()
                    except:
                        pass
                    continue
            except:
                print("There was some uncaught error..")
                try:
                    browser.close()
                except:
                    pass
                continue
        if success != -99:
            print("Company page " + str(success) + " could not be scraped because of an error. It is skipped but its ID is stored for later manual inspection.")
            missing_ids.append(start_page)
            pd.DataFrame(missing_ids).to_csv("missing_ids.csv", sep=';', index=False)
            print("So far skipped companies:")
            for x in missing_ids:
                print(x)
            start_page += 1
        elif success == -99:
            print("Ownership reports successfully scraped!"
                  ""
                  "However, the following companys could not be scraped and need to be checked manually: ")
            for x in missing_ids:
                print(x)
            print("This list was also printed to the 'missing_ids.csv' in the working directory.")
            break

######### Settings for scraping #####################
# year_to_get = list(range(2019,2009,-1))
start_page = 14800
#####################################################

start_time = time.time()
start_datetime = time.ctime()
big_round_time = start_time
hard_refresh_times = 0
# Select base variables
#try:
#    browser = sel_base_info_vars()
#except:
#    browser = sel_base_info_vars()

# Scrape the data
#scrape_table(browser)
#browser.close()

browser = login_orbis()
# export_directors(browser)

# if __name__ == '__main__':