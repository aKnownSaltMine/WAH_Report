# ## Import Dependencies and variables
# native imports
import os
import warnings
from datetime import date
from shutil import copy, move
from time import sleep
from Dependencies.setup import setup
from Dependencies import gvp_functions as gvp

try:
    import numpy as np
    import pandas as pd
    import pyodbc
    import win32com.client
    from dateutil.relativedelta import relativedelta, MO
except ImportError:
    setup()
    import numpy as np
    import pandas as pd
    import pyodbc
    import win32com.client
    from dateutil.relativedelta import relativedelta, MO

warnings.simplefilter(action='ignore')

def calc_weekend_work(shift):
    # function in order to calculate if the agent currently works weekends based on the days worked string
    if isinstance(shift, float):
        return np.nan
    elif ('Y' in shift) or ('S' in shift):
        return True
    else:
        return False
    


def main():
    # Declaring lookback dates
    today = date.today()
    # today = date(2023, 3, 20)
    last_month = today + relativedelta(months=-1)
    last_fiscal = gvp.decide_fm(last_month)
    fiscal_beginning = gvp.decide_fm_beginning(last_fiscal)
    fiscal_end = gvp.decide_fm_end(last_fiscal)
    lookback_beginning = gvp.decide_fm_beginning(
        last_fiscal + relativedelta(months=-2))
    lookback_end = gvp.decide_fm_end(last_fiscal)

    # hard code checks
    wah_date = date(2022, 8, 1)
    required_ot = 4
    final_shrink_date = lookback_end + relativedelta(days=14)
    ot_fiscal = date(2023, 4, 1)

    ot_applied = None
    while ot_applied == None:
        ot_check = input('Is Overtime being applied for this month? (y/n) ')
        ot_check = ot_check.lower()
        if ot_check == 'y':
            ot_applied = True
        elif ot_check == 'n':
            ot_applied = False
        else:
            print(f'{ot_check} is not an acceptable answer. Please only input y for Yes or n for No.')


    print("-"*25)
    print(
        f'Running the WAH Eligibility report for {last_fiscal.strftime("%B %Y")}')

    # checking to see if the date in which shink has been finalized has passed. If not display a message, then close
    if today < final_shrink_date:
        print(
            f'Shrink data has not finalized and will not until {final_shrink_date.strftime("%m/%d/%Y")}')
        print('Please try again running the report on or after that date.')
        sleep(60)
        exit()

    # getting the previous month of lookback period in order to calculate previous month to check for 2 month fails
    last_lookback_month = last_fiscal + relativedelta(months=-1)
    last_lookback_beginning = gvp.decide_fm_beginning(
        last_lookback_month + relativedelta(months=-2))
    last_lookback_end = gvp.decide_fm_end(last_lookback_month)


    # declares working paths
    cwd = os.path.dirname(__file__)

    # declaring queries paths
    query_folder = os.path.join(cwd, 'Queries')
    ot_query_name = 'Overtime_Query_Fiscal.sql'
    shrink_query_name = 'Shrink_Query.sql'
    roster_query_name = 'VR_Roster_Query.sql'
    ot_query_path = os.path.join(query_folder, ot_query_name)
    shrink_query_path = os.path.join(query_folder, shrink_query_name)
    roster_query_path = os.path.join(query_folder, roster_query_name)

    # getting data folders
    data_folder = os.path.join(cwd, 'Data')
    if os.path.exists(data_folder) == False:
        os.makedirs(data_folder)
    data_file = 'WAH Data.xlsx'
    data_path = os.path.join(data_folder, data_file)
    archive_folder = os.path.join(data_folder, 'Prior WAH Data')
    archive_file = f'{last_fiscal.strftime("%m%Y")}_{data_file}'
    archive_path = os.path.join(archive_folder, archive_file)
    performance_file = 'Percentile_ranks.xlsx'
    performance_path = os.path.join(data_folder, performance_file)


    """
    This section searches through the user's outlook email inbox in order to find the most
    recent disciplinary file that is emailed from HR every monday. 
    """
    # calculating the latest ca file inside of the data file
    # launching outlook in order to search through email items
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # default folder 6 is the inbox
    messages = inbox.Items
    today_string = today.strftime("%m/%d/%y")
    monday = today - relativedelta(weekday=MO(-1))
    print(monday)
    subject = "CA's"
    found = False  # flag for if it has been found in order to not search everything

    for message in messages:
        # loop through all of the messages and test the subject and sent date
        if subject in message.Subject and message.Senton.date()>=monday:
            # if the subject matches and was sent on day of running
            # then download all of the attachments ending in xlsb
            attachments = message.Attachments
            attachment = attachments.Item(1)
            print(f'Found message with subject: {message.Subject}')
            for attachment in message.Attachments:
                if str(attachment).endswith('.xlsx'):
                    download_path = os.path.join(data_folder, str(attachment))
                    ca_path = download_path
                    if os.path.exists(download_path):
                        os.remove(download_path)
                        print('Old Copy Removed')

                    # save directly into data folder
                    attachment.SaveAsFile(download_path)
                    print(f'Saved: {download_path}')
                    found = True  # set flag to true, then break out of loop to stop search
                    break
        if found:  # break out of loop to end search
            break
    else:
        print(
            f'No messages recieved {monday.strftime("%m/%d/%y")} found with the subject: {subject}')
        exit()

    ca_archive_folder = os.path.join(data_folder, 'Prior CA Files')
    ca_files = [value for value in os.listdir(data_folder) if value.startswith(
        'CA') and value.endswith('xlsx')]
    ca_paths = [os.path.join(data_folder, basename) for basename in ca_files]
    ca_path = max(ca_paths, key=os.path.getctime)

    for file in ca_files:
        file_path = os.path.join(data_folder, file)
        if file_path == ca_path:
            continue
        else:
            move(file_path, os.path.join(ca_archive_folder, file))

    print(f'Latest Path for CAs:{ca_path}')

    # declaring template file
    template_file = 'WAH_Eligibility_StatusCheck_Template.xlsx'
    template_folder = os.path.join(cwd, 'Templates')
    template_path = os.path.join(template_folder, template_file)

    # output paths
    reports_folder = os.path.join(cwd, 'Reports')
    server_folder = r'' # Network Share drive
    # server_folder = cwd
    save_name = f'WAH_Eligibility_StatusCheck_{fiscal_end.strftime("%m%d%y")}.xlsx'
    save_path = os.path.join(reports_folder, save_name)
    server_path = os.path.join(server_folder, save_name)
    src_folder = os.path.join(cwd, 'src')

    # ## Queries and Server connection

    # reading in each of the sql queries from the queries folder
    with open(ot_query_path, 'r') as query:
        ot_query = query.read()
    with open(shrink_query_path, 'r') as query:
        shrink_query = query.read()
    with open(roster_query_path, 'r') as query:
        roster_query = query.read()


    # connection string to access server and creating the server connection
    conn_str = ("Driver={SQL Server};"
                "Server=;" # Network Server Address
                "Database=Aspect;"
                "Trusted_Connection=yes;"
                "ApplicationIntent=ReadOnly")

    # creating connection to server
    conn = pyodbc.connect(conn_str)
    print('Connecting to Server')

    # ## Creating Source Dataframes

    # reading in the source dataframes
    # Reading the roster dataframe from the server
    print('Retrieving Roster')
    roster_df = pd.read_sql(roster_query, conn)
    print('Roster Dataframe Created')
    print('-'*25)
    # correcting roster dataframe
    roster_df = roster_df.loc[roster_df['TERMINATEDDATE'].isna()]
    roster_df['NETIQWORKERID'] = roster_df['NETIQWORKERID'].astype(int).astype(str)
    roster_df['HIREDATE'] = pd.to_datetime(roster_df['HIREDATE']).dt.date
    roster_df['WP Start Date'] = pd.to_datetime(roster_df['WP Start Date']).dt.date
    print('Roster corrected')

    # splitting the location into centers as well as correcting for Gran Vista
    for index, row in roster_df.iterrows():
        call_center = row['MGMTAREANAME']
        location = row['WorkLocation']
        city = ' '.join(location.split(' ')[1:])
        state = location.split(' ')[0]

        updated_location = f'{city} {state}'
        if 'Gran Vista' in call_center:
            updated_location = f'{updated_location} (Gran Vista)'
        roster_df.loc[index, 'MGMTAREANAME'] = updated_location  # type: ignore

    # dropping columns in this way if another is accidentially added, it wont break script
    columns_to_keep = ['BossName',
                    'BossBossName',
                    'EmpName',
                    'EmpTitle',
                    'MGMTAREANAME',
                    'NETIQWORKERID',
                    'HIREDATE',
                    'Days Worked',
                    'Start/Stop',
                    'WorkPlace',
                    'WP Start Date']
    drop_columns = [
        value for value in roster_df.columns if value not in columns_to_keep]
    roster_df = roster_df.drop(columns=drop_columns)

    # reading the percentile dataframe and correcting types
    print('Retrieving Percentile Data')
    percentile_df = pd.read_excel(performance_path, engine='openpyxl')
    print('Percentile Data Loaded')
    print('-'*25)
    percentile_df['Fiscal Month'] = pd.to_datetime(
        percentile_df['Fiscal Month']).dt.date
    percentile_df['PSID'] = percentile_df['PSID'].astype(str)
    percentile_df['Overall Rank'] = percentile_df['Overall Rank'].astype(float)

    # checking for the latest fiscal month and ensuring that it is in the dataset before moving forward
    max_date = percentile_df['Fiscal Month'].max()
    if max_date != last_fiscal:
        print("It looks like we do not have last month's data added in the Percentile Ranks. Please correct that by grabbing the lastest rankings from the ranking email and run again.")
        sleep(60)
        exit()

    # Reading in the CA dataframe and correcting datatypes
    print('Retrieving Corrective Action Data')
    ca_df = pd.read_excel(ca_path, dtype=object)
    print('Corrective Action Data Loaded')
    print('-'*25)
    date_list = ['Effective Date (Occurence Dt.)',
                'Purge Date',
                'Term Date']
    for column in date_list:
        ca_df[column] = pd.to_datetime(ca_df[column]).dt.date
    ca_df = ca_df.loc[ca_df['Purge Date'] >= last_lookback_beginning]
    ca_df['PSID'] = ca_df['PSID'].astype(int).astype(str)

    ca_df = ca_df.sort_values(
        'Effective Date (Occurence Dt.)').drop_duplicates('PSID', keep='last')

    # reading in overtime dataframe
    print('Retrieving Overtime Data')
    overtime_df = pd.read_sql(ot_query, conn)
    print('Ovetime Data loaded')
    print('-'*25)

    overtime_df['PSID'] = overtime_df['PSID'].astype(int).astype(str)
    overtime_df['FiscalMonth'] = pd.to_datetime(overtime_df['FiscalMonth']).dt.date

    # reading in the shrink dataframe and correcting datatypes
    print('Retrieving shrink data')
    shrink_df = pd.read_sql(shrink_query, conn)
    print('Shrink Data loaded')
    print('-'*25)
    shrink_df['FiscalMonth'] = pd.to_datetime(shrink_df['FiscalMonth']).dt.date
    shrink_df['EmpID'] = shrink_df['EmpID'].astype(str)

    # ## Creating Prior Month WAH eligibility
    # creating the prior percentile dataframe and calculating those who passed threshold
    print('-'*25)
    print('Calculating prior month')
    prior_percentile_df = percentile_df.loc[percentile_df['Fiscal Month'].between(last_lookback_beginning, last_lookback_end)].groupby('PSID').agg({
        'Overall Rank': 'mean'
    }).reset_index()
    prior_percentile_df['Over 50'] = prior_percentile_df['Overall Rank'].map(
        lambda x: True if x >= 50 else False)

    # creating a list of psid's that have a corrective during the prior lookback period
    prior_ca_list = ca_df.loc[ca_df['Effective Date (Occurence Dt.)'].between(
        last_lookback_beginning, last_lookback_end)]['PSID'].tolist()
    prior_ca_list = [str(value) for value in prior_ca_list]

    # creating the prior OT Dataframe and mapping those who are over the required OT
    prior_ot_df = overtime_df.loc[overtime_df['FiscalMonth']
                                == last_lookback_month]
    prior_ot_df['OT Met'] = prior_ot_df['OT Total'].map(
        lambda x: True if x > required_ot else False)

    # creating the prior shrink dataframe and creating columns to see who
    prior_shrink_df = shrink_df.loc[shrink_df['FiscalMonth'].between(last_lookback_beginning, last_lookback_end)].groupby('EmpID').agg({
        'Unplanned OOO': 'sum',
        'Scheduled': 'sum'
    }).reset_index()
    prior_shrink_df['Shrinkage'] = prior_shrink_df['Unplanned OOO'] / \
        prior_shrink_df['Scheduled']
    prior_shrink_df['Shrink Pass'] = prior_shrink_df['Shrinkage'].map(
        lambda x: True if x <= .07 else False)

    # ## Creating the WAH Dataframe

    # renaming the roster dataframe in order to create a more readable wah export
    roster_rename_dict = {'BossName': 'Supervisor',
                        'BossBossName': 'Manager',
                        'EmpName': 'Agent',
                        'EmpTitle': 'Title',
                        'NETIQWORKERID': 'PSID',
                        'MGMTAREANAME': 'Call Center'}
    wah_df = roster_df.rename(columns=roster_rename_dict)
    # filtering for just reps
    wah_df = wah_df.loc[wah_df['Title'].str.contains('Rep ')]

    # creating columns for wah and if they were prior to the cutoff date
    wah_df['Remote'] = wah_df['WorkPlace'].map(lambda x: x if (
        x == None) else True if (x.startswith('WAH')) else False)
    wah_df['WAH Prior'] = wah_df['WP Start Date'].map(
        lambda x: x if (x == None) else True if x <= wah_date else False)
    # correcting for the people who are not remote in order to not mark them eligible if they have been in center since before the cutoff date
    wah_df.loc[(wah_df['WAH Prior'] == True) & (
        wah_df['Remote']) == False, 'WAH Prior'] = False
    # parsing the Days Worked column to create a column of booleans if weekends are worked
    wah_df['Works_Weekend'] = wah_df['Days Worked'].map(
        lambda x: calc_weekend_work(x) if x != None else x)

    # ## Merging prior data with a cloned wah dataframe

    prior_wah_df = wah_df

    # joining the prior wah with the created prior dataframes
    prior_wah_df = prior_wah_df.merge(prior_percentile_df.loc[:, [
                                    'PSID', 'Over 50']], how='left', left_on='PSID', right_on='PSID')
    prior_wah_df = prior_wah_df.merge(prior_shrink_df.loc[:, [
                                    'EmpID', 'Shrink Pass']], how='left', left_on='PSID', right_on='EmpID').drop(columns='EmpID')
    prior_wah_df['No CA'] = prior_wah_df['PSID'].map(
        lambda x: True if x not in prior_ca_list else False)
    prior_wah_df = prior_wah_df.merge(
        prior_ot_df.loc[:, ['PSID', 'OT Met']], how='left', left_on='PSID', right_on='PSID')
    prior_wah_df.loc[prior_wah_df['Remote'] == False, 'OT Met'] = np.nan

    # looping through the prior wah df in order to create a column if they passed the prior month in order to see if the
    # agent should come back in office or not
    for index, row in prior_wah_df.iterrows():
        remote = row['Remote']
        weekends = row['Works_Weekend']
        performance = row['Over 50']
        shrink = row['Shrink Pass']
        ca = row['No CA']
        ot = row['OT Met']
        wah_prior = row['WAH Prior']

        results = []

        # rather than creating nested if statements, appending results to a list to check the length of the list to see if agent passed
        if (wah_prior == False) and (weekends == False):
            results.append('Weekends')
        if performance == False:
            results.append('Performance')
        if ca == False:
            results.append('CA')
        if shrink == False:
            results.append('Shrink')
        # if (remote == True) and (ot == False) and (ot_applied):
        #     results.append('OT')

        if len(results) == 0:
            prior_wah_df.loc[index, 'Pass Last FM'] = True  # type: ignore
        else:
            prior_wah_df.loc[index, 'Pass Last FM'] = False  # type: ignore
    else:
        print('Prior month has concluded calculating')
        print('-'*25)

    # ## Creating data used for this month
    # joining the results of the prior month onto the wah df
    print('-'*25)
    print('Calculating this month')
    wah_df = wah_df.merge(
        prior_wah_df.loc[:, ['PSID', 'Pass Last FM']], how='left', on='PSID')

    # creating a dataframe to calculate the number of scorecards inside of the percentile df exist for the agent
    data_df = percentile_df.loc[(percentile_df['Overall Rank'].notna()) & (
        percentile_df['Fiscal Month'] >= lookback_beginning)]
    data_df = data_df.groupby('PSID').count()['Fiscal Month'].reset_index().rename(
        columns={'Fiscal Month': 'Months_of_Data'})

    # dropping unnecessary columns
    columns_to_keep = ['PSID', 'Months_of_Data']
    drop_columns = [
        value for value in data_df.columns if value not in columns_to_keep]
    data_df = data_df.drop(columns=drop_columns)

    # averaging the last 3 months of percentile data, and adding in how many months of data are possesed by the agent
    percentile_df = percentile_df.loc[percentile_df['Fiscal Month'] >= (
        last_fiscal + relativedelta(months=-2))].groupby('PSID').mean()['Overall Rank'].reset_index()
    percentile_df = percentile_df.merge(data_df, how='left', on='PSID')
    # creating a boolean column of agents who passed the 50% mark
    percentile_df['Over 50'] = percentile_df['Overall Rank'].map(
        lambda x: True if x >= 50 else False)

    # joining the new percentile_df onto the wah_df
    wah_df = wah_df.merge(percentile_df.loc[:, [
                        'PSID', 'Overall Rank', 'Months_of_Data', 'Over 50']], how='inner', left_on='PSID', right_on='PSID')

    # filtering the ca dataframe for the last three fiscal months of ca's
    ca_df = ca_df.loc[ca_df['Effective Date (Occurence Dt.)']
                    >= lookback_beginning]

    # sending the unique ca values to a list and mapping that list to the wah_df to create a boolean column
    ca_list = ca_df['PSID'].unique().tolist()
    ca_list = [str(value) for value in ca_list]
    wah_df['No CA'] = wah_df['PSID'].map(
        lambda x: True if x not in ca_list else False)
    wah_df = wah_df.merge(ca_df, how='left', on='PSID')

    # creating a list of individuals above OT cutoff and then creating a column from the results
    overtime_list = overtime_df.loc[overtime_df['OT Total']
                                    >= required_ot, 'PSID'].unique().tolist()
    overtime_list = [str(value) for value in overtime_list]
    wah_df['Overtime Met'] = wah_df['PSID'].map(
        lambda x: True if x in overtime_list else False)
    wah_df.loc[wah_df['Remote'] == False, 'Overtime Met'] = np.nan
    wah_df = wah_df.merge(overtime_df.loc[overtime_df['FiscalMonth'] == last_fiscal, [
                        'PSID', 'OT Total']], how='left', on='PSID')

    # filtering the shrink dataframe and creating a column of people below the threshold.
    shrink_df = shrink_df.loc[shrink_df['FiscalMonth'] >= lookback_beginning].groupby('EmpID').agg({
        'Unplanned OOO': 'sum',
        'Scheduled': 'sum'
    }).reset_index()
    shrink_df['Shrinkage'] = shrink_df['Unplanned OOO'] / shrink_df['Scheduled']
    shrink_df['Shrink Pass'] = shrink_df['Shrinkage'].map(
        lambda x: True if x <= .07 else False)
    # joining the shrink dataframe with the main wah_df for the Shrink Pass Column
    wah_df = wah_df.merge(shrink_df.loc[:, ['EmpID', 'Shrinkage', 'Shrink Pass']],
                        how='left', left_on='PSID', right_on='EmpID').drop(columns='EmpID')

    # creating a column of people who have met the tenure requirement
    wah_df['Enough Data'] = wah_df['Months_of_Data'].map(
        lambda x: True if x >= 3 else False)
    wah_df.loc[(wah_df['Remote'] == True) & (
        wah_df['Months_of_Data'] >= 2), 'Enough Data'] = True


    # looping through the df to find the people who fit the criteria in order to mark them eligible
    # if they are not eligible, or did not pass this month, create a reason string in order to explain why
    # also account for the people who have failed two months in a row vs only one month
    for index, row in wah_df.iterrows():
        remote = row['Remote']
        data_check = row['Enough Data']
        weekends = row['Works_Weekend']
        performance = row['Over 50']
        ca = row['No CA']
        shrink = row['Shrink Pass']
        ot = row['Overtime Met']
        last_fm = row['Pass Last FM']
        wah_prior = row['WAH Prior']

        results = []
        if ca == False:
            results.append('CA')
        if data_check == False:
            results.append('Data')
        if (remote) and (wah_prior == False) and (weekends == False):
            results.append('Weekends')
        if performance == False:
            results.append('Performance')
        if shrink == False:
            results.append('Shrink')
        if (remote) and (ot == False) and (ot_applied):
            results.append('OT')

        if len(results) == 0:
            wah_df.loc[index, 'Pass This Month'] = True  # type: ignore
            wah_df.loc[index, 'WAH Eligible'] = 'Yes'  # type: ignore
        else:
            wah_df.loc[index, 'Pass This Month'] = False  # type: ignore

            data_str = 'not having enough data'
            weekend_str = 'not being scheduled for a weekend day'
            performance_str = 'not meeting average performance'
            ca_str = 'being on a Corrective Action'
            shrink_str = 'too many unplanned absences'
            ot_str = f'not working the required {required_ot} hours of overtime'

            dictionary = {'Weekends': weekend_str,
                        'Performance': performance_str,
                        'CA': ca_str,
                        'OT': ot_str,
                        'Data': data_str,
                        'Shrink': shrink_str}
            reason = ''

            # if the results are a certain length, the string will reference the dictionary to return the reason and build the sentence
            if len(results) == 1:
                reason = f'This agent did not pass this month due to {dictionary.get(results[0])}.'
            elif len(results) == 2:
                reason = f'This agent did not pass this month due to {dictionary.get(results[0])} and {dictionary.get(results[1])}.'
            elif len(results) == 3:
                reason = f'This agent did not pass this month due to {dictionary.get(results[0])}, {dictionary.get(results[1])}, and {dictionary.get(results[2])}.'
            elif len(results) == 4:
                reason = f'This agent did not pass this month due to {dictionary.get(results[0])}, {dictionary.get(results[1])}, {dictionary.get(results[2])}, and {dictionary.get(results[3])}.'
            elif len(results) == 5:
                reason = f'This agent did not pass this month due to {dictionary.get(results[0])}, {dictionary.get(results[1])}, {dictionary.get(results[2])}, {dictionary.get(results[3])}, and {dictionary.get(results[4])}.'
            elif len(results) == 6:
                reason = f'This agent did not pass this month due to {dictionary.get(results[0])}, {dictionary.get(results[1])}, {dictionary.get(results[2])}, {dictionary.get(results[3])}, {dictionary.get(results[4])}, and {dictionary.get(results[5])}.'

            # check to see if the person failed two months in a row to see if they are still eligble or must come in office.
            if remote and last_fm and ca:
                reason = f'{reason}  They have one month to correct before coming back to the office.'
                wah_df.loc[index, 'WAH Eligible'] = 'Yes'  # type: ignore
            elif remote and ca == False:
                reason = f'{reason} They must come into office due to their CA.'
                wah_df.loc[index, 'WAH Eligible'] = 'No'  # type: ignore
            elif remote and last_fm == False:
                reason = f'{reason} This is their second month in a row they have failed a WAH qualification. They must come into office.'
                wah_df.loc[index, 'WAH Eligible'] = 'No'  # type: ignore
            else:
                wah_df.loc[index, 'WAH Eligible'] = 'No'  # type: ignore
            wah_df.loc[index, 'Reason'] = reason  # type: ignore
    else:
        print('Current month completed in calculating')
        print('-'*25)

    wah_df.loc[(wah_df['WAH Eligible'] == 'Yes') & (wah_df['Remote'] == False) & (wah_df['Works_Weekend'] == False),
            'Reason'] = "This agent qualifies for WAH on performance, however they do not work weekends and must change their schedule before deployment."

    # adding a tilde to the beginning of schedules so that excel does not treat them as formulas
    wah_df['Days Worked'] = wah_df['Days Worked'].map(
        lambda x: "`" + x if x != None else x)
    print('Corrected schedule strings')

    # reordering columns in order to output and sorting values
    new_column_order = ['Call Center',
                        'Manager',
                        'Supervisor',
                        'Agent',
                        'Title',
                        'PSID',
                        'HIREDATE',
                        'Days Worked',
                        'Start/Stop',
                        'WP Start Date',
                        'WorkPlace',
                        'Overall Rank',
                        'Months_of_Data',
                        'Displinary Action/Reason',
                        'Effective Date (Occurence Dt.)',
                        'Purge Date',
                        'Discp Step Descr',
                        'Term Date',
                        'OT Total',
                        'Shrinkage',
                        'Remote',
                        'WAH Prior',
                        'Works_Weekend',
                        'Pass Last FM',
                        'Enough Data',
                        'Over 50',
                        'No CA',
                        'Overtime Met',
                        'Shrink Pass',
                        'Pass This Month',
                        'WAH Eligible',
                        'Reason']

    wah_df = wah_df.reindex(columns=new_column_order)
    wah_df = wah_df.sort_values(
        by=['Call Center', 'Manager', 'Supervisor', 'Agent'], ignore_index=True)
    wah_df.loc[wah_df['WorkPlace'] == 'WAH-PF', 'WorkPlace'] = 'Remote'
    wah_df.loc[wah_df['WorkPlace'] ==
            'WIC-WORK_IN_CENTER', 'WorkPlace'] = 'In-Center'
    print('Output columns corrected')

    print('Writing to excel data sheet')
    wah_df.to_excel(data_path, index=False, sheet_name='WAH_Stutus')
    wah_df.to_excel(archive_path, index=False, sheet_name='WAH_Stutus')
    print('Excel sheet written')
    print('-'*25)

    ca_max_date = ca_df['Effective Date (Occurence Dt.)'].max()

    disclaimer_string = f'*CA data through {ca_max_date.strftime("%m/%d")}, all other data through {lookback_end.strftime("%m/%d")}. Evaluation period: {lookback_beginning.strftime("%m/%d")} - {lookback_end.strftime("%m/%d")}'

    if ot_applied == True:
        disclaimer_string = f'{disclaimer_string} Overtime below goal, OT requirement applied to remote agents.'
    elif ot_applied == False: 
        disclaimer_string = f'{disclaimer_string} Over 100% of Overtime Goal, OT requirement waived for this month.'

    print("Opening Excel")
    xlapp = win32com.client.Dispatch('Excel.Application')
    xlapp.Visible = True
    xlapp.DisplayAlerts = False
    wb = xlapp.Workbooks.Open(template_path)
    print('Excel has been opened')

    # refreshing all queries
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    print('Excel Data has been refreshed.')

    # deleting connections for output file
    for conn in wb.Queries:
        conn.Delete()
    print('Connections have been removed.')

    ws = wb.Worksheets('WAH_Status')
    ws.Range('A10').Value = disclaimer_string
    print('Updated the Update Date')

    # saving file in the determined folder and quitting excel
    wb.SaveAs(save_path)
    print(f'Workbook has been saved here: {save_path}')
    xlapp.DisplayAlerts = True
    wb.Close()
    xlapp.Quit()
    print('Excel has been closed.')

    # outputting to the website
    copy(save_path, server_path)
    print(f'Report has been saved to {server_path}')

    # creating a display dataframe in order to embed into html for email
    display_df = wah_df.groupby(['Call Center', 'WorkPlace', 'WAH Eligible']).agg({
        'WAH Eligible': 'count'
    }).rename(columns={'WAH Eligible': 'Count'}).reset_index().pivot(index=['WorkPlace', 'WAH Eligible'], columns='Call Center', values='Count')
    display_df['Total'] = display_df.sum(axis=1)

    # creating a list of columns to loop through in order to create a total row
    not_cc = ['WorkPlace', 'WAH Eligible']
    cc_columns = [value for value in display_df.columns if value not in not_cc]

    for column in cc_columns:
        hc_total = display_df[column].sum()
        wah_total = display_df.loc[('Remote', 'Yes'), column].sum() + display_df.loc[[  # type: ignore
            ('In-Center', 'Yes')], column].sum()  # type: ignore
        wah_pct = wah_total / hc_total

        display_df.loc[('Grand Total', 'Eligible Total'), column] = wah_total
        display_df.loc[('Grand Total', 'WAH Eligible %'),
                    column] = '{:.2%}'.format(wah_pct)

    index_list = [('Remote', 'No'),
                ('Remote', 'Yes'),
                ('In-Center', 'No'),
                ('In-Center', 'Yes'),
                ('Grand Total', 'Eligible Total')]
    # looping through the indexes in order to correct from floats to integers
    for index in index_list:
        display_df.loc[index] = display_df.loc[index].astype(int)


    # creating email table
    email_display_df = display_df.style.set_table_styles(
        [{'selector': '',
        'props': 'border: 1px solid black; border-collapse: collapse; padding: 5px;'},
        {'selector': 'td',
        'props': 'border: 1px solid black; tborder-collapse: collapse; text-align: center;'},
        {'selector': '.row_heading',
        'props': 'text-align: left; font-weight: bold; border: 1px solid black; tborder-collapse: collapse;'},
        {'selector': 'thead',
        'props': 'background-color:#787878; color:white; border: 1px white; border-collapse: collapse;'},
        {'selector': '.index_name',
        'props': 'background-color:#787878; color:white; border: 1px white; border-collapse: collapse;'},
        {'selector': '.blank',
        'props': 'background-color:#787878; color:white; border: 1px white; border-collapse: collapse;'},
        {'selector': 'table',
        'props': 'border-collapse: collapse;'}])

    table = email_display_df.to_html()

    # declaring html to build email
    explainer = f"""
    <p>&nbsp;</p>
    <p class=MsoNormal><span style='color:black'>You can find the most recent snapshot for WAH Eligibility posted <a
                href=''><span style='font-size: 12.0pt'>here</span></a></span><span
            style='font-size:12.0pt;color:black'>.</span></p>
    <p class=MsoNormal><span style='color:black'>Below is a summary after the most recent data refresh.</span></p>
    <p><br>{table}</br></p>
    """
    subject = f"LEADER: WAH Eligibility - {last_fiscal.strftime('%B %Y')}"

    recipients_list = ['Network_Email_Address']

    gvp.generate_email(explainer, subject, 'leader', recipients_list)

if __name__ == '__main__':
    main()