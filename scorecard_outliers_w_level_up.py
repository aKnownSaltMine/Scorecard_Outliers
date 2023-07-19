# import native dependencies
import os
import warnings
from datetime import date
from pathlib import Path
from shutil import copy, move
from Dependencies.setup import setup
from Dependencies import gvp_functions as gvp

# import dependencies
try:
    import numpy as np
    import pandas as pd
    import pyodbc
    import win32com.client
    from dateutil.relativedelta import relativedelta
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    setup()
    import numpy as np
    import pandas as pd
    import pyodbc
    import win32com.client
    from dateutil.relativedelta import relativedelta
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager


warnings.simplefilter(action='ignore')
pd.set_option('display.max_columns', None)

# declares helper functions
def round_half_up(n, decimals=0):
    # function rounds up on a .5 split
    import math
    multiplier = 10 ** decimals
    return math.floor(n*multiplier + 0.5) / multiplier

def main():
    # calculating fiscal months and lookback periods
    today = date.today()
    yesterday = today + relativedelta(days=-1)
    current_fm = gvp.decide_fm(yesterday)

    shrink_date = yesterday + relativedelta(days=-30)
    shrink_fm = gvp.decide_fm(shrink_date)
    # shrink_fm = gvp.decide_fm(date(2023,1,1))


    months_to_pull = relativedelta(current_fm, shrink_fm).months + 1

    months_to_display = 9
    lookback_month = current_fm - relativedelta(months=months_to_display)
    print(f'Running for {current_fm.strftime("%B %Y")}')
    print(f'Lookback period through {lookback_month.strftime("%B %Y")}')

    # declare path and file names
    mstr_url = "" # Microstrategy URL
    saves_as = "Scorecard_Metrics.xlsx"

    cwd = os.path.dirname(__file__)
    # cwd = os.getcwd()
    data_folder = os.path.join(cwd, 'Data')
    queries_folder = os.path.join(cwd, 'Queries')

    threshold_file = 'Thresholds.xlsx'
    threshold_path = os.path.join(data_folder, threshold_file)
    old_scorecard_file = 'Old_Scorecard_Numbers.xlsx'
    old_scorecard_path = os.path.join(data_folder, old_scorecard_file)
    new_scorecard_file = 'New_Scorecard_Numbers.xlsx'
    new_scorecard_path = os.path.join(data_folder, new_scorecard_file)
    scorecard_file = 'Scorecard_Metrics.xlsx'
    scorecard_data = os.path.join(data_folder, scorecard_file)

    template_folder = os.path.join(cwd, 'Templates')
    template_file = 'Scorecard_Outlier_Template.xlsx'
    template_path = os.path.join(template_folder, template_file)

    save_folder = os.path.join(cwd, 'Reports')
    save_file = f'Scorecard Outliers - {yesterday.strftime("%m%d%y")}.xlsx'
    save_path = os.path.join(save_folder, save_file)
    server_folder = r'' # Network Share drive
    server_path = os.path.join(server_folder, save_file)
    roster_path = r"" # Network Share drive

    roster_query_file = 'VR_Roster_Query.sql'
    shrink_query_file = 'Shrink_Query.sql'
    hours_query_file = 'hours_query.sql'

    roster_query_path = os.path.join(queries_folder, roster_query_file)
    shrink_query_path = os.path.join(queries_folder, shrink_query_file)
    hours_query_path = os.path.join(queries_folder, hours_query_file)

    downloads_path = os.path.join(Path.home(), 'Downloads')


    # Pulling the Data
    # declare driver
    print('Downloading Drivers')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    print('Drivers downloaded')
    print('--------------------')

    gvp.download_reports(driver, mstr_url, saves_as, prompt='fm',
                        fiscal_month=current_fm, months=months_to_pull)

    download_file = os.path.join(downloads_path, saves_as)
    driver.quit()
    print(f'{saves_as} Report Downloaded. Moving...')

    # declaring save path from dataframe and renaming downloaded file to that path
    file_save_location = os.path.join(data_folder, saves_as)
    move(download_file, file_save_location)
    print(f'Saved: {file_save_location}')
    print('--------------------')

    # reading in queries to run against network server
    with open(shrink_query_path, 'r') as query:
        shrink_query = query.read()
    with open(roster_query_path, 'r') as query:
        roster_query = query.read()
    with open(hours_query_path, 'r') as query:
        hours_query = query.read()

    """
    Change the date boundaries of sql query using python rather than using variables in sql
    to allow for edits/adhoc runs of the report without having to edit too many fields.
    """
    shrink_query = shrink_query.replace('<<start>>', gvp.decide_fm_beginning(shrink_fm - relativedelta(months=2)).strftime("%m/%d/%Y"))
    shrink_query = shrink_query.replace('<<end>>', gvp.decide_fm_end(current_fm).strftime("%m/%d/%Y"))

    hours_query = hours_query.replace('<<start>>', gvp.decide_fm_beginning(shrink_fm).strftime("%m/%d/%Y"))
    hours_query = hours_query.replace('<<end>>', gvp.decide_fm_end(current_fm).strftime("%m/%d/%Y"))


    conn_str = ("Driver={SQL Server};"
                "Server=;" # Network Server Address
                "Database=Aspect;"
                "Trusted_Connection=yes;")
    conn = pyodbc.connect(conn_str)
    print('Connected to Server')


    # retreiving roster from ehh roster
    print('Retrieving Roster')
    roster_df = pd.read_sql(roster_query, conn)
    print('Roster Dataframe Created')
    print('-'*25)

    # correcting roster dataframe
    roster_df = roster_df.loc[roster_df['TERMINATEDDATE'].isna()]
    roster_df['NETIQWORKERID'] = roster_df['NETIQWORKERID'].astype(int).astype(str)
    roster_df['HIREDATE'] = pd.to_datetime(roster_df['HIREDATE']).dt.date
    # roster_df['WP Start Date'] = pd.to_datetime(roster_df['WP Start Date']).dt.date
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

    rename_dict = {'BossName': 'Supervisor',
                'BossBossName': 'Manager',
                'EmpName': 'Agent',
                'EmpTitle': 'Title',
                'NETIQWORKERID': 'PSID',
                'STATUSID': 'Status',
                'MGMTAREANAME': 'Call Center',
                'HIREDATE': 'Hire Date'}
    roster_df = roster_df.rename(columns=rename_dict)
    print('Corrected roster column names.')


    # reading in the shrink dataframe from traffic server and correcting datatypes
    print('Retrieving shrink data')
    shrink_df = pd.read_sql(shrink_query, conn)
    print('Shrink Data loaded')
    print('-'*25)
    shrink_df['Date'] = pd.to_datetime(shrink_df['Date']).dt.date
    shrink_df['EmpID'] = shrink_df['EmpID'].astype(str)

    # calculating the final shrink for fiscal months inside of the lookback period
    lookback_dates = [gvp.decide_fm_end(
        current_fm - relativedelta(months=value)) for value in range(months_to_pull)]

    final_shrink_df = pd.DataFrame()
    # calculating shrink for each fiscal month based on 3 month average
    for end_date in lookback_dates:
        if end_date.month == current_fm.month:
            end_date = yesterday

        start_date = gvp.decide_fm_beginning(end_date + relativedelta(months=-2))
        fiscal_month = gvp.decide_fm(end_date)

        print(
            f'Shrink Fiscal Month: {fiscal_month} \nStart Date: {start_date} \nEnd Date: {end_date}')
        print('-'*25)

        df = shrink_df.loc[shrink_df['Date'].between(start_date, end_date)]

        df['FiscalMonth'] = fiscal_month

        df = df.groupby(['FiscalMonth', 'EmpID']).agg({
            'Unplanned OOO': 'sum',
            'Scheduled': 'sum'
        }).reset_index()

        df['Attendance'] = 1 - (df['Unplanned OOO'] / df['Scheduled'])
        df = df.drop(columns=['Unplanned OOO', 'Scheduled'])

        final_shrink_df = pd.concat(
            [final_shrink_df, df], axis=0, ignore_index=True)
    print('Shrink has been calculated')


    # reading in the shrink dataframe from traffic server and correcting datatypes
    print('Retrieving hours data')
    hours_df = pd.read_sql(hours_query, conn)
    print('hours Data loaded')
    print('-'*25)
    hours_df['Date'] = pd.to_datetime(hours_df['Date']).dt.date
    hours_df['EmpID'] = hours_df['EmpID'].astype(str)
    hours_df = hours_df.fillna(0)
    hours_df['Hours Worked'] = (hours_df['Scheduled Hours'] - (
        hours_df['Out of Center - Planned'] + hours_df['Out of Center - Unplanned'])) / 3600
    hours_df['Hours Worked'] = hours_df['Hours Worked'].map(
        lambda x: 0 if x < 0 else x)
    hours_df['Fiscal Month'] = hours_df['Date'].map(gvp.decide_fm)

    hours_fm_df = hours_df.groupby(['Fiscal Month', 'EmpID']).agg({
        'Hours Worked': 'sum'
    }).reset_index()


    # reading in the mstr data excel sheet and merging with attendance data
    scorecard_df = pd.read_excel(scorecard_data, engine='openpyxl')
    scorecard_df['Agent - HR Number'] = scorecard_df['Agent - HR Number'].astype(
        str)
    scorecard_df['Fiscal Mth'] = pd.to_datetime(
        scorecard_df['Fiscal Mth'], format='%B %Y').dt.date
    scorecard_df = scorecard_df.merge(final_shrink_df, how='left', left_on=[
                                    'Fiscal Mth', 'Agent - HR Number'], right_on=['FiscalMonth', 'EmpID']).drop(columns=['FiscalMonth', 'EmpID'])
    print('Scorecard data merged with shrink data')

    # renaming columns
    rename_dict = {'Agent - HR Number': 'PSID',
                'Calls Handled': 'Calls',
                'Transfer Rate': 'Transfer Prevention',
                'FCR': 'FCR %',
                'Truck Roll Prevention': 'TRP %',
                'Attendance %': 'Attendance'}
    scorecard_df = scorecard_df.rename(columns=rename_dict)

    # calculating the transfer prefention, then adding in the roster data to the scorecard data
    scorecard_df['Transfer Prevention'] = 1 - scorecard_df['Transfer Prevention']
    scorecard_df = roster_df.loc[(roster_df['Title'].str.startswith('Rep ')) & ((roster_df['Title'].str.contains('Video')) | (roster_df['Title'].str.contains('Disability'))), [
        'Call Center', 'Manager', 'Supervisor', 'Agent', 'PID', 'PSID', 'Title', 'Hire Date']].merge(scorecard_df, how='inner', on='PSID')
    print('Scorecard merged with roster')

    scorecard_df = scorecard_df.merge(hours_fm_df, how='left', left_on=['PSID', 'Fiscal Mth'], right_on=[
                                    'EmpID', 'Fiscal Month']).drop(columns=['EmpID', 'Fiscal Month'])
    print('Scorecard merged with hours worked')


    # reading in the thresholds used to calculate agent's overall performance and correcting it
    threshold_df = pd.read_excel(threshold_path, engine='openpyxl')
    threshold_df.loc[threshold_df['Metric'] != 'AHT', 'Red'] = threshold_df.loc[threshold_df['Metric']
                                                                                != 'AHT', 'Red'].map(lambda x: ''.join(x[2:-1]))
    threshold_df.loc[threshold_df['Metric'] == 'AHT',
                    'Red'] = threshold_df.loc[threshold_df['Metric'] == 'AHT', 'Red'].map(lambda x: ''.join(x[2:]))
    threshold_df['Red'] = threshold_df['Red'].astype(float)
    threshold_df.loc[threshold_df['Metric'] != 'AHT',
                    'Red'] = threshold_df.loc[threshold_df['Metric'] != 'AHT', 'Red'] / 100
    date_list = ['StartDate', 'StopDate']
    for date_column in date_list:
        threshold_df[date_column] = pd.to_datetime(
            threshold_df[date_column]).dt.date
    print('Thresholds loaded')


    # looping through each unique fiscal month in the data, as well as each title, and each metric in order to calculate scorecard data
    title_list = scorecard_df['Title'].unique().tolist()
    metric_list = threshold_df['Metric'].unique().tolist()
    fiscal_mths = scorecard_df['Fiscal Mth'].unique().tolist()

    color_dict = {'Blue': 3,
                'Green': 3,
                'Yellow': 2,
                'Red': 1}
    """
    the below nested for loop iterates through each possible instance based on fiscal month,
    title, and metric in order to find those who match the qualifications and assign the color
    earned by their performance
    """

    for month in fiscal_mths:
        for title in title_list:
            for metric in metric_list:
                # grab the thresholds for each possibility of loop from the threshold_df 
                green = threshold_df.loc[(threshold_df['JobCodeDesc'] == title) & (threshold_df['Metric'] == metric) & (
                    threshold_df['StartDate'] <= month) & (threshold_df['StopDate'] >= month), 'Green'].values[0]  # type: ignore
                yellow = threshold_df.loc[(threshold_df['JobCodeDesc'] == title) & (threshold_df['Metric'] == metric) & (
                    threshold_df['StartDate'] <= month) & (threshold_df['StopDate'] >= month), 'Yellow'].values[0]  # type: ignore
                blue = threshold_df.loc[(threshold_df['JobCodeDesc'] == title) & (threshold_df['Metric'] == metric) & (
                    threshold_df['StartDate'] <= month) & (threshold_df['StopDate'] >= month), 'Level Up!'].values[0]  # type: ignore
                weight = threshold_df.loc[(threshold_df['JobCodeDesc'] == title) & (threshold_df['Metric'] == metric) & (
                    threshold_df['StartDate'] <= month) & (threshold_df['StopDate'] >= month), 'Weighting'].values[0]  # type: ignore

                if metric == 'AHT':
                    if blue != blue:  # checks to see if null
                        # if the blue is null, then find those that achieved green and mark the column as such
                        scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] <= green) & (
                            scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Green'
                    else:
                        # if blue isn not null, then find those who match blue achievement and mark the column
                        scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] <= blue) & (
                            scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Blue'
                        # find those that match for green and mark green
                        scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] <= green) & (
                            scorecard_df[metric] > blue) & (scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Green'
                    # find those who hit yellow and mark them for yellow
                    scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] <= yellow) & (
                        scorecard_df[metric] > green) & (scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Yellow'
                    # find those who hit below yellow, and mark them red
                    scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] > yellow) & (
                        scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Red'

                else:
                    if blue != blue:  # checks to see if null
                        # if the blue is null, then find those that achieved green and mark the column as such
                        scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] >= green) & (
                            scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Green'
                    else:
                        # find those that match for green and mark green
                        scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] >= blue) & (
                            scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Blue'
                        # find those that match for green and mark green
                        scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] >= green) & (
                            scorecard_df[metric] < blue) & (scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Green'
                    # find those who hit yellow and mark them for yellow
                    scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] >= yellow) & (
                        scorecard_df[metric] < green) & (scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Yellow'
                    # find those who hit below yellow, and mark them red
                    scorecard_df.loc[(scorecard_df['Title'] == title) & (scorecard_df[metric] < yellow) & (
                        scorecard_df['Fiscal Mth'] == month), f'{metric} color'] = 'Red'
                # create score column for the metric
                for key, value in color_dict.items():
                    scorecard_df.loc[scorecard_df[f'{metric} color']
                                    == key, f'{metric} score'] = value * weight

    scorecard_df = scorecard_df.replace([np.inf, -np.inf], 0)
    # create weighted score column by summing up score columns
    scores_columns = [
        value for value in scorecard_df.columns if value.endswith('score')]  # type: ignore
    scorecard_df['Weighted score'] = scorecard_df.loc[:,
                                                    scores_columns].sum(axis=1)

    # filling score columns with null if data is not complete
    for column in scores_columns:
        scorecard_df.loc[scorecard_df[column].isna(), 'Weighted score'] = np.nan

    # calculating overall color based on rounded numbers with .5 rounding up
    scorecard_df['Color score'] = scorecard_df['Weighted score'].loc[scorecard_df['Weighted score'].notna()].map(
        lambda x: round_half_up(x, decimals=0))
    scorecard_df.loc[scorecard_df['Color score'] == 3, 'Overall Color'] = 'Green'
    scorecard_df.loc[scorecard_df['Color score'] == 2, 'Overall Color'] = 'Yellow'
    scorecard_df.loc[scorecard_df['Color score'] == 1, 'Overall Color'] = 'Red'


    # creating a list of metrics that are used for level up
    level_up_list = threshold_df.loc[threshold_df['Level Up Metric']
                                    == 'Yes', 'Metric'].unique().tolist()
    level_up_list = [f'{value} color' for value in level_up_list]

    # mirroring the scorecard_df dataframe, then looping through all the metrics, and only keeping the ones that are blue in all metrics
    level_up_df = scorecard_df
    for metric in level_up_list:
        level_up_df = level_up_df.loc[level_up_df[metric] == 'Blue']
    # filtering out the people that have under 200 calls and under 80 hours worked
    level_up_df = level_up_df.loc[(level_up_df['Calls'] >= 200) & (
        level_up_df['Hours Worked'] >= 80)]

    level_up_df['Level Up'] = True  # creating a level up column
    scorecard_df = scorecard_df.merge(level_up_df.loc[:, ['PSID', 'Fiscal Mth', 'Level Up']], how='left', on=[
                                    'PSID', 'Fiscal Mth'])  # merging the level up dataframe onto the scorecard
    # those who are level up true, are marked blue for their overall color
    scorecard_df.loc[scorecard_df['Level Up'] == True, 'Overall Color'] = 'Blue'
    scorecard_df = scorecard_df.drop(columns='Level Up')  # drops level up column

    # pulling in the prior scorecard data, first the new data, and if there is still data missing, then the old scorecard data
    prior_scorecard_df = pd.read_excel(new_scorecard_path)

    prior_scorecard_df['Fiscal Mth'] = pd.to_datetime(
        prior_scorecard_df['Fiscal Mth']).dt.date
    prior_scorecard_df['PSID'] = prior_scorecard_df['PSID'].astype(str)

    max_month = scorecard_df['Fiscal Mth'].min() - relativedelta(months=1)

    cooking_df = scorecard_df.loc[:, ['PSID', 'Fiscal Mth', 'Overall Color', 'Weighted score']].rename(
        columns={'Weighted score': 'Overall Score'})
    cooked_df = prior_scorecard_df.loc[prior_scorecard_df['Fiscal Mth'] <= max_month]
    update_prior = [cooking_df, cooked_df]

    # update the new scorecard data file with the newly calculated values
    pd.concat(update_prior, axis=0, ignore_index=True).sort_values(
        by=['Fiscal Mth', 'PSID']).to_excel(new_scorecard_path, index=False)

    # dropping the score columns since they are unnecceary after having color
    scores_columns = [
        value for value in scorecard_df.columns if value.endswith('score')]  # type: ignore
    scorecard_df = scorecard_df.drop(columns=scores_columns)

    prior_scorecard_df = prior_scorecard_df.drop(
        columns='Overall Score').rename(columns={'Overall Color': 'Overall'})

    prior_scorecard_df = prior_scorecard_df.loc[prior_scorecard_df['Fiscal Mth'].between(
        lookback_month, max_month)]

    if prior_scorecard_df['Fiscal Mth'].min() > lookback_month:
        old_scorecard_df = pd.read_excel(old_scorecard_path)
        old_scorecard_df['Fiscal Mth'] = pd.to_datetime(
            old_scorecard_df['Fiscal Mth']).dt.date
        old_scorecard_df['PSID'] = old_scorecard_df['PSID'].astype(str)
        prior_scorecard_df = pd.concat(
            [prior_scorecard_df, old_scorecard_df], axis=0, ignore_index=True)
        prior_scorecard_df = prior_scorecard_df.loc[prior_scorecard_df['Fiscal Mth'].between(
            lookback_month, max_month)]

    flux_df = scorecard_df.rename(columns={'Overall Color': 'Overall'}).loc[scorecard_df['Fiscal Mth'] != current_fm, [
        'PSID', 'Fiscal Mth', 'Overall']]
    prior_scorecard_df = pd.concat(
        [prior_scorecard_df, flux_df], axis=0, ignore_index=True)

    month_list = prior_scorecard_df['Fiscal Mth'].unique().tolist()
    month_list.sort(reverse=True)
    month_list = [value.strftime('%b %y') for value in month_list]

    prior_scorecard_df['Fiscal Mth'] = pd.to_datetime(
        prior_scorecard_df['Fiscal Mth']).dt.strftime('%b %y')
    prior_scorecard_df = prior_scorecard_df.pivot(
        index='PSID', columns='Fiscal Mth', values='Overall').reindex(columns=month_list).reset_index()

    current_scorecard_df = scorecard_df.loc[scorecard_df['Fiscal Mth'] == current_fm]

    new_column_order = ['Call Center',  # changes the column order, and columns included
                        'Manager',
                        'Supervisor',
                        'Agent',
                        'PID',
                        'PSID',
                        'Title',
                        'Hire Date',
                        'Calls',
                        'Hours Worked',
                        'Overall Color',
                        'Attendance',
                        'Attendance color',
                        'FCR %',
                        'FCR % color',
                        'SAM %',
                        'SAM % color',
                        'Transfer Prevention',
                        'Transfer Prevention color',
                        'AHT',
                        'AHT color',
                        'TRP %',
                        'TRP % color']

    current_scorecard_df = current_scorecard_df.reindex(columns=new_column_order).rename(
        columns={'Overall Color': 'Overall', 'Transfer Prevention': 'Xfer Prev'}).sort_values(by=['Call Center', 'Manager', 'Supervisor', 'Agent'])

    current_scorecard_df['Agent'] = current_scorecard_df['Agent'].map(
        lambda x: x.title())
    current_scorecard_df['Supervisor'] = current_scorecard_df['Supervisor'].map(
        lambda x: x.title())
    current_scorecard_df['Manager'] = current_scorecard_df['Manager'].map(
        lambda x: x.title())


    final_scorecard_df = current_scorecard_df.merge(
        prior_scorecard_df, how='left', on='PSID')

    output_file = 'final_scorecard_data.xlsx'
    output_path = os.path.join(data_folder, output_file)
    final_scorecard_df.to_excel(
        output_path, index=False, sheet_name='Scorecard Data')


    print("Opening Excel")
    xlapp = win32com.client.Dispatch('Excel.Application')
    xlapp.Visible = True
    xlapp.DisplayAlerts = False
    wb = xlapp.Workbooks.Open(template_path)
    print('Excel has been opened')

    # refreshing all queries
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Worksheets('Summary Pivot').PivotTables(
        "PivotTable1").PivotCache().Refresh()
    print('Excel Data has been refreshed.')

    # deleting connections for output file
    for conn in wb.Queries:
        conn.Delete()
    print('Connections have been removed.')

    ws = wb.Worksheets('Scorecard Outliers')
    ws.Range('B9').Value = yesterday.strftime('%m/%d/%Y')
    print('Updated the Update Date')

    # hiding non used columns for when a new fiscal turns over and it slides formatting over with the inserted month
    ws.Range('AG:XFD').EntireColumn.Hidden = True
    # removing headers at top of column
    ws.ListObjects("Scorecard_Data").ShowAutoFilterDropDown = False

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

    # declaring html to build email
    explainer = """
        <p>&nbsp;</p>
        <p class=MsoNormal><span style='color:black'>You can find the most recent Scorecard Outlier report <a href=''><span
                style='font-size: 12.0pt'>here</span></a></span><span style='font-size:12.0pt;color:black'>.</span></p>
    """
    recipient_list = ['Network_Email_Address']
    subject = f"LEADER: Scorecard Outlier - {yesterday.strftime('%m/%d/%y')}"


    # generating email
    gvp.generate_email(explainer, subject, 'leader', recipient_list)



if __name__ == '__main__':
    main()