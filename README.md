# WAH_Report
## Overview
The business unit was moving away from distributing a relative stack rank to anyone below executive level. However, our system for Work at Home Agents still utilized this relative ranking system. So, I was asked to create a report that would cleanly state an agent's Work at Home eligibility as well as if they did not, explain why they did not using the Excel platform. With these contraints I created a script using Python and the Pandas library mostly in order to have a report that was easily replicatable, and able to be ran by any of our team members.

## Methodology
### The Rules
The rules for the agents to work from home were a bit complicated, but the were as follows:

#### For new deployment for work from home: 
* The agent must have been eligible for ranking for 3 months
    * This means that the agent took at least 200 calls and was staffed for 80 hours for each month being looked at
* In this relative stack rank, the agent must be at or above the 50th percentile for three month average
* The agent cannot have any displinary action taken against them in the three months being examained
* Schedule must include one weekend day worked
* Must be willing to work 4 hours minimum of Overtime per month if overtime is available (was not enforced if departmental overtime needs were met for the month)

#### Eligibility to Remain work from home:
* Agent must have been eligible for ranking for 2 months
    * This means that the agent took at least 200 calls and was staffed for 80 hours for 2 of the 3 months being looked at
* Must be at or above 50th percentile in 3 month average (if 2 months ranked, average of 2)
* No Disciplinary Action taken against the agent in any way
* Worked 4 or more OT hours in the prior month (can be waived if departmental overtime needs were met for the month)
* Must work at least one weekend day

### How it was accomplished
The report first calculates all of the dates needed based on how long it takes for data to finalize from the prior fiscal month and prevents the report from fulling running if the data has not finalized. 

Then it creates paths to all needed excel sheets and queries based on the path of the python file to allow the report folder to be moved around without breaking the script. 

The script then makes a connection to the Department SQL Server using pyodbc due to company's use of mssql. Pandas Dataframes are then created from the results of these queries as well as the data that is stored in Excel sheets. Then main dataframe for both the month being looked at and the preceding month are created and all of the rules are checked against in a series of True/False flags.

Once these flags are created for all rules, in order to create a reason string, I did not want to create a series of nested if statements in order to handle the construction of the string in an effort to keep the string at least semi-grammatically correct. 
This is why I decided to loop through each row, grab the True/False flags for each of the rules, then if any of these flags are False, then append them to a list. If the list has 0 items in it, then the agent passed and was eligible to work from home that month. Else, the script will start building the reason string by referencing the length of the list, and grabbing using the list items as keys in a string dictionary to append these using f strings. 

```
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

```

Then at the end of the loop, the script appends at the end of any reason string, whether they must come in to the office from home or if they have to fail one more month before returning. 

After these reason strings were created, then the dataframe was corrected to export to excel in two locations, one to archive the month for reference by the team if needed, and one as a data file for the template to use Power Query to refresh the data into the report template after which it was saved in our report repository, and our website file location for the end user to download, as well as generate an email with embeded stats that was sent to alert the end user of the refresh.

## Summary
Overall, the report was well recieved by the end users for its clean output and clear direction as to if the agent qualified for work at home and if not why. It did not require the end user to look through all of the data themselves to decide if the agent qualified or not and made it quite simple for them to understand.