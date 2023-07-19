# Scorecard_Outliers
## Overview
The department was moving away from a relative stack rank that was used to coach agents. In place of that, they were given their "Scorecard." This is a method to guage agent performance against the company goals, which are reset every quarter. The aim of this was to allow agents to become more collabrative rather than competitive. Agent's were able to check their own scorecard progress through access of a web app, however, the aim of this report was to give visibility to the agent's scorecard to the agent's reporting hierarchy in order to monitor agent performance month over month in an Excel platform as requested. 

## Scorecard System
* Every quarter, the department would get performance goals for the metrics along with their weights. 
* There were three bands that an agent would be able to hit: 
    * Green if they were meeting goal 
    * Yellow if they were below goal but close 
    * Red if they were missing goal entirely
* There was also a fourth color, blue, that if the agent would hit above goal at the decided level, as well as hit a few other qualifications, the agent would be eligible for promotion. 

* There were the Red, Yellow, Green for each metric as well as Blue for those that was examined for promotion. 
* Each color would correlate to a number
    * Red = 1
    * Yellow = 2
    * Green = 3
    * Blue = 3
* These scores would then be multiplied by the weight for the metric, and the products then summed to get the "Overall Score" which could range from 1-3.
* This overall score was then rounded to a whole number with .5 rounding up and the number was correlated with a color above, giving the agent their overall color. 
* If the agent hit a minimum of 200 calls and 80 hours staffed as well as hitting the blue tier in the metrics that had them, then that agent would be Overall Blue. 4 months in a row of this, and the agent would be eligible for promotion. Or the agent could also achieve it 6 times in a 9 month lookback to achieve the same effect.

## Methodology
The script starts by calculating the current fiscal month utilizing the the custom library before launching a Selenium instance and downloading a custom report from Microstrategy using their URL API. The method would answer the different prompts depending on the parameters fed to the method. 

The script would then grab the remaining required data from the Unit MSSQL Server and existing excel files maintained in the script folder. 

Once all of the data is then loaded into Pandas Dataframes, cleaned, and joined, the script uses three tiered nested for loop in order to iterate through each of the fiscal months that new data was pulled for, each of the titles that have different goals, and each of the metrics examined. Within each of these, it grabs the weights for the metrics and thresholds for the current fiscal month and title, then using .loc, finds those that hit the band and create a column to store the color. 

Then once color is assigned, there is a loop that goes through a dictionary and multiplies the color score by the weight and creates a column for that. 

The script then sums all score columns to create the weighted score column and to create a color score column, it rounds the overall scores in the method mentioned above. 

Following that, the script creates a list of metrics that are used for promotion. Using that list, it mirrors the main dataframe, then loops through each of those metrics and filters out anyone who does not have "Blue" in that metric's columns. Then it filters out those who did not meet the call requirement and hours staffed requirement, and for those remaining, it creates a True flaged column called Level Up. This column is then joined with the main dataframe and those who have True in that column, their "Overall Color" is changed to blue, and the column is dropped.

The prior overall scorecard data is loaded into a seperate pandas dataframe where the new data, or refreshed data is replaced, and then saved again in an excel sheet. Then the months as a part of the rolling 9 month lookback are kept and the overall color column is pivoted with the months as columns, rows are employee numbers, and values are the overall color. This is then joined with the scorecard dataframe before exported as a datafile which is used in a power query to bring in the data and utilize conditional formatting to display all formatting and lookback in company color guidelines. 

Example of report layout without identifying agent data:
![example screenshot](https://github.com/aKnownSaltMine/Scorecard_Outliers/blob/main/screenshots/report_screenshot.png)

## Summary
This became one of our most utilized and asked for reports. The feedback recieved from the end user was extremely positive and aimed at expansion to apply all the way up the agent reporting hierarchy. And has assisted heavily in the department's move away from a relative stack rank to one focused around company goals. 