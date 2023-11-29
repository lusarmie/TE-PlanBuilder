# TE-PlanBuilder
Python Script to fill the Implementation Plan 2023 document with the information related to what is currently configured on TE platform.

## Info Provided
The script will pull the info for Agents (Enterprise and Endpoint), tests, %Unit consumption, customer name and date of doc creation.

## How to run?
1. Make sure you are added to the Account Group with Org Admin permissions
2. Clone this repo to your computer
3. Run python script
   ```
   python3 TE-PlanBuilder.py
   ```
4. Script will ask for your token, provide it
5. After a few seconds, it will show a complete list of the available Account Groups. Type (or even better copy & paste) the name of the AG you are interested on.
6. After a few seconds/minutes, the script will exit and the file will be located under TE-PlanBuilder dir.
   `<orgName>_ImplementationPlan.xlsx`

## Author

`lusarmie@cisco.com`
