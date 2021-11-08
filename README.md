[![Run on Repl.it](https://repl.it/badge/github/jscacco/wpn_report_generator)](https://repl.it/github/jscacco/wpn_report_generator)

# Welcome!

This program automates the generation of monthly WPN Premium reports using raw Lightspeed line reports. Although it is specific to one store (Fair Game) and one service, the general program could certainly be adapted for use by a broader clientele.

Please refer to wpn_report_generator.py for the program itself. 
files/wotc_sku_dict.txt contains product information needed to fill out each report. 
A blank version of these reports can be found at files/FairGame_POSData_TEMPLATE.xlsx.

## Instructions
0) Navigate to the folder containing wpn_report_generator.py ('cd ~/WPN/wpn_report_generator', or something like that)
1) To prepare the input file, download the line report from Lightspeed (filtering by WotC and store) and save it as an excel workbook (.xlsx file extension). Then, move this file into the same folder as the project (Eric - this folder is titled 'WPN' and is pinned on the left panel in File Explorer)
2) To prepare the output file, copy+paste the template file in /files/ to the main directory and rename it by replacing the capitals with the relevant info.
3) Now is a good time to pull up the SKU master list from WotC - you will need to reference it while the program runs.
4) Finally, we can run the program. Type 'py wpn_report_generator.py -h' for help on templating. Here are some helpful reminders: 1) The first flag is the letter L, not the number 1. 2) Each flag is preceded by a minus sign '-', not an equal sign '='. 3) You can auto-complete file names by pressing tab. This is easier if the folder contains only those files you need.
