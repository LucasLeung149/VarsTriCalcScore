# VarsTriCalcScore
 This is a Python code to calculate Varsity duathlon/triathlon scores.
 The current release is VarsTriCalcScore.v1.0, updated on 28 May 2023.

## How to use the program
This program is set up as a short code to download on a local computer with a python installation to compute varsity triathlon and duathlon scores.
To use the program follow the following steps.
1. Download the file and unzip the package.
2. Edit two excel files, one for men (open) and one for women. The excel files must have a .xlsx ending and must contain only three columns: Name & Time & Allegience(i.e. either Oxf/Cam). The first line should be left as titles (i.e. the line Name & Time & Allegience). The results should start from the second line. The required format can be found by looking at the two test files - test_results_men.xlsx and test_results_women.xlsx.
3. Put the two excel files in the file containing the .py file.
4. Run the terminal. You should run the following lines:
   cd <user>/.../<script_directory>
   ls
   (At this point - make sure you can see the .py file and the .xlsx files you have just edited and put in.)
   python3 VarsTriCalcScore.py
5. When prompted, enter the file names as requested. Make sure you enter the .xlsx ending for both of the files.
6. The program should display the results as computed.

The scores have yet to be introduced (I don't know how the scoring system works). This will be introduced in a later version.
