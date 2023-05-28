#!/usr/bin/env python
# coding: utf-8

# # Below is a code for computing the Varsity scores.


import numpy as np
import pandas as pd
import datetime


#Intro:
print("Welcome to VarsTriScoreCalc.v1.0. This programme is developed by Lucas Leung and last revised in 28 May 2023.")
print("The programme calculates the results for Varsity Dualathlon and Triathlon for both men and women.")
print("The programme takes data from an Excel file with columns Name, Time and Allegiance and outputs the results.")
print("")



#Enter file name:
print("Please enter the excel file name for the men results (and put it in the same directory as the script).")
dir_excel_men = input("File Name:")
if dir_excel_men.endswith(".xlsx"):
    print("Thank you!")
else:
    print("Wrong file type. Ending script.")
    quit()

print("Please enter the excel file name for the women results (and put it in the same directory as the script).")
dir_excel_women = input("File Name:")
if dir_excel_women.endswith(".xlsx"):
    print("Thank you! Now calculating scores...")
else:
    print("Wrong file type. Ending script.")
    quit()

#Import results
df_men = pd.read_excel(dir_excel_men)
df_women = pd.read_excel(dir_excel_women)
resultsarr_men = np.asarray(df_men)
resultsarr_women = np.asarray(df_women)


# Define extract functions,
def extract_oxf(results_array):
    #extract all oxford people from array
    oxfarray = []
    oxfarrayboo = []
    allegarr = results_array[:,-1]
    for element in allegarr:
        if element == 'Oxford':
            oxfarrayboo.append(True)
        else:
            oxfarrayboo.append(False)
    return results_array[oxfarrayboo]

def extract_tab(results_array):
    #extract all cambridge people from array
    tabarray = []
    tabarrayboo = []
    allegarr = results_array[:,-1]
    for element in allegarr:
        if element == 'Cambridge':
            tabarrayboo.append(True)
        else:
            tabarrayboo.append(False)
    return results_array[tabarrayboo]


#Extract 1st, 2nd and Mob Teams
def extract_oxf1st2nd(results_oxford_array):
    sorted_oxf = sorted(results_oxford_array,key=lambda x: x[1])
    oxf1st = np.hstack(sorted_oxf[0:6])
    oxf2nd = np.hstack(sorted_oxf[6:9])
    return oxf1st.reshape(6,-1),oxf2nd.reshape(3,-1)

def extract_tab1st2nd(results_cambridge_array):
    sorted_tab = sorted(results_cambridge_array,key=lambda x: x[1])
    tab1st = np.hstack(sorted_tab[0:6])
    tab2nd = np.hstack(sorted_tab[6:9])
    return tab1st.reshape(6,-1),tab2nd.reshape(3,-1)

def extract_mobs(results_oxford_array,results_cambridge_array):
    sorted_oxf = sorted(results_oxford_array,key=lambda x: x[1])
    sorted_tab = sorted(results_cambridge_array,key=lambda x: x[1])
    oxf_num_mob = len(results_oxford_array)-9
    tab_num_mob = len(results_cambridge_array)-9
    num_mob = min(oxf_num_mob,tab_num_mob)
    oxf_mob = np.hstack(sorted_oxf[9:9+num_mob])
    tab_mob = np.hstack(sorted_tab[9:9+num_mob])
    return oxf_mob.reshape(num_mob,-1),tab_mob.reshape(num_mob,-1)

def extract_teams(results_oxford_array,results_cambridge_array):
    sorted_oxf = sorted(results_oxford_array,key=lambda x: x[1])
    sorted_tab = sorted(results_cambridge_array,key=lambda x: x[1])
    oxf1st = np.hstack(sorted_oxf[0:6])
    oxf2nd = np.hstack(sorted_oxf[6:9])
    tab1st = np.hstack(sorted_tab[0:6])
    tab2nd = np.hstack(sorted_tab[6:9])
    oxf_num_mob = len(results_oxford_array)-9
    tab_num_mob = len(results_cambridge_array)-9
    num_mob = min(oxf_num_mob,tab_num_mob)
    oxfmob = np.hstack(sorted_oxf[9:9+num_mob])
    tabmob = np.hstack(sorted_tab[9:9+num_mob])
    return oxf1st.reshape(6,-1),oxf2nd.reshape(3,-1),oxfmob.reshape(num_mob,-1),tab1st.reshape(6,-1),tab2nd.reshape(3,-1),tabmob.reshape(num_mob,-1)
    

#Compute total time in each team
def convert_time_timedelta(x):
    y = datetime.timedelta(hours=x.hour, minutes=x.minute, seconds=x.second, microseconds=x.microsecond)
    return y


def compute_total_time(results_team):
    times_array = results_team[:,1]
    times_array_conv = list(map(convert_time_timedelta,times_array))
    tot_time = sum(times_array_conv, datetime.timedelta())
    return tot_time

def display_timedelta(td):
    minutes, seconds = divmod(td.seconds + td.days * 86400, 60)
    hours, minutes = divmod(minutes, 60)
    return '{:d}:{:02d}:{:02d}'.format(hours, minutes, seconds)

def display_team(results_team):
    team_members = results_team[:,0]
    return team_members

#Determine winning teams
def det_win_team(oxf_time,tab_time):
    if oxf_time > tab_time:
        return "Cambridge",tab_time
    if oxf_time < tab_time:
        return "Oxford",oxf_time
    else:
        return "Draw",tab_time


#Compute scores - men
oxford_all_men = extract_oxf(resultsarr_men)
cambridge_all_men = extract_tab(resultsarr_men)
all_teams_men = extract_teams(oxford_all_men,cambridge_all_men)

oxf_1st_time_men = compute_total_time(all_teams_men[0])
oxf_2nd_time_men = compute_total_time(all_teams_men[1])
oxf_mob_time_men = compute_total_time(all_teams_men[2])
cam_1st_time_men = compute_total_time(all_teams_men[3])
cam_2nd_time_men = compute_total_time(all_teams_men[4])
cam_mob_time_men = compute_total_time(all_teams_men[5])

oxf_1st_team_men = display_team(all_teams_men[0])
oxf_2nd_team_men = display_team(all_teams_men[1])
cam_1st_team_men = display_team(all_teams_men[3])
cam_2nd_team_men = display_team(all_teams_men[4])


win_team_1st_men = det_win_team(oxf_1st_time_men,cam_1st_time_men)
win_team_2nd_men = det_win_team(oxf_2nd_time_men,cam_2nd_time_men)
win_team_mob_men = det_win_team(oxf_mob_time_men,cam_mob_time_men)

num_mob_men = min(len(all_teams_men[2]),len(all_teams_men[5]))

#Compute scores - women
oxford_all_women = extract_oxf(resultsarr_women)
cambridge_all_women = extract_tab(resultsarr_women)
all_teams_women = extract_teams(oxford_all_women,cambridge_all_women)

oxf_1st_time_women = compute_total_time(all_teams_women[0])
oxf_2nd_time_women = compute_total_time(all_teams_women[1])
oxf_mob_time_women = compute_total_time(all_teams_women[2])
cam_1st_time_women = compute_total_time(all_teams_women[3])
cam_2nd_time_women = compute_total_time(all_teams_women[4])
cam_mob_time_women = compute_total_time(all_teams_women[5])

oxf_1st_team_women = display_team(all_teams_women[0])
oxf_2nd_team_women = display_team(all_teams_women[1])
cam_1st_team_women = display_team(all_teams_women[3])
cam_2nd_team_women = display_team(all_teams_women[4])


win_team_1st_women = det_win_team(oxf_1st_time_women,cam_1st_time_women)
win_team_2nd_women = det_win_team(oxf_2nd_time_women,cam_2nd_time_women)
win_team_mob_women = det_win_team(oxf_mob_time_women,cam_mob_time_women)

num_mob_women = min(len(all_teams_women[2]),len(all_teams_women[5]))

#Determine individual results
all_results_men = sorted(resultsarr_men,key=lambda x: x[1])
top_3_men = np.hstack(all_results_men)[0:9]
top_3_data_men = top_3_men.reshape(3,-1)

all_results_women = sorted(resultsarr_women,key=lambda x: x[1])
top_3_women = np.hstack(all_results_women)[0:9]
top_3_data_women = top_3_women.reshape(3,-1)



#Printing results.
print("It is my pleasure and honour to now announce the results for the competition.")

print("")
print("First, the individual results.")
print("We begin with the women.")
print("The woman who came first is:",top_3_data_women[0,0], "from",top_3_data_women[0,2], "with a time of",display_timedelta(convert_time_timedelta(top_3_data_women[0,1])),"!")
print("The woman who came second is:",top_3_data_women[1,0], "from",top_3_data_women[1,2], "with a time of",display_timedelta(convert_time_timedelta(top_3_data_women[1,1])),"!")
print("The woman who came third is:",top_3_data_women[2,0], "from",top_3_data_women[2,2], "with a time of",display_timedelta(convert_time_timedelta(top_3_data_women[2,1])),"!")
print("Now for the men.")
print("The man who came first is:",top_3_data_men[0,0], "from",top_3_data_men[0,2], "with a time of",display_timedelta(convert_time_timedelta(top_3_data_men[0,1])),"!")
print("The man who came second is:",top_3_data_men[1,0], "from",top_3_data_men[1,2], "with a time of",display_timedelta(convert_time_timedelta(top_3_data_men[1,1])),"!")
print("The man who came third is:",top_3_data_men[2,0], "from",top_3_data_men[2,2], "with a time of",display_timedelta(convert_time_timedelta(top_3_data_men[2,1])),"!")
print("Let us congratulate them!")
print("")

print("Now team results.")
print("")
print("We begin with the women again.")
print("We will first announce the results of the Women's First team.")
print("With a total winning time of",display_timedelta(win_team_1st_women[1]),"the team that won Men's First is",win_team_1st_women[0],"!")
print("The members of their team are:")
if win_team_1st_women[0]=='Cambridge':
    print(cam_1st_team_women)
    print("The Oxford team members are:",oxf_1st_team_women)
if win_team_1st_women[0]=='Oxford':
    print(oxf_1st_team_women)
    print("The Cambridge team members are:",cam_1st_team_women)
print("")
print("Now for Women's Second team.")
print("With a total winning time of",display_timedelta(win_team_2nd_women[1]),"the team that won Men's Second is",win_team_2nd_women[0],"!")
print("The members of their team are:")
if win_team_1st_women[0]=='Cambridge':
    print(cam_2nd_team_women)
    print("The Oxford team members are:",oxf_2nd_team_women)
if win_team_1st_women[0]=='Oxford':
    print(oxf_2nd_team_women)
    print("The Cambridge team members are:",cam_2nd_team_women)
print("")
print("Now for Women's mob team...")
print("The number of mobs counted is:",num_mob_women,".")
print("With a total winning time of",display_timedelta(win_team_mob_women[1]),"the team that won Men's Mob is",win_team_mob_women[0],"!")
print("")
print("Now for the men!")
print("We will first announce the results of the Men's First team.")
print("With a total winning time of",display_timedelta(win_team_1st_men[1]),"the team that won Men's First is",win_team_1st_men[0],"!")
print("The members of their team are:")
if win_team_1st_men[0]=='Cambridge':
    print(cam_1st_team_men)
    print("The Oxford team members are:",oxf_1st_team_men)
if win_team_1st_men[0]=='Oxford':
    print(oxf_1st_team_men)
    print("The Cambridge team members are:",cam_1st_team_men)
print("")
print("Now for Men's Second team.")
print("With a total winning time of",display_timedelta(win_team_2nd_men[1]),"the team that won Men's Second is",win_team_2nd_men[0],"!")
print("The members of their team are:")
if win_team_1st_men[0]=='Cambridge':
    print(cam_2nd_team_men)
    print("The Oxford team members are:",oxf_2nd_team_men)
if win_team_1st_men[0]=='Oxford':
    print(oxf_2nd_team_men)
    print("The Cambridge team members are:",cam_2nd_team_men)
print("")
print("Now for Men's mob team...")
print("The number of mobs counted is:",num_mob_women,".")
print("With a total winning time of",display_timedelta(win_team_mob_men[1]),"the team that won Men's Mob is",win_team_mob_men[0],"!")
print("")
print("This is it! Thank you for using VarsTriScoreCalc.v1.0!")

