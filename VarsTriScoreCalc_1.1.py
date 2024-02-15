# !/usr/bin/env python
# coding: utf-8

# # Below is a code for computing the Varsity scores.


import numpy as np
import pandas as pd
import datetime

# Settings
dir_excel_men = "results_open_duathlon_24.xlsx"
dir_excel_women = "results_women_duathlon_24.xlsx"
students_rule = True
matric_on_tri = False


#Intro:
print("Welcome to VarsTriScoreCalc.v1.0. This programme is developed by Lucas Leung and last revised in 28 May 2023.")
print("The programme calculates the results for Varsity Dualathlon and Triathlon for both men and women.")
print("The programme takes data from an Excel file with columns Name, Time and Allegiance and outputs the results.")
print("")



#Enter file name:
"""
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
"""

#Import results
df_men = pd.read_excel(dir_excel_men)
df_women = pd.read_excel(dir_excel_women)
resultsarr_men = np.asarray(df_men)
resultsarr_women = np.asarray(df_women)



# Define extract functions
def extract_oxf(results_array):
    #extract all oxford people from array
    oxfarray = []
    oxfarrayboo = []
    allegarr = results_array[:,2]
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
    allegarr = results_array[:,2]
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
    

#Extract teams with additional rules
def extract_team_rules_duathlon(results_oxford_array,results_cambridge_array):
    ## first sort the arrays with time as keys
    #sorted_oxf = sorted(results_oxford_array,key=lambda x: x[1])
    #sorted_tab = sorted(results_cambridge_array,key=lambda x: x[1])
    ## sort arrays - first filter to get student array
    oxf_stu = [row for row in results_oxford_array if row[3]]
    sorted_oxf_stu = sorted(oxf_stu,key=lambda x: x[1])
    tab_stu = [row for row in results_cambridge_array if row[3]]
    sorted_tab_stu = sorted(tab_stu,key=lambda x: x[1])
    ## extract the 1st and 2nd teams
    oxf1st = np.hstack(sorted_oxf_stu[0:6])
    oxf2nd = np.hstack(sorted_oxf_stu[6:9])
    tab1st = np.hstack(sorted_tab_stu[0:6])
    tab2nd = np.hstack(sorted_tab_stu[6:9])
    mob_oxf_arr = sorted(results_oxford_array,key=lambda x: x[1])
    mob_tab_arr = sorted(results_cambridge_array,key=lambda x: x[1])
    for row in sorted_oxf_stu[0:9]:
        index_to_delete = np.where(np.all(mob_oxf_arr == row, axis=1))
        mob_oxf_arr = np.delete(mob_oxf_arr,index_to_delete, axis=0)
    for row in sorted_tab_stu[0:9]:
        index_to_delete = np.where(np.all(mob_tab_arr == row, axis=1))
        mob_tab_arr = np.delete(mob_tab_arr,index_to_delete, axis=0)
    ## now sort the mob teams
    oxf_num_mob = len(mob_oxf_arr)
    tab_num_mob = len(mob_tab_arr)
    num_mob = min(oxf_num_mob,tab_num_mob)
    ## sort mob
    sorted_oxf_mob = sorted(mob_oxf_arr,key=lambda x: x[1])
    sorted_tab_mob = sorted(mob_tab_arr,key=lambda x: x[1])
    oxfmob = np.hstack(sorted_oxf_mob[0:num_mob])
    tabmob = np.hstack(sorted_tab_mob[0:num_mob])
    return oxf1st.reshape(6,-1),oxf2nd.reshape(3,-1),oxfmob.reshape(num_mob,-1),tab1st.reshape(6,-1),tab2nd.reshape(3,-1),tabmob.reshape(num_mob,-1)

def extract_team_rules_triathlon(results_oxford_array,results_cambridge_array):
    ## first sort the arrays with time as keys
    #sorted_oxf = sorted(results_oxford_array,key=lambda x: x[1])
    #sorted_tab = sorted(results_cambridge_array,key=lambda x: x[1])
    ## sort arrays - first filter to get matriculated student array
    oxf_stu_mat = [row for row in results_oxford_array if row[3] and row[4]]
    sorted_oxf_stu_mat = sorted(oxf_stu_mat,key=lambda x: x[1])
    tab_stu_mat = [row for row in results_cambridge_array if row[3] and row[4]]
    sorted_tab_stu_mat = sorted(tab_stu_mat,key=lambda x: x[1])
    ## extract the 1st teams
    oxf1st = np.hstack(sorted_oxf_stu_mat[0:6])
    tab1st = np.hstack(sorted_tab_stu_mat[0:6])  
    ## delete 1sts
    mobsec_oxf_arr = sorted(results_oxford_array,key=lambda x: x[1])
    mobsec_tab_arr = sorted(results_cambridge_array,key=lambda x: x[1])
    for row in sorted_oxf_stu_mat[0:6]:
        index_to_delete = np.where(np.all(mobsec_oxf_arr == row, axis=1))
        mobsec_oxf_arr = np.delete(mobsec_oxf_arr,index_to_delete, axis=0)
    for row in sorted_tab_stu_mat[0:6]:
        index_to_delete = np.where(np.all(mobsec_tab_arr == row, axis=1))
        mobsec_tab_arr = np.delete(mobsec_tab_arr,index_to_delete, axis=0)
    ## extract 2nd teams
    oxf_stu = [row for row in mobsec_oxf_arr if row[3]]
    sorted_oxf_stu = sorted(oxf_stu,key=lambda x: x[1])
    tab_stu = [row for row in mobsec_tab_arr if row[3]]
    sorted_tab_stu = sorted(tab_stu,key=lambda x: x[1])
    oxf2nd = np.hstack(sorted_oxf_stu[0:3])
    tab2nd = np.hstack(sorted_tab_stu[0:3])
    ## delete 2nds
    mob_oxf_arr = sorted(mobsec_oxf_arr,key=lambda x: x[1])
    mob_tab_arr = sorted(mobsec_tab_arr,key=lambda x: x[1])
    for row in sorted_oxf_stu[0:3]:
        index_to_delete = np.where(np.all(mob_oxf_arr == row, axis=1))
        mob_oxf_arr = np.delete(mob_oxf_arr,index_to_delete, axis=0)
    for row in sorted_tab_stu[0:3]:
        index_to_delete = np.where(np.all(mob_tab_arr == row, axis=1))
        mob_tab_arr = np.delete(mob_tab_arr,index_to_delete, axis=0)
    ## now sort the mob teams
    oxf_num_mob = len(mob_oxf_arr)
    tab_num_mob = len(mob_tab_arr)
    num_mob = min(oxf_num_mob,tab_num_mob)
    ## sort mob
    sorted_oxf_mob = sorted(mob_oxf_arr,key=lambda x: x[1])
    sorted_tab_mob = sorted(mob_tab_arr,key=lambda x: x[1])
    oxfmob = np.hstack(sorted_oxf_mob[0:num_mob])
    tabmob = np.hstack(sorted_tab_mob[0:num_mob])
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
        return "Cambridge",tab_time,oxf_time,(oxf_time-tab_time)
    if oxf_time < tab_time:
        return "Oxford",oxf_time,tab_time,(tab_time-oxf_time)
    else:
        return "Draw",tab_time,oxf_time,(oxf_time-tab_time)




#Compute scores - men
oxford_all_men = extract_oxf(resultsarr_men)
cambridge_all_men = extract_tab(resultsarr_men)
if students_rule:
    if matric_on_tri:
        all_teams_men = extract_team_rules_triathlon(oxford_all_men,cambridge_all_men)
    else:
        all_teams_men = extract_team_rules_duathlon(oxford_all_men,cambridge_all_men)
else:
    all_teams_men = extract_team(oxford_all_men,cambridge_all_men)

oxf_1st_time_men = compute_total_time(all_teams_men[0])
oxf_2nd_time_men = compute_total_time(all_teams_men[1])
oxf_mob_time_men = compute_total_time(all_teams_men[2])
cam_1st_time_men = compute_total_time(all_teams_men[3])
cam_2nd_time_men = compute_total_time(all_teams_men[4])
cam_mob_time_men = compute_total_time(all_teams_men[5])

oxf_1st_team_men = display_team(all_teams_men[0])
oxf_2nd_team_men = display_team(all_teams_men[1])
oxf_mob_team_men = display_team(all_teams_men[2])
cam_1st_team_men = display_team(all_teams_men[3])
cam_2nd_team_men = display_team(all_teams_men[4])
cam_mob_team_men = display_team(all_teams_men[5])



win_team_1st_men = det_win_team(oxf_1st_time_men,cam_1st_time_men)
win_team_2nd_men = det_win_team(oxf_2nd_time_men,cam_2nd_time_men)
win_team_mob_men = det_win_team(oxf_mob_time_men,cam_mob_time_men)

num_mob_men = min(len(all_teams_men[2]),len(all_teams_men[5]))

#Compute scores - women
oxford_all_women = extract_oxf(resultsarr_women)
cambridge_all_women = extract_tab(resultsarr_women)
all_teams_women = extract_teams(oxford_all_women,cambridge_all_women)

if students_rule:
    if matric_on_tri:
        all_teams_women = extract_team_rules_triathlon(oxford_all_women,cambridge_all_women)
    else:
        all_teams_women = extract_team_rules_duathlon(oxford_all_women,cambridge_all_women)
else:
    all_teams_women = extract_team(oxford_all_women,cambridge_all_women)

oxf_1st_time_women = compute_total_time(all_teams_women[0])
oxf_2nd_time_women = compute_total_time(all_teams_women[1])
oxf_mob_time_women = compute_total_time(all_teams_women[2])
cam_1st_time_women = compute_total_time(all_teams_women[3])
cam_2nd_time_women = compute_total_time(all_teams_women[4])
cam_mob_time_women = compute_total_time(all_teams_women[5])

oxf_1st_team_women = display_team(all_teams_women[0])
oxf_2nd_team_women = display_team(all_teams_women[1])
oxf_mob_team_women = display_team(all_teams_women[2])
cam_1st_team_women = display_team(all_teams_women[3])
cam_2nd_team_women = display_team(all_teams_women[4])
cam_mob_team_women = display_team(all_teams_women[5])


win_team_1st_women = det_win_team(oxf_1st_time_women,cam_1st_time_women)
win_team_2nd_women = det_win_team(oxf_2nd_time_women,cam_2nd_time_women)
win_team_mob_women = det_win_team(oxf_mob_time_women,cam_mob_time_women)

num_mob_women = min(len(all_teams_women[2]),len(all_teams_women[5]))

#Determine individual results
all_results_men = sorted(resultsarr_men,key=lambda x: x[1])
top_3_men = np.hstack(all_results_men)[0:15]
top_3_data_men = top_3_men.reshape(3,-1)

all_results_women = sorted(resultsarr_women,key=lambda x: x[1])
top_3_women = np.hstack(all_results_women)[0:15]
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
print("With a total winning time of",display_timedelta(win_team_1st_women[1]),"the team that won Women's First is",win_team_1st_women[0],"!")
print("The members of their team are:")
if win_team_1st_women[0]=='Cambridge':
    print(cam_1st_team_women)
    print("The Oxford team members are:",oxf_1st_team_women)
    print("Their finishing time was",display_timedelta(win_team_1st_women[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_1st_women[3]),".")
if win_team_1st_women[0]=='Oxford':
    print(oxf_1st_team_women)
    print("The Cambridge team members are:",cam_1st_team_women)
    print("Their finishing time was",display_timedelta(win_team_1st_women[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_1st_women[3]),".")
print("")
print("Now for Women's Second team.")
print("With a total winning time of",display_timedelta(win_team_2nd_women[1]),"the team that won Women's Second is",win_team_2nd_women[0],"!")
print("The members of their team are:")
if win_team_2nd_women[0]=='Cambridge':
    print(cam_2nd_team_women)
    print("The Oxford team members are:",oxf_2nd_team_women)
    print("Their finishing time was",display_timedelta(win_team_2nd_women[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_2nd_women[3]),".")
if win_team_2nd_women[0]=='Oxford':
    print(oxf_2nd_team_women)
    print("The Cambridge team members are:",cam_2nd_team_women)
    print("Their finishing time was",display_timedelta(win_team_2nd_women[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_2nd_women[3]),".")
print("")
print("Now for Women's mob team...")
print("The number of mobs counted is:",num_mob_women,".")
print("With a total winning time of",display_timedelta(win_team_mob_women[1]),"the team that won Women's Mob is",win_team_mob_women[0],"!")
print("The members of their team are:")
if win_team_mob_women[0]=='Cambridge':
    print(cam_mob_team_women)
    print("The Oxford team members are:",oxf_mob_team_women)
    print("Their finishing time was",display_timedelta(win_team_mob_women[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_mob_women[3]),".")
if win_team_mob_women[0]=='Oxford':
    print(oxf_mob_team_women)
    print("The Cambridge team members are:",cam_mob_team_women)
    print("Their finishing time was",display_timedelta(win_team_mob_women[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_mob_women[3]),".")
print("")
print("")
print("Now for the men!")
print("We will first announce the results of the Men's First team.")
print("With a total winning time of",display_timedelta(win_team_1st_men[1]),"the team that won Men's First is",win_team_1st_men[0],"!")
print("The members of their team are:")
if win_team_1st_men[0]=='Cambridge':
    print(cam_1st_team_men)
    print("The Oxford team members are:",oxf_1st_team_men)
    print("Their finishing time was",display_timedelta(win_team_1st_men[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_1st_men[3]),".")
if win_team_1st_men[0]=='Oxford':
    print(oxf_1st_team_men)
    print("The Cambridge team members are:",cam_1st_team_men)
    print("Their finishing time was",display_timedelta(win_team_1st_men[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_1st_men[3]),".")
print("")
print("Now for Men's Second team.")
print("With a total winning time of",display_timedelta(win_team_2nd_men[1]),"the team that won Men's Second is",win_team_2nd_men[0],"!")
print("The members of their team are:")
if win_team_2nd_men[0]=='Cambridge':
    print(cam_2nd_team_men)
    print("The Oxford team members are:",oxf_2nd_team_men)
    print("Their finishing time was",display_timedelta(win_team_2nd_men[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_2nd_men[3]),".")
if win_team_2nd_men[0]=='Oxford':
    print(oxf_2nd_team_men)
    print("The Cambridge team members are:",cam_2nd_team_men)
    print("Their finishing time was",display_timedelta(win_team_2nd_men[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_2nd_men[3]),".")
print("")
print("Now for Men's mob team...")
print("The number of mobs counted is:",num_mob_men,".")
print("With a total winning time of",display_timedelta(win_team_mob_men[1]),"the team that won Men's Mob is",win_team_mob_men[0],"!")
print("The members of their team are:")
if win_team_mob_men[0]=='Cambridge':
    print(cam_mob_team_men)
    print("The Oxford team members are:",oxf_mob_team_men)
    print("Their finishing time was",display_timedelta(win_team_mob_men[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_mob_men[3]),".")
if win_team_mob_men[0]=='Oxford':
    print(oxf_mob_team_men)
    print("The Cambridge team members are:",cam_mob_team_men)
    print("Their finishing time was",display_timedelta(win_team_mob_men[2]),".")
    print("The time difference between the two teams was",display_timedelta(win_team_mob_men[3]),".")
print("")
print("")
print("This is it! Thank you for using VarsTriScoreCalc.v1.2!")
