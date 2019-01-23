import numpy as np
import pandas as pd
import statsmodels.api as sm
from matplotlib import pyplot as plt
import json

def get_frame():

    schedule = pd.read_excel('BPG - List of Centrally Scheduled Classrooms (CSC).xlsx')
    spring14 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_3349.xlsx', header = 1)
    fall14 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_2527.xlsx', header = 1)
    spring15 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_6241.xlsx', header = 1)
    fall15 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_4743.xlsx', header = 1)
    spring16 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_7859.xlsx', header = 1)
    fall16 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_1251.xlsx', header = 1)
    spring17 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_5453.xlsx', header = 1)
    fall17 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_1525.xlsx', header = 1)
    spring18 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_5390.xlsx', header = 1)
    fall18 = pd.read_excel('UA_SR_ALL_SEC_MAIN_TUCSON_1548.xlsx', header = 1)

    list_year = [spring14, fall14, spring15, fall15, spring16, fall16, spring17, fall17, spring18, fall18]

    return list_year

def drop_zero_enrol(years):

    for i in range(len(years)):
        for row in years[i].index:
            if years[i].loc[row, 'Tot Enrl'] == 0:
                years[i] = years[i].drop([row])
    
def classes_by_semester_dictionary(years):

    outer_dict = {}

    for i in range(len(years)):
        inner_dict = {}
        counts = years[i]['Course ID'].value_counts()
        for row in years[i].index:
            if years[i].loc[row, "Course ID"] in inner_dict.keys():
                inner_dict[years[i].loc[row,"Course ID"]] += years[i].loc[row,"Tot Enrl"]
            else:
                inner_dict[years[i].loc[row,"Course ID"]] = years[i].loc[row,"Tot Enrl"]
        for course in counts.index:
            inner_dict[course] = inner_dict[course] / counts[course]
        outer_dict[i] = inner_dict

    return outer_dict
    
def declining_increasing_classes(dict_class_by_semester):
    
    courses_list = []
    declining_class_index, increasing_class_index = [], []
    declining_class_val, increasing_class_val = [], []

    for year in sorted(dict_class_by_semester.keys()):
        for course in dict_class_by_semester[year].keys():
            if course not in courses_list:
                courses_list.append(course)
    
    for course in courses_list:
        tot_enrl_mean = []
        for year in sorted(dict_class_by_semester.keys()):
            if course in dict_class_by_semester[year].keys():
                tot_enrl_mean.append(dict_class_by_semester[year][course])
        if len(tot_enrl_mean) > 1:
            years_array = sm.add_constant(np.arange(len(tot_enrl_mean)))
            model = sm.OLS(tot_enrl_mean,years_array)
            result = model.fit()
            if result.params[1] < 0:
                declining_class_index.append(course)
                declining_class_val.append(result.params[1])
            elif result.params[1] > 0:
                increasing_class_index.append(course)
                increasing_class_val.append(result.params[1])

    dictionary = {}
    for course in increasing_class_index:
        tot = 0
        count = 0
        for year in sorted(dict_class_by_semester.keys()):
            if course in dict_class_by_semester[year].keys():
                tot += dict_class_by_semester[year][course]
                count += 1
        tot = tot / count
        dictionary[course] = tot

    dec_classes = pd.Series(declining_class_val, declining_class_index)
    inc_classes = pd.Series(increasing_class_val, increasing_class_index)

    return dec_classes, inc_classes, courses_list, dictionary

def inc_class_plot(inc_classes):

    files = get_frame()
    plt.figure()

    inc_classes = inc_classes.sort_values(ascending=False).head(10)
    course_name, course_num = [], []

    for course in inc_classes.index:
        for i in range(len(files)):
            for row in files[i].index:
                if files[i].loc[row, 'Course ID'] == course:
                    if course == files[i].loc[row, 'Course ID'] and files[i].loc[row,'Subject'] + str(files[i].loc[row,'Catalog']) not in course_name:
                        course_name.append(files[i].loc[row,'Subject'] + str(files[i].loc[row,'Catalog']))  

    new_inc_classes = pd.Series(inc_classes.values,course_name)
    new_inc_classes.plot(kind = "bar")

    ax = plt.gca()
    ax.set_xlabel('Name of Class', fontsize = 20)
    ax.set_ylabel('Slope for Class', fontsize = 20)
    ax.set_title("Top 10 Increasing classes", fontsize = 24)


def open_json_create_dict(filename, files_list):

    with open(filename) as fp:
        course_college = json.load(fp)

    dept = sorted(course_college.keys())
    colleges = list(set(course_college.values()))
    year_sub_dict = {}

    for i in range(len(files_list)):
        inner_dict = {}
        counts = files_list[i]['Subject'].value_counts()
        for row in files_list[i].index:
            if files_list[i].loc[row,'Subject'] not in inner_dict.keys():
                inner_dict[files_list[i].loc[row,'Subject']] = files_list[i].loc[row,'Tot Enrl']
            else:
                inner_dict[files_list[i].loc[row,'Subject']] += files_list[i].loc[row,'Tot Enrl']
        for course in counts.index:
            inner_dict[course] = inner_dict[course] / counts[course]
        year_sub_dict[i] = inner_dict

    yr_col_sub = {}
    for year in year_sub_dict.keys():
        by_college = {}
        for college in colleges:
            by_sub = {}
            for sub in dept:
                #by_sub = {}
                if course_college[sub] == college:
                    if sub in year_sub_dict[year].keys():
                        by_sub[sub] = year_sub_dict[year][sub]
            by_college[college] = by_sub
        yr_col_sub[year] = by_college
    return yr_col_sub

def declining_increasing_department(yr_col_sub_dict):

    dept_list = []
    declining_dept_index, increasing_dept_index = [], []
    declining_dept_val, increasing_dept_val = [], []

    for year in sorted(yr_col_sub_dict.keys()):
        for college in yr_col_sub_dict[year].keys():
            for dept in yr_col_sub_dict[year][college].keys():
                if dept not in dept_list:
                    dept_list.append(dept)
    for dept in dept_list:
        tot_enrl_mean = []
        for year in sorted(yr_col_sub_dict.keys()):
            for col in yr_col_sub_dict[year].keys():
                if dept in yr_col_sub_dict[year][col].keys():
                    tot_enrl_mean.append(yr_col_sub_dict[year][col][dept])
        if len(tot_enrl_mean) > 1:
            years_array = sm.add_constant(np.arange(len(tot_enrl_mean)))
            model = sm.OLS(tot_enrl_mean,years_array)
            result = model.fit()
            if result.params[1] < 0:
                declining_dept_index.append(dept)
                declining_dept_val.append(result.params[1])

            elif result.params[1] > 0:
                increasing_dept_index.append(dept)
                increasing_dept_val.append(result.params[1])
    dec_dept = pd.Series(declining_dept_val, declining_dept_index)
    inc_dept = pd.Series(increasing_dept_val, increasing_dept_index)

    college_list = []
    declining_college_index, increasing_college_index = [], []
    declining_college_val, increasing_college_val = [], []
    new_col_dict = {}
    yr_col, yr_tot = [], []
    for year in yr_col_sub_dict.keys():
        inner_dict = {}
        for college in yr_col_sub_dict[year].keys():
            tot = 0
            for sub in yr_col_sub_dict[year][college].keys():
                tot += yr_col_sub_dict[year][college][sub]
            count = len(yr_col_sub_dict[year][college].keys())
            if count != 0:
                tot = tot / len(yr_col_sub_dict[year][college].keys())
            inner_dict[college] = tot
        new_col_dict[year] = inner_dict

    for year in sorted(new_col_dict.keys()):
        for college in new_col_dict[year].keys():
            if college not in college_list:
                college_list.append(college)
    for college in college_list:
        tot_enrl_mean = []
        for year in sorted(new_col_dict.keys()):
            if college in new_col_dict[year].keys():
                tot_enrl_mean.append(new_col_dict[year][college])
        if len(tot_enrl_mean) > 1:
            years_array = sm.add_constant(np.arange(len(tot_enrl_mean)))
            model = sm.OLS(tot_enrl_mean,years_array)
            result = model.fit()
            if result.params[1] < 0:
                declining_college_index.append(college)
                declining_college_val.append(result.params[1])
            elif result.params[1] > 0:
                increasing_college_index.append(college)
                increasing_college_val.append(result.params[1])

    dec_college = pd.Series(declining_college_val, declining_college_index)
    inc_college = pd.Series(increasing_college_val, increasing_college_index).sort_values(ascending=False)

    plt_dict = {}
    for college in college_list:
        plt_list = []
        for year in new_col_dict.keys():
            if college in new_col_dict[year].keys():
                plt_list.append(new_col_dict[year][college])
        if len(plt_list) == 10:
            plt_dict[college] = plt_list
    
    return inc_dept, dec_dept, inc_college, dec_college, plt_dict
    print(inc_dept, dec_dept, inc_college, dec_college, plt_dict)
    
def main():
    
    list_year = get_frame()
    drop_zero_enrol(list_year)
    decline_classes_dict = classes_by_semester_dictionary(list_year)
    decline_classes, increase_classes, courses_lis, dictionary_incr = declining_increasing_classes(decline_classes_dict)
    yr_col_dep_dict = open_json_create_dict('colleges.json',list_year)
    in_dept, de_dept, in_col, de_col, plt_dict = declining_increasing_department(yr_col_dep_dict)
    print(inc_dept, dec_dept, inc_college, dec_college, plt_dict)

main()