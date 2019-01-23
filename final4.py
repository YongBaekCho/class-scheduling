import numpy as np
import pandas as pd
import statsmodels.api as sm
import json
import matplotlib.pyplot as plt


def create_df():
    
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
    
    first = []
    second = []
    third = []
    fourth = []
    fifth = []
    sixth = []
    seventh = []
    eighth = []
    a = []
    b = []
    c = []
    d = []
    e = []
    f = []
    g = []
    h = []
    i = []
    j = []
    k = []
    l = []

    for row in spring14.index:
        if spring14.loc[row, 'Component'] == 'LEC':
            first.append(spring14.loc[row, 'Tot Enrl'])
        elif spring14.loc[row, 'Component'] == 'LAB':
            second.append(spring14.loc[row, 'Tot Enrl'])
    mean_lec_sp14 = (sum(first) / len(first))    
    mean_lab_sp14 = (sum(second) / len(second))

    for row in fall14.index:
        if fall14.loc[row, 'Component'] == 'LEC':
            third.append(fall14.loc[row, 'Tot Enrl'])
        elif fall14.loc[row, 'Component'] == 'LAB':
            fourth.append(fall14.loc[row, 'Tot Enrl'])
    mean_lec_fall14 = (sum(third) / len(third))    
    mean_lab_fall14 = (sum(fourth) / len(fourth))

    for row in spring15.index:
        if spring15.loc[row, 'Component'] == 'LEC':
            fifth.append(spring15.loc[row, 'Tot Enrl'])
        elif spring15.loc[row, 'Component'] == 'LAB':
            sixth.append(spring15.loc[row, 'Tot Enrl'])
    mean_lec_sp15 = (sum(fifth) / len(fifth))    
    mean_lab_sp15 = (sum(sixth) / len(sixth))

    for row in fall15.index:
        if fall15.loc[row, 'Component'] == 'LEC':
            seventh.append(fall15.loc[row, 'Tot Enrl'])
        elif fall15.loc[row, 'Component'] == 'LAB':
            eighth.append(fall15.loc[row, 'Tot Enrl'])
    mean_lec_fall15 = (sum(seventh) / len(seventh))    
    mean_lab_fall15 = (sum(eighth) / len(eighth))

    for row in spring16.index:
        if spring16.loc[row, 'Component'] == 'LEC':
            a.append(spring16.loc[row, 'Tot Enrl'])
        elif spring16.loc[row, 'Component'] == 'LAB':
            b.append(spring16.loc[row, 'Tot Enrl'])
    mean_lec_sp16 = (sum(a) / len(a))    
    mean_lab_sp16 = (sum(b) / len(b))

    for row in fall16.index:
        if fall16.loc[row, 'Component'] == 'LEC':
            c.append(fall16.loc[row, 'Tot Enrl'])
        elif fall16.loc[row, 'Component'] == 'LAB':
            d.append(fall16.loc[row, 'Tot Enrl'])
    mean_lec_fall16 = (sum(c) / len(c))    
    mean_lab_fall16 = (sum(d) / len(d))

    for row in spring17.index:
        if spring17.loc[row, 'Component'] == 'LEC':
            e.append(spring17.loc[row, 'Tot Enrl'])
        elif spring17.loc[row, 'Component'] == 'LAB':
            f.append(spring17.loc[row, 'Tot Enrl'])
    mean_lec_sp17 = (sum(e) / len(e))    
    mean_lab_sp17 = (sum(f) / len(f))

    for row in fall17.index:
        if fall17.loc[row, 'Component'] == 'LEC':
            g.append(fall17.loc[row, 'Tot Enrl'])
        elif fall17.loc[row, 'Component'] == 'LAB':
            h.append(fall17.loc[row, 'Tot Enrl'])
    mean_lec_fall17 = (sum(g) / len(g))    
    mean_lab_fall17 = (sum(h) / len(h))

    for row in spring18.index:
        if spring18.loc[row, 'Component'] == 'LEC':
            i.append(spring18.loc[row, 'Tot Enrl'])
        elif spring18.loc[row, 'Component'] == 'LAB':
            j.append(spring18.loc[row, 'Tot Enrl'])
    mean_lec_sp18 = (sum(i) / len(i))    
    mean_lab_sp18 = (sum(j) / len(j))

    for row in fall18.index:
        if fall18.loc[row, 'Component'] == 'LEC':
            k.append(fall18.loc[row, 'Tot Enrl'])
        elif fall18.loc[row, 'Component'] == 'LAB':
            l.append(fall18.loc[row, 'Tot Enrl'])
    mean_lec_fall18 = (sum(k) / len(k))    
    mean_lab_fall18 = (sum(l) / len(l))

    
    data = [{'a:sp14':mean_lab_sp14,'b:fall14':mean_lab_fall14, 'c:sp15':mean_lab_sp15, 'd:fall15': mean_lab_fall15, 'e:sp16':mean_lab_sp16, 'f:fall16':mean_lab_fall16, 'g:sp17':mean_lab_sp17, 'h:fall17':mean_lab_fall17, 'i:sp18':mean_lab_sp18, 'j:fall18':mean_lab_fall18}]
    df = pd.DataFrame(data)
    return df
    print(df)

def make_fig(df):
    pltDF = create_df()
    plt.plot(pltDF.loc[0,:], linestyle = 'dashed')
    
    ax = plt.gca()
    ax.set_ylabel('Total enrollemnt ')
    ax.yaxis.label.set_fontsize(24)
    ax.set_title("Labs for semsesters", fontsize = 24)
    ax.set_xlabel('Semesters', fontsize = 24)
    pass

def main():
    df = create_df()
    print(df)
main()