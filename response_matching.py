#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created in Feb 2018

"""

##Please note that this is an example to match knife responses, and it's similar when we deal with brick responses

##Part I: Gather our ratings into a dictionary called dict

import xlrd
import numpy as np

# import the excel file data to python
data = xlrd.open_workbook("Brick cleaned.xlsx")
# choose the workbook we need
sheet = data.sheet_by_index(0)

# create an empty list to store keys as the responses and values as the average of the scores
dict = {}

for row in range(1, 848):            #rows of different  brick/knife responses
    dict[sheet.cell_value(row, 0)] = sheet.cell_value(row, 1)


##Part II: Match the responses from Reddit_Microdosing_Survey_Num.csv to the ratings in dict
    
# import the excel file data to python
#survey = open("Reddit_Microdosing_Survey_Num.csv")
survey = xlrd.open_workbook("brick.xlsx")
# choose the workbook we need
survey_sheet = survey.sheet_by_index(0)

# initialize an array with zeros to store the ratings of each response of each person
ratings = np.zeros((1379, 30))              #1387 people, each person 30 responses at most


n = 0       #number of non-empty responses
n1 = 0      #number of responses that are contained in the key
n2 = 0      #number of responses that contain key

for row in range(3, 1380):
    for column in range(30):
        if survey_sheet.cell_value(row, column) != '':
            n += 1
        if survey_sheet.cell_value(row, column) in dict:
            ratings[row-3][column] = dict[survey_sheet.cell_value(row, column)]
            n1 += 1
        else:
            for key in dict:
                if key in survey_sheet.cell_value(row, column):
                    ratings[row-3][column] = dict[key]
                    n2 += 1
                else:
                    ratings[row-3][column] = 0

import xlsxwriter

workbook = xlsxwriter.Workbook('arrays.xlsx')
worksheet = workbook.add_worksheet()
row = 0

for col, data in enumerate(ratings):
    worksheet.write_column(row, col, data)
    
workbook.close()

##Finally, we can get all the ratings of 1387 people, each person 30 responses at most for knife          
#np.savetxt('matching_output', ratings)       
            
def indict(response, dict):
    """
    Return True is the response is contained or contains any key in the dict, False otherwise.
    """            
    result = False
    if response in dict:
        result = True
    for key in dict:
        if key in response:
            result = True
    return result

    

unmatched = []

for row in range(3, 1380):
    for column in range(30):
        if survey_sheet.cell_value(row, column) != '' and indict(survey_sheet.cell_value(row, column), dict)==False and survey_sheet.cell_value(row, column) not in unmatched:
            unmatched.append(survey_sheet.cell_value(row, column))
                    
            
            


with open("file.txt", "w") as output:
    output.write(str(unmatched))
    
    
    
    
    