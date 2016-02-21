import xlrd
from xlutils3.copy import copy
import os

workbook_male = xlrd.open_workbook('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS')
workbook_female = xlrd.open_workbook('Data/WPP2015_POP_F01_3_TOTAL_POPULATION_FEMALE.XLS')
worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
worksheet_female = workbook_female.sheet_by_name('ESTIMATES')
workbook_pop_growth = xlrd.open_workbook('Data/WPP2015_POP_F02_POPULATION_GROWTH_RATE.XLS')
workbook_pop_growth = workbook_pop_growth.sheet_by_name('ESTIMATES')

def find_countries(sheet, year):
    keywords = ['Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
                'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
                'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
                'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
                'Micronesia','Polynesia']
    data = {}
    for i in range(sheet.nrows):
        if i > 11:
             row = sheet.row_values(i)
             for j in range(len(row)):
                if not(row[2] in keywords):
                    data[row[2]] = row[(year-1950)/5 + 5]*100

    sorted(data.items(), key=lambda x:x[1])
    print(data)
    return None




def get_data_by_country_year(worksheet_male, worksheet_female, country, year):
    row_country_m, col_country_m = find_row_col_index(country, worksheet_male)
    row_year_m, col_year_m = find_row_col_index(year, worksheet_male)

    val_male = worksheet_male.cell(row_country_m, col_year_m).value

    row_country_f, col_country_f = find_row_col_index(country, worksheet_female)
    row_year_f, col_year_f = find_row_col_index(year, worksheet_female)

    val_female = worksheet_female.cell(row_country_f, col_year_f).value

    return val_male, val_female


def change_data_by_country_year(workbook_male, workbook_female, country, year, male_or_female, new_val):

    row_country_m, col_country_m = find_row_col_index(country, workbook_male.sheet_by_name('ESTIMATES'))
    row_year_m, col_year_m = find_row_col_index(year, workbook_male.sheet_by_name('ESTIMATES'))

    row_country_f, col_country_f = find_row_col_index(country, workbook_female.sheet_by_name('ESTIMATES'))
    row_year_f, col_year_f = find_row_col_index(year, workbook_female.sheet_by_name('ESTIMATES'))

    wb_male = copy(workbook_male)
    wb_female = copy(workbook_female)

    if male_or_female == 'male':
        wb_male.get_sheet(0).write(row_country_m, col_year_m, new_val)
        # os.remove('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS')
        wb_male.save('male.XLS')
    elif male_or_female == 'female':
        wb_female.get_sheet(0).write(row_country_f, col_year_f, new_val)
        # os.remove('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_FEMALE.XLS')
        wb_female.save('female.XLS')

    return 0


def find_row_col_index(string_value, sheet):
    for i in range(sheet.nrows):
         row = sheet.row_values(i)
         for j in range(len(row)):
              if row[j] == string_value:
                    return i,j
    return None

while True:
    print('please enter command number:')
    print('1. get population information for male and female.')
    print('2. change population information for a country at special year.')
    print('3. plot population information of of a country.')
    print('4. plot population information for future.')
    print('5. sort population information.')
    print('6. exit.')
    command = input('enter command: ')
    if command == '1':
        country = input('enter country: ')
        year = input('enter year: ')
        male_res, female_res = get_data_by_country_year(worksheet_male, worksheet_female, country, year)
        print('Male: '+str(male_res)+'\n'+'Female: '+str(female_res)+'\n')
    else:
        if command == '2':
            country = input('enter country: ')
            year = input('enter year: ')
            male_or_female = input('enter male or female: ')
            new_val = input('enter new value: ')
            change_data_by_country_year(workbook_male, workbook_female, country, year, male_or_female, new_val)
            print('Done.')
        else:
            if command == '3':
               # request3()
                print('3')
            else:
                if command == '4':
                    #request4()
                    print('4')
                else:
                    if command == '5':
                        command = input('please insert year:')
                        find_countries(workbook_pop_growth, command)
                    else:
                        if command == '6':
                            #request5()
                            print('6')
                            break
                        else:
                            print('invalid instruction!')
