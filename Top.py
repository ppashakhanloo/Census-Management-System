import xlrd
from xlutils.copy import copy
import os
import Diagrammer


workbook_male = xlrd.open_workbook('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS', formatting_info=True)
workbook_female = xlrd.open_workbook('Data/WPP2015_POP_F01_3_TOTAL_POPULATION_FEMALE.XLS', formatting_info=True)
worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
worksheet_female = workbook_female.sheet_by_name('ESTIMATES')
workbook_pop_growth = xlrd.open_workbook('Data/WPP2015_POP_F02_POPULATION_GROWTH_RATE.XLS')
workbook_pop_growth = workbook_pop_growth.sheet_by_name('ESTIMATES')

non_country_keywords = ['WORLD','Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
               'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
               'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
               'Micronesia','Polynesia']

def find_negative_growth_countries(sheet, year):
    keywords = ['WORLD','Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
               'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
               'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
               'Micronesia','Polynesia']
    data = {}
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                x = int(float(row[(year-1950)//5 + 5])*100)
                if x < 0 :
                    data[row[2]] = x
    items = [k for k, v in data.items()]
    print(items)
    return None


def find_sorted_countries_interval(sheet, first, last):
    keywords = ['WORLD','Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
               'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
               'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
               'Micronesia','Polynesia']
    data = {}
    if (last%5)==0:
            last = last-1
    first = (first-1950)//5 + 5
    last = (last -1950)//5 + 5
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                x = 0
                for j in range(first,last+1) :
                    x = x+int(float(row[j])*100)
                data[row[2]] = x
    items = [(v, k) for k, v in data.items()]
    items.sort()
    items.reverse()             # so largest is first
    items = [k for v, k in items]
    print(items)
    return None

def find_countries(sheet, year):
    keywords = ['WORLD','Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
               'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
               'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
               'Micronesia','Polynesia']
    data = {}
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                data[row[2]] = int(float(row[(year-1950)//5 + 5])*100)
    items = [(v, k) for k, v in data.items()]
    items.sort()
    items.reverse()             # so largest is first
    items = [k for v, k in items]
    print(items)
    return None


def get_data_by_country_year(worksheet_male, worksheet_female, country, year):

    row_country_m, col_country_m = find_row_col_index(country, worksheet_male)
    row_year_m, col_year_m = find_row_col_index(year, worksheet_male)

    val_male = worksheet_male.cell(row_country_m, col_year_m).value

    row_country_f, col_country_f = find_row_col_index(country, worksheet_female)
    row_year_f, col_year_f = find_row_col_index(year, worksheet_female)

    val_female = worksheet_female.cell(row_country_f, col_year_f).value

    return val_male, val_female


def get_data_by_country(worksheet_male, worksheet_female, country, start_year, end_year):

    male_population = []
    female_population = []

    # for each year in range
    for year in range(start_year, end_year + 1):
        male_val, female_val = get_data_by_country_year(worksheet_male, worksheet_female, country, str(year))
        male_population.append(male_val)
        female_population.append(female_val)

    return male_population, female_population


def get_data_by_year(data_sheet, year):

    row_year_m, col_year_m = find_row_col_index(year, data_sheet)

    data = []
    for i in range(28, data_sheet.nrows):
            row = data_sheet.row_values(i)
            if not(row[2] in non_country_keywords):
                data.append(data_sheet.cell(i, col_year_m).value)

    return data


def change_data_by_country_year(workbook_male, workbook_female, country, year, male_or_female, new_val):
    row_country_m, col_country_m = find_row_col_index(country, workbook_male.sheet_by_name('ESTIMATES'))
    row_year_m, col_year_m = find_row_col_index(year, workbook_male.sheet_by_name('ESTIMATES'))

    row_country_f, col_country_f = find_row_col_index(country, workbook_female.sheet_by_name('ESTIMATES'))
    row_year_f, col_year_f = find_row_col_index(year, workbook_female.sheet_by_name('ESTIMATES'))

    wb_male = copy(workbook_male)
    wb_female = copy(workbook_female)

    if male_or_female == 'male':
        wb_male.get_sheet(0).write(row_country_m, col_year_m, new_val)
        wb_male.save('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS')
    elif male_or_female == 'female':
        wb_female.get_sheet(0).write(row_country_f, col_year_f, new_val)
        wb_female.save('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_FEMALE.XLS')


def find_row_col_index(string_value, sheet):
    for i in range(sheet.nrows):
         row = sheet.row_values(i)
         for j in range(len(row)):
              if row[j] == string_value:
                    return i,j
    return None

# for getting a valid option from user, options is the list of valid choices
def get_input_option(options, explanation):

    while True:
        # getting option
        option = raw_input(explanation)
        # if option is among options return true
        if option in options:
            return option
        else:
            print("Invalid choice. You can only choose from ")
            print(options)



while True:
    print('please enter command number:')
    print('1. get population information for male and female.')
    print('2. change population information for a country at special year.')
    print('3. plot population information of of a country.')
    print('4. plot population information for future.')
    print('5. sort population information.')
    print('6. exit.')
    print('7. Plot population of countries in boxplot.')
    print('8. exit.')
    print('9. exit.')
    print('10. exit.')
    command = get_input_option(["1", "2", "3", "4", "5", "6"], 'enter command: ')

    if command == '1':
        country = raw_input('enter country: ')
        year = raw_input('enter year: ')
        male_res, female_res = get_data_by_country_year(worksheet_male, worksheet_female, country, year)
        print('Male: '+str(male_res)+'\n'+'Female: '+str(female_res)+'\n')

    if command == '2':
        country = raw_input('enter country: ')
        year = raw_input('enter year: ')
        male_or_female = raw_input('enter male or female: ')
        new_val = raw_input('enter new value: ')
        change_data_by_country_year(workbook_male, workbook_female, country, year, male_or_female, new_val)
        print('Done.')

    if command == '3':
        country = raw_input('enter country')
        sex = raw_input('enter sex, m for male, f for female or anything else for both')
        output_dir = raw_input('output dir?')
        male_population, female_population = get_data_by_country(worksheet_male, worksheet_female, country, 1950, 2015)
        draw_male = (sex != "f")
        draw_female = (sex != "m")

        if draw_male:
            Diagrammer.draw_diagram(range(1950, 2016), male_population, 'male_population', 'year', '1000 persons', output_dir + "male.pdf")
        if draw_female:
            Diagrammer.draw_diagram(range(1950, 2016), male_population, 'female_population', 'year', '1000 persons', output_dir + "female.pdf")

        print('Diagram was created successfully')

    if command == '4':

        country = raw_input('Country?')
        estimate_methods = ['MEDIUM VARIANT', 'HIGH VARIANT', 'LOW VARIANT',
                            'CONSTANT-FERTILITY', 'INSTANT-REPLACEMENT', 'ZERO-MIGRATION',
                            'CONSTANT-MORTALITY', 'NO CHANGE']
        print('choose a method among :')
        print(estimate_methods)
        method = get_input_option(estimate_methods, 'method?')
        estimate_methods.index(method)
        worksheet_male = workbook_male.sheet_by_name(method)
        worksheet_female = workbook_female.sheet_by_name(method)
        male_data, female_data = get_data_by_country(worksheet_male, worksheet_female, country, 2015, 2100)
        total_population = []
        for i in range(len(male_data)):
            total_population.append(male_data[i] + female_data[i])

        output_dir = raw_input('output directory?')
        Diagrammer.draw_diagram(range(2015, 2101), total_population, 'population', 'year', '1000 persons',
                                output_dir + 'population.pdf')
        print('Diagram was drawn successfully!')

    if command == '5':
        year = raw_input('please insert year number:')
        year = int(year)
        find_countries(workbook_pop_growth, year)

    if command == '6':
        year = raw_input('year?')
        output_dir = raw_input('output directory?')
        data_male = get_data_by_year(worksheet_male, year)
        data_female = get_data_by_year(worksheet_female, year)

        Diagrammer.draw_box_diagram(data_male, 'Countries Male Population', '1000 persons', output_dir + 'countries_male_population.pdf')
        Diagrammer.draw_box_diagram(data_female, 'Countries Female Population', '1000 persons', output_dir + 'countries_female_population.pdf')
        print('Diagrams were drawn successfully!')

    if command == '7':
        year = raw_input('please insert year number:')
        year = int(year)
        find_negative_growth_countries(workbook_pop_growth, year)
        break

    if command == '8':
        #request5()
        print('6')
        break

    if command == '9':
        first = raw_input('please insert first of interval:')
        first = int(first)
        last = raw_input('please insert last of interval:')
        last = int(last)
        find_sorted_countries_interval(workbook_pop_growth, first, last)
        break
