import xlrd
from xlutils.copy import copy
import os
import Diagrammer
import xlwt


workbook_male = xlrd.open_workbook('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS', formatting_info=True)
workbook_female = xlrd.open_workbook('Data/WPP2015_POP_F01_3_TOTAL_POPULATION_FEMALE.XLS', formatting_info=True)
worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
worksheet_female = workbook_female.sheet_by_name('ESTIMATES')
workbook_pop_growth_main = xlrd.open_workbook('Data/WPP2015_POP_F02_POPULATION_GROWTH_RATE.XLS')
workbook_pop_growth = workbook_pop_growth_main.sheet_by_name('ESTIMATES')

non_country_keywords = ['WORLD','Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
               'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
               'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
               'Micronesia','Polynesia']


def find_negative_growth_countries():
    keywords = ['WORLD','Sub-Saharan Africa','AFRICA','Eastern Africa','Middle Africa', 'Northern Africa','Southern Africa','Western Africa',
               'ASIA','Eastern Asia','South-Central Asia','Central Asia','Southern Asia','South-Eastern Asia','Western Asia',
               'EUROPE','Eastern Europe','Northern Europe','Southern Europe','Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean','Central America','South America','NORTHERN AMERICA','OCEANIA','Australia/New Zealand','Melanesia',
               'Micronesia','Polynesia']
    estimate_methods = ['MEDIUM VARIANT', 'HIGH VARIANT', 'LOW VARIANT',
                            'CONSTANT-FERTILITY', 'INSTANT-REPLACEMENT', 'ZERO-MIGRATION',
                            'CONSTANT-MORTALITY', 'NO CHANGE']
    data = {}
    while (True):
        estimate_type = raw_input('please insert estimation type:')
        if not(estimate_type in estimate_methods):
            print('invalid estimation name!')
        else:
            break
    sheet = xlrd.open_workbook('Data/WPP2015_POP_F02_POPULATION_GROWTH_RATE.XLS')
    sheet = sheet.sheet_by_name(estimate_type)
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                for j in range(4,21):
                    x = int(float(row[j])*100)
                    if x < 0 :
                        data[row[2]] = x
                        break
    item = [(k, v/100.0) for k, v in data.items()]
    print(item)
    return None

def find_max(sheet, first, last):
    keywords = ['WORLD', 'Sub-Saharan Africa', 'AFRICA', 'Eastern Africa', 'Middle Africa', 'Northern Africa', 'Southern Africa', 'Western Africa',
               'ASIA', 'Eastern Asia', 'South-Central Asia', 'Central Asia', 'Southern Asia', 'South-Eastern Asia', 'Western Asia',
               'EUROPE', 'Eastern Europe','Northern Europe', 'Southern Europe', 'Western Europe', 'LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean', 'Central America', 'South America', 'NORTHERN AMERICA', 'OCEANIA', 'Australia/New Zealand', 'Melanesia',
               'Micronesia', 'Polynesia']
    data = {}
    if (last % 5) == 0:
            last -= 1

    first = (first-1950)//5 + 5
    last = (last -1950)//5 + 5
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                x = int(float(row[first])*100)
                for j in range(first+1,last+1) :
                    if x < int(float(row[j])*100):
                        x = int(float(row[j])*100)
                data[row[2]] = x
    items = [(v, k) for k, v in data.items()]
    items.sort()
    items.reverse()             # so largest is first
    items = [(k, v/100.0) for v, k in items]
    print(items)
    return None

def find_sorted_countries_interval(sheet, first, last):
    keywords = ['WORLD', 'Sub-Saharan Africa', 'AFRICA', 'Eastern Africa', 'Middle Africa', 'Northern Africa', 'Southern Africa', 'Western Africa',
               'ASIA', 'Eastern Asia', 'South-Central Asia', 'Central Asia', 'Southern Asia', 'South-Eastern Asia', 'Western Asia',
               'EUROPE', 'Eastern Europe','Northern Europe', 'Southern Europe', 'Western Europe', 'LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean', 'Central America', 'South America', 'NORTHERN AMERICA', 'OCEANIA', 'Australia/New Zealand', 'Melanesia',
               'Micronesia', 'Polynesia']
    data = {}
    if (last % 5) == 0:
            last -= 1

    first = (first-1950)//5 + 5
    last = (last -1950)//5 + 5
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                x = 0
                for j in range(first,last+1) :
                    x += int(float(row[j])*100)
                x = x//(last-first+1)
                data[row[2]] = x
    items = [(v, k) for k, v in data.items()]
    items.sort()
    items.reverse()             # so largest is first
    items = [(k, v/100.0) for v, k in items]
    print(items)
    return None



def find_countries(sheet, year):
    keywords = ['WORLD', 'Sub-Saharan Africa', 'AFRICA', 'Eastern Africa', 'Middle Africa', 'Northern Africa', 'Southern Africa', 'Western Africa',
               'ASIA', 'Eastern Asia', 'South-Central Asia', 'Central Asia', 'Southern Asia', 'South-Eastern Asia', 'Western Asia',
               'EUROPE', 'Eastern Europe', 'Northern Europe', 'Southern Europe', 'Western Europe','LATIN AMERICA AND THE CARIBBEAN',
               'Caribbean', 'Central America', 'South America', 'NORTHERN AMERICA', 'OCEANIA', 'Australia/New Zealand', 'Melanesia',
               'Micronesia', 'Polynesia']
    data = {}
    for i in range(sheet.nrows):
        if i > 27:
            row = sheet.row_values(i)
            if not(row[2] in keywords):
                data[row[2]] = int(float(row[(year-1950)//5 + 5])*100)
    items = [(v, k) for k, v in data.items()]
    items.sort()
    items.reverse()# so largest is first
    items = [(k, v/100.0) for v, k in items]
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


def get_growth_data_by_country(worksheet, country, years_ranges):

    data = []

    # for each year in range
    for year_range in years_ranges:
        row_country_m, col_country_m = find_row_col_index(country, worksheet)
        row_year_m, col_year_m = find_row_col_index(year_range, worksheet)
        data.append(worksheet.cell(row_country_m, col_year_m).value)

    return data


def get_data_by_year(data_sheet, year):

    row_year_m, col_year_m = find_row_col_index(str(year), data_sheet)

    data = []
    for i in range(28, data_sheet.nrows):
            row = data_sheet.row_values(i)
            if not(row[2] in non_country_keywords):
                data.append(data_sheet.cell(i, col_year_m).value)

    return data


def change_data_by_country_year(workbook_male, workbook_female, country, year, male_or_female, new_val):

    style = xlwt.XFStyle()
    style.num_format_str = '0.00'

    row_country_m, col_country_m = find_row_col_index(country, workbook_male.sheet_by_name('ESTIMATES'))
    row_year_m, col_year_m = find_row_col_index(year, workbook_male.sheet_by_name('ESTIMATES'))

    row_country_f, col_country_f = find_row_col_index(country, workbook_female.sheet_by_name('ESTIMATES'))
    row_year_f, col_year_f = find_row_col_index(year, workbook_female.sheet_by_name('ESTIMATES'))

    wb_male = copy(workbook_male)
    wb_female = copy(workbook_female)

    if male_or_female == 'male':
        wb_male.get_sheet(0).write(row_country_m, col_year_m, new_val, style)
        wb_male.save('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS')
    elif male_or_female == 'female':
        wb_female.get_sheet(0).write(row_country_f, col_year_f, new_val, style)
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


def make_protected(workbook_male, workbook_female, worksheet_male, worksheet_female, country):

    row_country_m, col_country_m = find_row_col_index(country, worksheet_male)
    row_country_f, col_country_f = find_row_col_index(country, worksheet_female)

    wb_male = copy(workbook_male)
    wb_female = copy(workbook_female)

    for i in range(17, 258):
        for j in range(3, 50):
            if i != row_country_m:
                wb_male.get_sheet(0).write(i, j, worksheet_male.cell(i, j).value, xlwt.easyxf('protection: cell_locked false;'))
            else:
                wb_male.get_sheet(0).write(i, j, worksheet_male.cell(i, j).value, xlwt.easyxf('protection: cell_locked true;'))
    wb_male.get_sheet(0).set_protect(True)
    wb_male.save('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS')

    for i in range(17, 258):
        for j in range(3, 50):
            if i != row_country_f:
                wb_female.get_sheet(0).write(i, j, worksheet_female.cell(i, j).value, xlwt.easyxf('protection: cell_locked false;'))
            else:
                wb_female.get_sheet(0).write(i, j, worksheet_female.cell(i, j).value, xlwt.easyxf('protection: cell_locked true;'))
    wb_female.get_sheet(0).set_protect(True)
    wb_female.save('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_FEMALE.XLS')



while True:
    print('Please enter command number:')
    print('1. View population information for male and female.')
    print('2. Change population information for a country at special year.')
    print('3. Plot population information of a country.')
    print('4. Plot population information for future.')
    print('5. Sort population information.')
    print('6. Plot population of countries in box plot.')
    print('7. View countries with negative growth rate.')
    print('8. Different growth estimates diagram for a country.')
    print('9. View sorted list of countries with regarding to average growth rate.')
    print('10. View sorted list of countries with regarding to maximum average growth rate.')
    print('11. Choose a country to protect.')
    print('12. Exit.')

    command = get_input_option(["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"], 'enter command: ')

    if command == '1':

        worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
        worksheet_female = workbook_female.sheet_by_name('ESTIMATES')

        country = raw_input('Enter country: ')
        year = raw_input('Enter year: ')
        male_res, female_res = get_data_by_country_year(worksheet_male, worksheet_female, country, year)
        print('Male: '+str(male_res)+'\n'+'Female: '+str(female_res)+'\n')

    if command == '2':

        workbook_male = xlrd.open_workbook('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS', formatting_info=True)
        workbook_female = xlrd.open_workbook('Data/WPP2015_POP_F01_3_TOTAL_POPULATION_FEMALE.XLS', formatting_info=True)

        country = raw_input('Enter country: ')
        year = raw_input('Enter year: ')
        male_or_female = raw_input('Enter male or female: ')
        new_val = raw_input('Enter new value as a floating point number: ')
        change_data_by_country_year(workbook_male, workbook_female, country, year, male_or_female, float(new_val))
        print('Done.')

        # re opening
        workbook_male = xlrd.open_workbook('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS', formatting_info=True)
        workbook_female = xlrd.open_workbook('Data/WPP2015_POP_F01_3_TOTAL_POPULATION_FEMALE.XLS', formatting_info=True)
        worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
        worksheet_female = workbook_female.sheet_by_name('ESTIMATES')

    if command == '3':
        worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
        worksheet_female = workbook_female.sheet_by_name('ESTIMATES')

        country = raw_input('Enter country: ')
        sex = raw_input('Enter male or female or anything else for both: ')
        output_dir = raw_input('Output directory: ')
        male_population, female_population = get_data_by_country(worksheet_male, worksheet_female, country, 1950, 2015)
        draw_male = (sex != "female")
        draw_female = (sex != "male")

        if draw_male:
            Diagrammer.draw_diagram(range(1950, 2016), male_population, country+' male population', 'year',
                                    '1000 persons', output_dir + "male_population.pdf")
        if draw_female:
            Diagrammer.draw_diagram(range(1950, 2016), female_population, country + ' female population',
                                    'year', '1000 persons', output_dir + "female_population.pdf")

        print('Diagram was successfully created.')

    if command == '4':

        country = raw_input('Enter country: ')
        estimate_methods = ['MEDIUM VARIANT', 'HIGH VARIANT', 'LOW VARIANT',
                            'CONSTANT-FERTILITY', 'INSTANT-REPLACEMENT', 'ZERO-MIGRATION',
                            'CONSTANT-MORTALITY', 'NO CHANGE']
        print('Choose a method:')
        print(estimate_methods)
        method = get_input_option(estimate_methods, 'method?')
        estimate_methods.index(method)
        m_worksheet = workbook_male.sheet_by_name(method)
        f_worksheet = workbook_female.sheet_by_name(method)
        male_data, female_data = get_data_by_country(m_worksheet, f_worksheet, country, 2015, 2100)
        total_population = []
        for i in range(len(male_data)):
            total_population.append(male_data[i] + female_data[i])

        output_dir = raw_input('Output directory: ')
        Diagrammer.draw_diagram(range(2015, 2101), total_population, country + ' estimated total population based on ' + method
                                , 'year', '1000 persons',
                                output_dir + 'estimated_total_population.pdf')
        print('Diagram was drawn successfully!')

    if command == '5':
        year = raw_input('Enter year: ')
        year = int(year)
        find_countries(workbook_pop_growth, year)

    if command == '6':

        worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
        worksheet_female = workbook_female.sheet_by_name('ESTIMATES')

        year = raw_input('Enter year: ')
        output_dir = raw_input('Output directory: ')

        data_male = get_data_by_year(worksheet_male, year)
        data_female = get_data_by_year(worksheet_female, year)

        Diagrammer.draw_box_diagram(data_male, 'Countries Male Population in ' + str(year), '1000 persons', output_dir + 'countries_male_population.pdf')
        Diagrammer.draw_box_diagram(data_female, 'Countries Female Population in ' + str(year), '1000 persons', output_dir + 'countries_female_population.pdf')
        print('Diagrams were successfully drawn.')

    if command == '7':
        find_negative_growth_countries()

    if command == '8':
        country = raw_input('Enter country: ')
        estimate_methods = ['MEDIUM VARIANT', 'HIGH VARIANT', 'LOW VARIANT',
                            'CONSTANT-FERTILITY', 'INSTANT-REPLACEMENT', 'ZERO-MIGRATION',
                            'CONSTANT-MORTALITY', 'NO CHANGE']
        colors = ['red', 'green', 'blue', 'cyan', 'magenta', 'yellow', 'orange', 'black']

        year_ranges = []
        year_mids = []
        for i in range(2015, 2100, 5):
            year_ranges.append(str(i) + '-' + str(i + 5))
            year_mids.append(i + 2.5)

        output_dir = raw_input('Output directory: ')

        for i in range(len(estimate_methods)):
            data = get_growth_data_by_country(workbook_pop_growth_main.sheet_by_index(i + 1), country, year_ranges)
            Diagrammer.draw_diagram(year_mids, data, country +' population growth estimates', 'year', 'percentage',
                                    output_dir + 'population_growth.pdf', False, colors[i], True, estimate_methods[i])

        print('Diagram was successfully drawn.')

    if command == '9':
        first = raw_input('Enter first of interval: ')
        first = int(first)
        last = raw_input('Enter last of interval: ')
        last = int(last)
        find_sorted_countries_interval(workbook_pop_growth, first, last)

    if command == '10':
        first = raw_input('Enter first of interval: ')
        first = int(first)
        last = raw_input('Enter last of interval: ')
        last = int(last)
        find_max(workbook_pop_growth, first, last)

    if command == '11':
        worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
        worksheet_female = workbook_female.sheet_by_name('ESTIMATES')

        country = raw_input('Enter country: ')
        make_protected(workbook_male, workbook_female, worksheet_male, worksheet_female, country)
        print('Done.')

    if command == '12':
        break
