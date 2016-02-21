import xlrd

workbook_male = xlrd.open_workbook('Data/WPP2015_POP_F01_2_TOTAL_POPULATION_MALE.XLS')
workbook_female = xlrd.open_workbook('Data/WPP2015_POP_F01_3_TOTAL_POPULATION_FEMALE.XLS')
worksheet_male = workbook_male.sheet_by_name('ESTIMATES')
worksheet_female = workbook_female.sheet_by_name('ESTIMATES')


def get_data_by_country_year(worksheet_male, worksheet_female, country, year):
    row_country_m, col_country_m = find_row_col_index(country, worksheet_male)
    row_year_m, col_year_m = find_row_col_index(year, worksheet_male)

    val_male = worksheet_male.cell(row_country_m, col_year_m).value

    row_country_f, col_country_f = find_row_col_index(country, worksheet_female)
    row_year_f, col_year_f = find_row_col_index(year, worksheet_female)

    val_female = worksheet_female.cell(row_country_f, col_year_f).value

    return val_male, val_female


def find_row_col_index(string_value, sheet):
    for i in range(sheet.nrows):
         row = sheet.row_values(i)
         for j in range(len(row)):
              if row[j] == string_value:
                    return i,j
    return None


print(get_data_by_country_year(worksheet_male, worksheet_female, "Djibouti", "1950"))





while True:
    print('please enter command number:')
    print('1. get population information for male and female.')
    print('2. change population information for a country at special year.')
    print('3. plot population information of of a country.')
    print('4. plot population information for future.')
    print('5. sort population information.')
    print('6. exit.')
    command = input('enter command:')
    if command == '1':
        #request1()
        print('1')
    else:
        if command == '2':
            #request2()
            print('2')
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
                        #request5()
                        print('5')
                    else:
                        if command == '6':
                            #request5()
                            print('6')
                            break
                        else:
                            print('invalid instruction!')
