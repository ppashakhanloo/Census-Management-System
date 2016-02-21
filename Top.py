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