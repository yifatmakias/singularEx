import sys
import xlrd


def print_date_dict(dict):
    counter = 0
    for old_date in dict.keys():
        if counter == len(dict) - 1:
            break
        y, m, d, h, ms, s = xlrd.xldate_as_tuple(old_date, wb.datemode)
        print("{0}-{1}-{2}".format(y, m, d))
        print(dict[old_date])
        counter = counter + 1


def print_dict(dict):
    counter = 0
    for key in dict.keys():
        if counter == len(dict) - 1:
            break
        print(key)
        print(dict[key])
        counter = counter + 1


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Please enter only one argument: A report file path.")
    else:
        file_path = sys.argv[1]
        try:
            wb = xlrd.open_workbook(file_path)
            sheet = wb.sheet_by_index(0)
            last_row = sheet.nrows

            # Total installs
            installs_cell = sheet.cell(last_row-1, 5)
            installs = installs_cell.value
            print("Total Installs: " + str(installs))

            # Total cost
            cost_cell = sheet.cell(last_row-1, 4)
            cost = cost_cell.value
            print("Total Cost: " + str(cost))

            # Total cost an installs by App, platform, date
            app_and_platform_cost_dictionary = {}
            app_and_platform_installs_dictionary = {}
            date_cost_dictionary = {}
            date_installs_dictionary = {}

            for i in range(4, last_row):
                row = sheet.row_values(i)
                date = row[0]
                app = row[1]
                platform = row[2].split(' ')[0]
                row_cost = row[4]
                row_installs = row[5]
                app_and_platform = app + ',' + platform

                if app_and_platform in app_and_platform_cost_dictionary:
                    app_and_platform_cost_dictionary[app_and_platform] = app_and_platform_cost_dictionary[app_and_platform] + row_cost
                else:
                    app_and_platform_cost_dictionary[app_and_platform] = row_cost

                if app_and_platform in app_and_platform_installs_dictionary:
                    app_and_platform_installs_dictionary[app_and_platform] = app_and_platform_installs_dictionary[app_and_platform] + row_installs
                else:
                    app_and_platform_installs_dictionary[app_and_platform] = row_installs

                if date in date_cost_dictionary:
                    date_cost_dictionary[date] = date_cost_dictionary[date] + row_cost
                else:
                    date_cost_dictionary[date] = row_cost

                if date in date_installs_dictionary:
                    date_installs_dictionary[date] = date_installs_dictionary[date] + row_installs
                else:
                    date_installs_dictionary[date] = row_installs

            print('\n')
            print("Total installs by app and platform:")
            print_dict(app_and_platform_installs_dictionary)
            print('\n')

            print("Total cost by app and platform:")
            print_dict(app_and_platform_cost_dictionary)
            print('\n')

            print("Total installs by date:")
            print_date_dict(date_installs_dictionary)
            print('\n')

            print("Total cost by date:")
            print_date_dict(date_cost_dictionary)

        except IOError:
            print("Please enter a valid path")
