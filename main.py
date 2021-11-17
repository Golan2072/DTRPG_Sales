# Stellagama Publishing Account Analysis Software

import csv
import openpyxl
from openpyxl.styles import Font

with open("dtrpg-report.csv", "r", errors="ignore") as raw_data:
    reader = csv.DictReader(raw_data)
    dict_list = []
    for line in reader:
        dict_list.append(line)


class Book_product:
    def __init__(self, product_name):
        self.name = product_name
        self.total_count = 0
        self.monthly_count = 0
        self.total_revenue = 0
        self.monthly_revenue = 0


def product_lister(dict_list):
    products = []
    for dict in dict_list:
        if dict["Name"] not in products:
            products.append(dict["Name"])
    products.sort()
    products.remove("")
    return products


def excel_output(product_dictionary):
    filename = "output" + ".xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Monthly Output"
    sheet['A1'] = "Monthly Output"
    sheet['A1'].font = Font(size=16, bold=True)
    for cell in ["A3", "B3", "C3", "D3", "E3"]:
        sheet[cell].font = Font(bold=True)
    sheet['A3'] = "Product"
    sheet['B3'] = "Total Paid Sales"
    sheet['C3'] = "Monthly Paid Sales"
    sheet['D3'] = "Total Revenue"
    sheet['E3'] = "Monthly Revenue"
    excel_row = 3
    for product in product_dictionary:
        excel_row += 1
        sheet.cell(row=excel_row, column=1).value = product_dictionary[product].name
        sheet.cell(row=excel_row, column=2).value = product_dictionary[product].total_count
        sheet.cell(row=excel_row, column=3).value = product_dictionary[product].monthly_count
        sheet.cell(row=excel_row, column=4).value = product_dictionary[product].total_revenue
        sheet.cell(row=excel_row, column=5).value = product_dictionary[product].monthly_revenue
    workbook.save(filename)


if __name__ == '__main__':
    products = product_lister(dict_list)
    product_dictionary = {}
    for product in products:
        product_dictionary[product] = Book_product(product)
        for product_dict in dict_list:
            if product_dict["Name"] == product:
                if product_dict["Earnings"] != "0":
                    product_dictionary[product].total_revenue += float(product_dict["Earnings"])
                    product_dictionary[product].total_count += int(product_dict["Quantity"])
                if product_dict["Earnings"] == "0":
                    pass
                else:
                    pass
                if product_dict["Date"][3:10] == "10/2021":
                    if product_dict["Earnings"] != "0":
                        product_dictionary[product].monthly_revenue += float(product_dict["Earnings"])
                        product_dictionary[product].monthly_count += int(product_dict["Quantity"])
                    if product_dict["Earnings"] == "0":
                        pass
                    else:
                        pass
            else:
                pass
    total_revenue = 0
    monthly_total_revenue = 0
    for product in product_dictionary:
        if product_dictionary[product].total_revenue != 0 and product_dictionary[product].total_count != 0:
            print(product, "- Total: ", "{:.2f}".format(product_dictionary[product].total_revenue), "Count: ", product_dictionary[product].total_count, "Month: ", product_dictionary[product].monthly_count, "Month Revenue: ", "{:.2f}".format(product_dictionary[product].monthly_revenue))
            total_revenue += product_dictionary[product].total_revenue
            monthly_total_revenue += product_dictionary[product].monthly_revenue
        else:
            pass
    print("Total Stellagama Revenue since January 2016: ", "{:.2f}".format(total_revenue), "Total monthly revenue: ", "{:.2f}".format(monthly_total_revenue))
    excel_output(product_dictionary)