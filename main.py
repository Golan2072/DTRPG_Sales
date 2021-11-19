# Stellagama Publishing Account Analysis Software

import csv
import openpyxl
from openpyxl.styles import numbers
from openpyxl.styles import Font


class Book_product:
    def __init__(self, product_name):
        self.name = product_name
        self.total_count = 0
        self.monthly_count = 0
        self.total_revenue = 0
        self.monthly_revenue = 0
        self.author = ""
        self.cost = 0
        self.release_month = 0
        self.release_year = 0


class Data_item:
    def __init__(self, data_name):
        self.name = data_name
        self.author = ""
        self.cost = 0
        self.release_month = 0
        self.release_year = 0


def month_calculation(start_month, start_year, end_month, end_year):
    year_difference = int(end_year) - int(start_year)
    if year_difference < 0:
        year_difference = 0
    return year_difference*12 + int(end_month-start_month)


def product_lister(dict_list):
    products = []
    for dict in dict_list:
        if dict["Name"] not in products:
            products.append(dict["Name"])
    products.sort()
    products.remove("")
    return products


def merge_products(main_product, secondary_product):
    new_product = Book_product(main_product)
    new_product.name = main_product.name
    new_product.total_count = main_product.total_count + secondary_product.total_count
    new_product.monthly_count = main_product.monthly_count + secondary_product.monthly_count
    new_product.total_revenue = main_product.total_revenue + secondary_product.total_revenue
    new_product.monthly_revenue = main_product.monthly_revenue + secondary_product.monthly_revenue
    return new_product


def product_cleaner(product_dictionary):
    output_product_dictionary = {}
    for product in product_dictionary:
        if product_dictionary[product].total_revenue != 0:
            output_product_dictionary[product] = product_dictionary[product]
        else:
            pass
    output_product_dictionary["TSAO: Wreck in the Ring"] = merge_products(product_dictionary["TSAO: Wreck in the Ring"], product_dictionary["Borderlands Adventure 1: Wreck in the Ring"])
    output_product_dictionary.pop("Borderlands Adventure 1: Wreck in the Ring")
    output_product_dictionary["Character Options for Stars Without Number"] = merge_products(product_dictionary["Character Options for Stars Without Number"], product_dictionary["Character Options - compatible with Stars Without Number"])
    output_product_dictionary.pop("Character Options - compatible with Stars Without Number")
    output_product_dictionary["Cheating Death 2nd Edition"] = merge_products(product_dictionary["Cheating Death 2nd Edition"], product_dictionary["Cheating Death"])
    output_product_dictionary.pop("Cheating Death")
    output_product_dictionary["From the Ashes 2nd Edition"] = merge_products(product_dictionary["From the Ashes 2nd Edition"], product_dictionary["From the Ashes"])
    output_product_dictionary.pop("From the Ashes")
    output_product_dictionary["TSAO: Stationery and Heraldry Pack"] = merge_products(product_dictionary["TSAO: Stationery and Heraldry Pack"], product_dictionary["TSAO: Stationary and Heraldry Pack"])
    output_product_dictionary.pop("TSAO: Stationary and Heraldry Pack")
    output_product_dictionary["TSAO: Liberty Ship"] = merge_products(product_dictionary["TSAO: Liberty Ship"], product_dictionary["Liberty Ship"])
    output_product_dictionary.pop("Liberty Ship")
    output_product_dictionary["TSAO: 50 Wonders of the Reticulan Empire"] = merge_products(product_dictionary["TSAO: 50 Wonders of the Reticulan Empire"], product_dictionary["50 Wonders of the Reticulan Empire"])
    output_product_dictionary.pop("50 Wonders of the Reticulan Empire")
    output_product_dictionary["The Sword of Cepheus"] = merge_products(product_dictionary["The Sword of Cepheus"], product_dictionary["Sword of Cepheus"])
    output_product_dictionary.pop("Sword of Cepheus")
    output_product_dictionary["TSAO: Wreck in the Ring"] = merge_products(product_dictionary["TSAO: Wreck in the Ring"], product_dictionary["TSAO: Borderlands Adventure 1: Wreck in the Ring"])
    output_product_dictionary.pop("TSAO: Borderlands Adventure 1: Wreck in the Ring")
    output_product_dictionary["TSAO: These Stars Are Ours!"] = merge_products(product_dictionary["TSAO: These Stars Are Ours!"], product_dictionary["These Stars Are Ours!"])
    output_product_dictionary.pop("These Stars Are Ours!")
    output_product_dictionary["Barbaro!"] = output_product_dictionary["Â¡BÃ¡rbaro!"]
    output_product_dictionary.pop("Â¡BÃ¡rbaro!")
    return output_product_dictionary


def excel_output(product_dictionary):
    filename = "output" + ".xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Monthly Output"
    sheet['A1'] = "Monthly Output"
    sheet['A1'].font = Font(size=16, bold=True)
    for cell in ["A3", "B3", "C3", "D3", "E3", "F3", "G3", "H3", "I3", "J3"]:
        sheet[cell].font = Font(bold=True)
    sheet['A3'] = "Product"
    sheet['B3'] = "Author"
    sheet['C3'] = "Production Cost"
    sheet['D3'] = "Total Paid Sales"
    sheet['E3'] = "Monthly Paid Sales"
    sheet['F3'] = "Total Revenue"
    sheet['G3'] = "Monthly Revenue"
    sheet['H3'] = "Release Year"
    sheet['I3'] = "Release Month"
    sheet['J3'] = "Months Since Release"
    for row in sheet.iter_rows(min_row=2, max_row=1000):
        for column in sheet.iter_cols(min_row=2, max_row=150, min_col=6, max_col=7):
            for cell in column:
                cell.number_format = '0.00'
        for column in sheet.iter_cols(min_row=2, max_row=150, min_col=3, max_col=3):
            for cell in column:
                cell.number_format = '0.00'
    sheet.column_dimensions['A'].width = 52
    sheet.column_dimensions['B'].width = 32
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20
    sheet.column_dimensions['F'].width = 20
    sheet.column_dimensions['G'].width = 20
    sheet.column_dimensions['H'].width = 20
    sheet.column_dimensions['I'].width = 20
    sheet.column_dimensions['J'].width = 20
    excel_row = 3
    for product in product_dictionary:
        excel_row += 1
        sheet.cell(row=excel_row, column=1).value = str(product_dictionary[product].name)
        sheet.cell(row=excel_row, column=2).value = str(product_dictionary[product].author)
        sheet.cell(row=excel_row, column=3).value = float(product_dictionary[product].cost)
        sheet.cell(row=excel_row, column=4).value = float(product_dictionary[product].total_count)
        sheet.cell(row=excel_row, column=5).value = float(product_dictionary[product].monthly_count)
        sheet.cell(row=excel_row, column=6).value = float(product_dictionary[product].total_revenue)
        sheet.cell(row=excel_row, column=7).value = float(product_dictionary[product].monthly_revenue)
        sheet.cell(row=excel_row, column=8).value = float(product_dictionary[product].release_year)
        sheet.cell(row=excel_row, column=9).value = float(product_dictionary[product].release_month)
        sheet.cell(row=excel_row, column=10).value = month_calculation(int(product_dictionary[product].release_month), int(product_dictionary[product].release_year), 11, 2021)
    workbook.save(filename)


if __name__ == '__main__':
    with open("dtrpg-report.csv", "r", errors="ignore") as raw_data:
        reader = csv.DictReader(raw_data)
        dict_list = []
        for line in reader:
            dict_list.append(line)
    products = product_lister(dict_list)
    product_dictionary = {}
    for product in products:
        product_dictionary[product] = Book_product(product)
        for product_dict in dict_list:
            if product_dict["Name"] == product:
                product_dictionary[product].total_revenue += float(product_dict["Earnings"])
                product_dictionary[product].total_count += int(product_dict["Quantity"])
                if product_dict["Date"][0:7] == "2021-10" or product_dict["Date"][3:10] == "10/2021":
                    product_dictionary[product].monthly_revenue += float(product_dict["Earnings"])
                    product_dictionary[product].monthly_count += int(product_dict["Quantity"])
                if product_dict["Earnings"] == "0":
                    pass
                else:
                    pass
            else:
                pass
    product_dictionary = product_cleaner(product_dictionary)
    data_list = []
    with open("database.csv", "r", errors="ignore") as base_data:
        reader = csv.DictReader(base_data)
        item_dictionary = {}
        for line in reader:
            if line["ï»¿Product"] != "Ã‚Â¡BÃƒÂ¡rbaro!":
                product_dictionary[line["ï»¿Product"]].cost = line["Cost"]
                product_dictionary[line["ï»¿Product"]].author = line["Author"]
                product_dictionary[line["ï»¿Product"]].release_year = line["Release_Year"]
                product_dictionary[line["ï»¿Product"]].release_month = line["Release_Month"]
            else:
                pass
    excel_output(product_dictionary)
