import openpyxl
import csv


def get_spx_index():
    spx_index_file = "spx.csv"
    data = []
    with open(spx_index_file, "r") as f:
        reader = csv.reader(f, delimiter=";")
        for row in reader:
            if row:
                data.append(float(row[0].replace(",", ".")))
    return data


def read_data_from_file_to_object(filename, prop_name,  write_object, series = 40):
    wb = openpyxl.open(filename)
    sheet = wb.active
    for r in sheet.iter_rows(4):
        ticket = r[0].value
        company_name = r[1].value
        if ticket in write_object:
            company = write_object[ticket]
        else:
            company = {"name": company_name}
            write_object[ticket] = company
        data = []
        for c in range(series):
            val = r[2 + c].value or 0
            data.append(val)
        while len(data) > 0 and not data[-1]:
            data.pop()
        company[prop_name] = data