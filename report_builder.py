import openpyxl
from typing import List, Tuple

from openpyxl.styles import PatternFill


class DDMReportBuilder:
    def __init__(self, results: List[Tuple], output_file: str):
        wb = openpyxl.Workbook()
        idx = 0
        wb.remove(wb.active)
        for (result, scenario) in results:
            self.create_result_sheet(wb, result, scenario, idx)
            idx += 1
        wb.save(output_file)

    def create_result_sheet(self, wb: openpyxl.Workbook, scenario_result, scenario, scenario_number):
        headers = ["Ticker", "Name", "Beta", "R_e", "Market Price", "DDM Price", "Error"]
        sheet = wb.create_sheet(f"Scenario {scenario_number}")
        sheet.append(headers)
        lower_border = 0.5
        high_border = 2
        good_results = 0

        idx = 0
        for i, v in scenario_result.items():
            idx += 1
            ddm_price = v.get("ddm_price")
            close_price = v["close"]
            sheet.append([i, v["name"], v.get("beta"), v.get("r_e"), close_price, ddm_price, v.get("error")])
            if ddm_price:
                ratio = ddm_price / close_price
                if 1 <= ratio < high_border or lower_border <= ratio <= 1:
                    good_results += 1
                    for cell in sheet.iter_cols(min_col=0, max_col=7, min_row=idx + 1, max_row=idx + 1):
                        cell[0].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

        sheet.append([])
        sheet.cell(1, len(headers) + 2).value = "Good results:"
        sheet.cell(1, len(headers) + 3).value = good_results
        row = 2
        for k, v in scenario.items():
            sheet.cell(row, len(headers) + 2).value = k
            sheet.cell(row, len(headers) + 3).value = v
            row += 1

        dims = {}
        for row in sheet.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, value in dims.items():
            sheet.column_dimensions[col].width = value
