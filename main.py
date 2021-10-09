from data_reader import get_spx_index, read_data_from_file_to_object
from ddm_model import DDMModel
from ddm_retrospective import DDMRetrospective
from report_builder import DDMReportBuilder
from retrospective_report_builder import RetrospectiveReportBuilder
from settings import scenarios

companys = {}

spx_quotes = get_spx_index()
read_data_from_file_to_object('snp500 - price close 480 series.xlsx', 'close',  companys, series=480)
read_data_from_file_to_object('snp500 market cap 40y.xlsx', 'market_cap',  companys)
read_data_from_file_to_object('snp500 net income 40y.xlsx', 'net_income',  companys)
read_data_from_file_to_object('snp500 payout ratio 40y.xlsx', 'payout_ratio',  companys)
read_data_from_file_to_object('snp500 shares_outstanding 40y.xlsx', 'shares_outstanding',  companys)
read_data_from_file_to_object('snp500 total revenue 40y.xlsx', 'total_revenue', companys)


def run_all():
    scenarios_results = []
    for scenario in scenarios:
        scenario_result = dict()
        for k, v in companys.items():
            try:
                print(f"Company {k}")
                scenario_result[k] = DDMModel(v, spx_quotes, scenario).results
            except Exception as e:
                v["error"] = str(e)
                print(f"Error : {str(e)}")
        scenarios_results.append((scenario_result, scenario))
    return scenarios_results

results = run_all()
DDMReportBuilder(results, "output.xlsx")

retro_results = []
for scenario in scenarios:
    for k, v in companys.items():
        try:
                if "AON" not in k:
                    pass
  #          if k in scenario["advanced"]:
                retro_result = DDMRetrospective(company=v, scenario=scenario, index_quotes=spx_quotes, retrospective_window=30)
                retro_result.data["ticker"] = k
                retro_results.append(retro_result)
        except:
            import traceback
            traceback.print_exc()


RetrospectiveReportBuilder(retro_results, "retrospective_output.xlsx")



#print(companys)