from copy import deepcopy

from ddm_model import DDMModel
from settings import r_0, implied_erp


FAIR_PRICE_DATA_WINDOW = 15


class DDMRetrospective:
    def __init__(self, company, index_quotes, scenario, retrospective_window):
        print(f"Building retrospective for {company['name']}")
        self.data = company
        self.results = []
        self.error = ""
        try:
            self.build_retrospective(company, index_quotes, scenario, retrospective_window)
        except Exception as e:
            raise
            # TODO check if there's a need in such thing
            # self.error = str(e)

    def build_retrospective(self, company, index_quotes, scenario, retrospective_window):
        for retrospective_index in range(retrospective_window):
            retrospective_data = {
                "name": company["name"]
            }

            opts = ["total_revenue", "net_income", "market_cap", "payout_ratio", "shares_outstanding"]

            for o in opts:
                retrospective_data[o] = company[o][retrospective_index:]
            retrospective_data["close"] = company["close"][retrospective_index*12:]
            retrospective_index_quotes = index_quotes[12*retrospective_index:]
            new_scenario = deepcopy(scenario)
            new_scenario["r_0"] = r_0[-retrospective_index - 1]
            new_scenario["implied_erp"] = implied_erp[-retrospective_index - 1]
            new_scenario["phase_two_growth"] = new_scenario["r_0"] * 2
            calculate_fair_price = False
            fwd_dividends = []
            share_price = None
            if retrospective_index > FAIR_PRICE_DATA_WINDOW:
                calculate_fair_price = True
                share_price = company["close"][(retrospective_index - FAIR_PRICE_DATA_WINDOW)*12]
                min_idx = retrospective_index - FAIR_PRICE_DATA_WINDOW
                max_idx = retrospective_index
                ni = company["net_income"][min_idx:max_idx]
                pr = company["payout_ratio"][min_idx:max_idx]
                so = company["shares_outstanding"][min_idx:max_idx]
                for i in range(len(min(min(pr, ni), so))):
                    fwd_dividends.append(ni[i]*pr[i]/so[i])
            self.results.append(DDMModel(company=retrospective_data,
                                         scenario=scenario,
                                         index=retrospective_index_quotes,
                                         calculate_fair_price=calculate_fair_price,
                                         fwd_dividends=fwd_dividends[::-1],
                                         share_price=share_price).results)


