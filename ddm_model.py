import math
import numpy as np


class DDMModel:
    def __init__(self, company, index, scenario):
        self.company = company
        self.scenario = scenario
        self.results = dict()
        self.results["name"] = self.company["name"]
        self.count_beta(index)
        self.count_means()
        self.calculate_forward_prices()

    def count_beta(self, index):
        company = self.company
        shares = company["close"]
        index_quotes = index
        if len(shares) < len(index):
            index_quotes = index[:len(shares)]
        shares_log = []
        for i in range(0, len(shares) - 1):
            shares_log.append(math.log(shares[i] / shares[i + 1]))
        index_log = []
        for i in range(0, len(index_quotes) - 1):
            index_log.append(math.log(index_quotes[i] / index_quotes[i + 1]))
        beta = np.cov(shares_log, index_log)[0][1] / np.var(index_log)
        self.results["beta"] = beta
        print(f"beta = {beta}")
        r_e = self.scenario["r_0"] + self.scenario["implied_erp"] * beta
        self.results["r_e"] = r_e

    def count_means(self):
        company = self.company
        tr = company["total_revenue"]
        ni = company["net_income"]
        period = len(tr)
        market_cap = company["market_cap"]
        total_revenue_mean = np.mean(tr)
        payout_ratio = company["payout_ratio"]
        payout_ratio_mean = 0
        if payout_ratio and period:
            payout_ratio_mean = sum(payout_ratio[0:3])/3
        ros = 0
        for i in range(len(tr)):
            ni_i = ni[i] if i < len(ni) else 0
            ros += ni_i/tr[i]
        ros /= len(tr)

        p_s_ratio = 0
        p_s_ratio_period = period // 2
        for i in range(p_s_ratio_period):
            market_cap_i = market_cap[i] if i < len(market_cap) else 0
            if market_cap_i:
                p_s_ratio += math.log(market_cap_i/tr[i])
        p_s_mean = math.exp(p_s_ratio/p_s_ratio_period)

        self.results["ROS"] = ros
        self.results["total_revenue_mean"] = total_revenue_mean
        self.results["P/S mean"] = p_s_mean
        self.results["payout_ratio_mean"] = payout_ratio_mean

    def calculate_forward_prices(self):
        scenario = self.scenario
        forward_tr = []
        company = self.company
        results = self.results
        tr_mean = self.results["total_revenue_mean"]
        #tr_mean = company["total_revenue"][0]
        for i in range(scenario["total_periods"]):
            if i < scenario["phase_one_len"]:
                forward_tr.append(tr_mean*(1 + scenario["phase_one_growth"])**(i+1))
            elif i < scenario["phase_one_len"] + scenario["phase_two_len"]:
                forward_tr.append(forward_tr[-1]*(1 + scenario["phase_two_growth"]))
            else:
                forward_tr.append(forward_tr[-1]*(1 + scenario["r_0"]))
        price = 0
        k_dt = results["payout_ratio_mean"]
        for i in range(scenario["total_periods"]):
            sps = (forward_tr[i]/company["shares_outstanding"][0])
            eps = results["ROS"] * sps
            price += k_dt * eps/(1+results["r_e"])**(i+1)

        price += (forward_tr[-1]/company["shares_outstanding"][0])*results["P/S mean"]/(1+results["r_e"])**scenario["total_periods"]
        print(f"Model price = {price}, market price = {company['close'][0]}")
        self.results["close"] = company["close"][0]
        self.results["ddm_price"] = price
