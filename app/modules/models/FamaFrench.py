"""
Fama and French Three-Factor Model
https://www.investopedia.com/terms/f/famaandfrenchthreefactormodel.asp

r = rf + β1(rm - rf) + β2(SMB) + β3(HML) + ↋
r = Expected rate of return
rf = Risk-free rate
ß = Factor’s coefficient (sensitivity)
(rm – rf) = Market risk premium
SMB (Small Minus Big) = Historic excess returns of small-cap companies over large-cap companies
HML (High Minus Low) = Historic excess returns of value stocks (high book-to-price ratio) over growth stocks (low book-to-price ratio)
↋ = Risk
"""
import tushare as ts
import pandas as pd
from sklearn import linear_model

# 导入本地库
from ..config import tushareToken

# 初始化pro接口
pro = ts.pro_api(tushareToken)


class FamaFrench:

    def __init__(self, stock: str):
        self.records = None
        self.factors = None
        self.data = None
        self.weights = None
        self.stock = stock

    def __str__(self):
        if self.records is None:
            return "没有从TuShare拉取股票行情数据，请使用pull()获取数据！"
        if self.factors is None or self.data is None:
            return "没有建立多因子DataFrame，请使用build()构建！"
        if self.weights is None:
            return "没有进行多因子回归，请使用regression()回归！"
        return f"{self.stock} 的Fama French三因子相关性为：\r\n市场风险因子：{self.weights[0]}\r\n" \
               f"规模风险因子：{self.weights[1]}\r\n账面市值比风险因子：{self.weights[2]}\r\n回归公式：" \
               f"r = rf + {self.weights[0]}(rm - rf) + {self.weights[1]}(SMB) + {self.weights[2]}(HML) + ↋"

    # 获取股票的月K数据
    def pull(self):
        df = pro.monthly(**{
            "ts_code": self.stock,
            "trade_date": "",
            "start_date": "",
            "end_date": "",
            "limit": 12,
            "offset": ""
        }, fields=[
            "trade_date",
            "pct_chg"
        ])
        df['trade_date'] = df['trade_date'].str[0:6].astype(int)
        df = df.rename(columns={'pct_chg': 'r'})
        self.records = df

    # 插入同期因子收益
    def build(self):
        df = pd.read_csv('./app/modules/models/data/fivefactor_monthly.csv',
                         usecols=['trdmn', 'mkt_rf', 'smb', 'hml', 'rf'])
        df = df.rename(columns={'trdmn': 'trade_date'})
        self.factors = df
        df = pd.merge(self.records, self.factors, on='trade_date', how='left')
        df['r_rf'] = df['r'] - df['rf']
        self.data = df

    # 多因子回归
    def regression(self):

        X = self.data[['mkt_rf', 'smb', 'hml']]
        y = self.data['r_rf']

        regr = linear_model.LinearRegression()
        regr.fit(X, y)

        self.weights = regr.coef_
