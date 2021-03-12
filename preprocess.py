import os
from datetime import datetime
from glob import glob

import pandas as pd


def bls():
    file = glob('factor/dirty/bls/*.xlsx')
    for f in file:
        tmp = pd.read_excel(f)
        file_name = tmp.columns[0]
        # 切出數據部分
        tmp = tmp[tmp[tmp['Unnamed: 1'].isin(['Jan'])].index.values[0]:].reset_index(drop=True)
        tmp = tmp.set_axis(list(tmp.iloc[0]), axis=1, inplace=False)[1:].reset_index(drop=True)
        # 組成日期
        start = datetime.strptime(str(tmp.iloc[0][0]) + '-' + tmp.columns[1], "%Y-%b")
        end = str(tmp.dropna(axis='columns').iloc[-1][0]) + '-' + tmp.dropna(axis='columns').columns[-1]
        # 將表格轉為直式
        idx = pd.date_range(start=start, end=end, freq='BMS')
        tmp_list = []
        for i in range(len(tmp)):
            tmp_list += list(tmp.iloc[i][1:].dropna().values)
        df = pd.DataFrame(tmp_list, columns=[file_name], index=idx)
        # 非農數據用差值表示
        if df.iloc[0][0] > 1000:
            # df['diff'] = df[file_name].diff(1)
            df[file_name] = df[file_name].diff(1)
            df = df.dropna()
        df.to_csv(f'factor/clean/{file_name}.csv', index=True, index_label='Date')


def eia():
    # file = glob('factor/dirty/eia/*.xls')
    tmp = pd.read_excel('factor/dirty/eia/psw04.xls', sheet_name='Data 1')
    # 切出數據
    tmp = tmp[tmp[tmp['Back to Contents'].isin(['Date'])].index.values[0]:].reset_index(drop=True)
    tmp = tmp.set_axis(list(tmp.iloc[0]), axis=1, inplace=False)[1:].reset_index(drop=True)
    tmp = tmp.set_index('Date')
    df = pd.DataFrame(tmp['Weekly U.S. Ending Stocks excluding SPR of Crude Oil  (Thousand Barrels)'])
    # df['Crude Oil diff'] = df['Weekly U.S. Ending Stocks excluding SPR of Crude Oil  (Thousand Barrels)'].diff(1)
    df = df.diff(1).dropna()
    df.to_csv('factor/clean/U.S. Crude Oil Inventories.csv', index=True)


def yahoo():
    """
    可用API取代
    """
    file = glob('factor/dirty/yahoo/*.csv')
    for f in file:
        tmp = pd.read_csv(f)
        tmp = tmp.dropna().drop(columns=['Adj Close']).reset_index(drop=True)
        filename = f.split('\\')[-1]
        tmp.to_csv(f"factor/clean/{filename}", index=False)


def dq2():
    file = glob('factor/dirty/DQ2/*.csv')
    for f in file:
        # UTF-8會亂碼
        tmp = pd.read_csv(f, skiprows=[0], encoding='big5')
        tmp = tmp.dropna(axis='columns')
        # 調整日期格式
        tmp['日期'] = tmp['日期'].apply(lambda x: str(x // 10000) + '/' + str(x % 10000 // 100) + '/' + str(x % 100))
        # 與Yahoo統一欄位名稱
        tmp = tmp.rename(columns={"日期": "Date",
                                  "開盤價": "Open",
                                  "最高價": "High",
                                  "最低價": "Low",
                                  "收盤價": "Close",
                                  "成交量": "Volume",
                                  "未平倉量": "OI"})
        filename = f.split('\\')[-1].split('(')[-2].replace(')', '.csv')
        # 捨棄全0欄位
        tmp = tmp.loc[:, (tmp != 0).any(axis=0)]
        tmp.to_csv(f"factor/clean/{filename}", index=False)


if __name__ == "__main__":
    if not os.path.exists('factor/clean/'):
        os.makedirs('factor/clean/')
    bls()
    eia()
    yahoo()
    dq2()
