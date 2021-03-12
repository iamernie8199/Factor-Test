# Factor Test

## Usage
1. 放置MC9/12報表到 StrategyReport-1
2. 放置因子資料到 factor/dirty
3. 執行preprocess.py進行因子資料前處理
4. 執行factor.py輸出因子報表
5. 執行report.py輸出績效報表
6. 執行factor_test2.py

## 檔案/資料夾說明

### StrategyReport

放置MC9/12報表

路徑: StrategyReport-1/類/商品/*.xls(xlsx)

### period_report

產出之各期間原始分析報表

路徑: period_report/類/策略名.xlsx

### output.xlsx

依設定標準篩除後績效報表, 每一因子一張工作表

數值為變化百分比(四捨五入到小數點後兩位)

績效 = $\frac{(篩後績效 - 原始績效) \times 100}{原始績效}$

MDD = $\frac{(原始MDD - 篩後MDD) \times 100}{原始績效}$

### setting.py

- freq(list): 回測窗格長度
    - 預設為[月, 季, 半年]
        - [30, 90, 180]
- freq_dict: 窗格長度對應中文
- criteria_dict(dict): 篩選標準
    - 編輯list更改
    - x%: 該點在該時間點前(含)的統計百分位數
- criteria_operator(dict): 定義篩選標準用運算子
- longterm: 略過之外部因子(週期: 周/月)
- econ_criteria_dict: 外部因子篩選標準
    - log_ret: 每日log return
    - return: 期間return(未減1)
    - abs_r: 期間漲跌幅絕對值
    - range: 期間最高-期間最低
    - ATR

### factor.py

統計因子並輸出報表

* log return: 日log return
* return: exp(區間log return.sum()) (未減1)
* range: 區間最高-區間最低
* ATR: ATR(區間長度)
* abs_r: 漲跌幅(絕對值)

流程:
1. 讀取因子數據
2. 計算收益率,對數收益率,TR...
   * 輕原油2020有負值, 以複數計算對數收益率
3. 繪製並輸出(jpg)時序統計圖
4. 補上缺少日期
   * 缺失值以上一筆資料填補, o=h=l=c
5. 依補值重新計算每日數據
6. 依區間長度計算結果
7. 輸出報表

#### tsplot()

輸出因子時間序列統計圖, 包括:
* 時序圖
* ACF
* PACF
* Q–Q plot
* probability plot
* 直方統計圖

### factor_test.py

流程:

1. 取得報表檔名
2. 建立篩後結果excel
3. 初步構建excel表格(篩選標準/期間)並上色
4. 檢查資料夾
5. 建立各期間報表用excel
6. 從MC報表切出日報酬
7. 初步處理日報酬(轉換type,drop未使用資料...)
8. 計算各期間日報酬因子
9. 輸出各期間報表
10. 依條件篩出符合期間並計算變化
11. 將獲益/MDD變化寫入篩後結果excel並上色
12. 將說明欄寫入篩後結果報表最下方

### factor_test2.py

直接讀取運算後績效及外部因子各期間報表

略過factor_test第6~9步

### report.py

產生績效各期間報表

略過factor_test第2~3、10~12步

### preprocess.py

parser, 以便後續搭配爬蟲更新及擴充, 以資料來源區分

統一格式及column名稱, 並刪除缺失值

### factor/dirty

因子raw data, 依來源分類

### factor/clean

因子處理後data

### factor/report

外部因子各期間報表

### factor/report/pic

外部因子時序統計圖

## 報表變數說明

|   變數   |   說明   |
| ------- |--------- |
| 獲利(¤)  | 當天獲利  |
| max_pnl | 權益高點  |
| DD     | DD(點數) |
| DD_pct | DD(%)   |
| mean    | 平均獲利   |
| volatility | 波動度(點數) |
| win_count | 獲利次數 |
| loss_count | 虧損次數 |
| pnl_w   | 毛利      |
| pnl_l   | 毛損      |
| avg_w | 平均獲利交易 |
| avg_l | 平均虧損交易 |
| avg_wl  | 賺賠比    |
| kelly   | 凱利值    |
| winrate | 勝率      |
| hazard  | 風報比    |
| pf_factor | 獲利因子 |
| sharpe | sharpe ratio |
| last | 前期間獲利>=0 |

## 因子

* BLS(略)
    * parser: bls()
    * 原始資料: [bls](https://data.bls.gov/cgi-bin/surveymost?bls)
        * 失業率: Labor Force Statistics from the Current Population Survey
        * 非農就業(差值): Employment, Hours, and Earnings from the Current Employment Statistics survey (National)
* EIA(略)
    * parser: eia()
    * 原始資料: [EIA - Stocks of Crude Oil by PAD District, and Stocks of Petroleum Products, U.S. Totals](https://www.eia.gov/petroleum/supply/weekly/)
  * 原油庫存周變動(差值): Weekly U.S. Ending Stocks excluding SPR of Crude Oil  (Thousand Barrels)
* yahoo
    * parser: yahoo()
    * 原始資料: [yahoo finance](https://finance.yahoo.com/)
        * VIX
        * 10年公債殖利率(^TNX)
        * 5年公債殖利率(^FVX)
        * 比特幣(BTC-USD)
* DQ2
    * parser: dq2()
    * 原始資料: DQ2
        * 天然氣近一(N1NG&)
        * 輕原油近一(N1CL&)
        * 汽油近一(N1RB&)
        * 日圓(JPY)
        * 美元指數(YDX)
        * 黃金近一(O1GC&)
        * 歐元(EUR)