from utils import *
import time


start = time.time()

# 初期実行 1.21s
init_process = InitProcess()
logger = init_process.set_log()
driver = init_process.set_selenium()
init_process.execute()
# finvizから銘柄の配列を取得 87.18s 1.45m
pick_finviz = PickFinviz(logger)
pick_finviz.execute()
# ヒストリカルデータをダウンロードする 3800.04s 63.33m
hist_data = HistData(logger, driver)
hist_data.execute()
# NASDAQのヒストリカルデータの列を増やす
process_nasdaq = ProcessNASDAQ(logger)
process_nasdaq.execute()
# 個別株のヒストリカルデータの列を増やす 610.45s 10.17m
process_histdata = ProcessHistData(logger)
process_histdata.execute()
# RSを計算する 91.94s 1.53m
calculate_rs = CalculateRS(logger)
calculate_rs.execute()
# BuyingStock.csvを出力する 0.3s
buying_stock = BuyingStock(logger)
buying_stock.execute()
# CとAを取得する 1881.23s 31.36m
current_annual = CurrentAnnual(logger, driver)
current_annual.execute()
# 機関投資家の増加数を取得  2345.98s 39.09m, 581.66s 240stocks
institutional = Institutional(logger, driver)
institutional.execute()
# BuyingStock.csvに値を追加する 61.6s → 12s
append_data = AppendData(logger)
append_data.execute()
# Excelに変換後視覚情報を調整 47.93s(セルの自動調整追加時)
convert_excel = ConvertExcel(logger)
convert_excel.execute()

end = time.time()
print(f"All Time : {round(end - start, 2)}")
