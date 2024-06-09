from utils import *
import time


start = time.time()

# 初期実行 1.21s
init_process = InitProcess()
logger = init_process.set_log()
driver = init_process.set_selenium()
# init_process.execute()
# # finvizから銘柄の配列を取得 87.18s 1.45m
# tmp_start = time.time()
# pick_finviz = PickFinviz(logger)
# pick_finviz.execute()
# print(round(time.time() - tmp_start, 2))
# # ヒストリカルデータをダウンロードする 3800.04s 63.33m
# tmp_start = time.time()
# hist_data = HistData(logger, driver)
# hist_data.execute()
# print(round(time.time() - tmp_start, 2))
# # NASDAQのヒストリカルデータの列を増やす
# tmp_start = time.time()
# process_nasdaq = ProcessNASDAQ(logger)
# process_nasdaq.execute()
# print(round(time.time() - tmp_start, 2))
# # 個別株のヒストリカルデータの列を増やす 610.45s 10.17m
# tmp_start = time.time()
# process_histdata = ProcessHistData(logger)
# process_histdata.execute()
# print(round(time.time() - tmp_start, 2))
# # RSを計算する 91.94s 1.53m
# tmp_start = time.time()
# calculate_rs = CalculateRS(logger)
# calculate_rs.execute()
# print(round(time.time() - tmp_start, 2))
# # BuyingStock.csvを出力する 0.3s
# tmp_start = time.time()
# buying_stock = BuyingStock(logger)
# buying_stock.execute()
# print(round(time.time() - tmp_start, 2))
# CとAを取得する 2874.17s 47.90m
tmp_start = time.time()
current_annual = CurrentAnnual(logger, driver)
current_annual.execute()
print(round(time.time() - tmp_start, 2))
# 機関投資家の増加数を取得  2345.98s 39.09m
tmp_start = time.time()
institutional = Institutional(logger, driver)
institutional.execute()
print(round(time.time() - tmp_start, 2))
# BuyingStock.csvに値を追加する 61.6s → 12s
tmp_start = time.time()
append_data = AppendData(logger)
append_data.execute()
print(round(time.time() - tmp_start, 2))
# Excelに変換後視覚情報を調整 47.93s(セルの自動調整追加時)
tmp_start = time.time()
convert_excel = ConvertExcel(logger)
convert_excel.execute()
print(round(time.time() - tmp_start, 2))


end = time.time()
print(f"All Time : {round(end - start, 2)}")
