# 各種必要なパッケージのロード
library(openxlsx)

# 保有契約情報の読み込み
pol <- read.xlsx("満期パターン一覧(5000).xlsm",sheet="OUT")

# 全契約の出力
write.xlsx(pol, file="output_all.xlsx", sheetName="sheet1", row.names=F)

# 条件分岐(保険金額5万円以下はエラー扱い)
error_data <- pol[pol$SAMT <= 50000,]
correct_data <- pol[pol$SAMT > 50000,]

# エラー契約と正常契約の出力
write.xlsx(error_data, file="output_error.xlsx", sheetName="sheet1", row.names=F)
write.xlsx(correct_data, file="output_correct.xlsx", sheetName="sheet1", row.names=F)
