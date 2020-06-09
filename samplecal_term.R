# 用意されてなかったら必要なものです。(用意されていれば無視してください。)
install.packages("devtools")
install.packages("openxlsx")
install.packages("remotes")

#指数表現を回避します（必要に応じて）
options(scipen=100)　

# ここから必要な処理です。------------------------------

# 各種必要なパッケージのロード
library(devtools)
library(openxlsx)
library(remotes)

# 予定利率の設定
i <- 0.005    #直接入力

# 計算基数表の作成
install_github("Shoichi-Hoshino-46/CreateCommTBL")
library(CreateCommTBL)
# 男性
qxm <- as.matrix(read.xlsx("Lifetablem.xlsx"))
comm_man <- cal_comm(qxm,i)
# 女性
qxw <- as.matrix(read.xlsx("Lifetablew.xlsx"))
comm_woman <- cal_comm(qxw,i)
comm_woman

# PVW算出パッケージのロード
install_github("Shoichi-Hoshino-46/CalPVW.term")
library(CalPVW.term)

# サンプル(保有契約情報を読み込んで責準を算出)
pol <- read.xlsx("満期パターン一覧(5000).xlsm",sheet="OUT")

# 年齢
agex <- pol[ ,14]

# 保険期間
per <- pol[ ,16]
mat <- ifelse(per == 1, pol[ ,17]/100, pol[ ,17]/100 - pol[ ,14])

# 経過期間
dur <- 1      #直接入力

# 保険金
sum_ins <- pol[ ,9]

# 性別で場合分け
sex <- pol[ ,12]
ifelse(sex == 1, comm_man, comm_woman)

cal_polreserve <- v(agex, mat, dur, sum_ins) #契約者毎の責準算出

polreserve <- data.frame(契約番号 = pol[ ,1],責任準備金 = cal_polreserve)
reservecurve <- plot(polreserve,xlim = c(0,5080), type ="l", main ="責準曲線")

write.xlsx(polreserve, file="output.xlsx", sheetName="sheet1", row.names=F)

# 責準曲線の出力
png("polreserve.png")
plot(polreserve,xlim = c(0,5080), type ="l", main ="責準曲線")
dev.off()