# install.packages('openxlsx') # パッケージのインストール（最初だけ）
library(openxlsx)  # パッケージのロード
options(scipen=100) #指数表現を回避

# 生存率、死亡率を算出
qx <- read.xlsx("Lifetablem.xlsx") # 2018年標準生命表(男性)の年齢と死亡率データを読み込み
px <- 1 - qx[ ,2]  # 生存率を計算

#生存数を計算
lx <- 100000*cumprod(c(1, px))

#死亡数を計算
length(lx) <- length(px)
dx <- lx - c(lx,0)[2:(length(lx)+1)]

# 平均余命を算出
Lx <- rev(cumsum(rev(lx)))
ex <- -0.5 + Lx/lx

# 予定利率の設定、現価率の算出
i <- 0.005
v <- 1/(1+i)

# 基数表の作成
Dx <- v^(0:109)*lx
Cx <- v^(0:109+0.5)*dx
Nx <- rev(cumsum(rev(Dx)))
Mx <- rev(cumsum(rev(Cx)))
data.frame(Dx,Cx,Nx,Mx)

# 一時払い終身保険のPVW算出

# Prateの算出
Prate <-function(x){
  (Mx[x+1]/Dx[x+1])
}

# Vrateの算出
Vrate <- function(x,t){
  (Mx[x+t+1]/Dx[x+t+1])
}

# Wrateの算出
Wrate <- function(x,t){
  (Mx[x+t+1]/Dx[x+t+1])
}
#PVW算出例
format(Prate(30)*10000000, big.mark=",", scientific=F)
format(Vrate(30,10)*10000000, big.mark=",", scientific=F)
format(Wrate(30,10)*10000000, big.mark=",", scientific=F)