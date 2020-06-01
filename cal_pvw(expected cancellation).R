# install.packages('openxlsx') # パッケージのインストール（最初だけ）
library(openxlsx)  # パッケージのロード
options(scipen=100) #指数表現を回避

# 生存率,死亡率,解約率を算出
qx <- read.xlsx("Lifetablem.xlsx") # 年齢,死亡率,解約率データを読み込み

qxdz <- qx[ ,2] #絶対死亡率
qxwz <- qx[ ,3] #絶対解約率

qxd <- qxdz*(1-0.5*qxwz) #死亡率
qxw <- qxwz*(1-0.5*qxdz) #解約率
qxt <- qxd + qxw #全脱退率

px <- 1 - qxt  # 生存率

# 生存数、脱退数を算出
#生存数を計算
lx <- 100000*cumprod(c(1, px))
length(lx) <- length(px) 

#脱退数を計算
dx <- lx*qxd #死亡数
wx <- lx*qxw #解約数
dxt <- dx + wx #脱退数

#lx - c(lx,0)[2:(length(lx)+1)]

data.frame(qxd,qxw,qxt,px,lx,dx,wx,dxt)

# 死亡解約脱退残存表の作成
data.frame(年齢 = qx[ ,1], 死亡率 = qxd, 脱退率 = qxw, 生存率 = px, 生存数 = lx, 死亡数 = dx, 解約数 = wx, 脱退数 = dxt)

# 予定利率の設定、現価率の算出
i <- 0.005
v <- 1/(1+i)

# 基数表の作成
Dx <- v^(0:109)*lx
Nx <- rev(cumsum(rev(Dx)))

Cxd <- v^(0:109+0.5)*dx
Cxw <- v^(0:109+0.5)*wx
Mxd <- rev(cumsum(rev(Cxd)))
Mxw <- rev(cumsum(rev(Cxw)))

data.frame(Dx,Nx,Cxd,Cxw,Mxd,Mxw)

# 一時払終身保険のPVW算出 解約したケースを考慮してません（死亡保険金にのみ対応）
# Pの算出
P <-function(x,S){
  (Mxd[x+1]/Dx[x+1])*S
}
# Vの算出
V <- function(x,t,S){
  (Mxd[x+t+1]/Dx[x+t+1])*S
}
# Wの算出
W <- function(x,t,S){
  (Mxd[x+t+1]/Dx[x+t+1])*S
}

#PVW算出例
format(P(6,10000000), big.mark=",", scientific=F)
format(V(5,1,10000000), big.mark=",", scientific=F)
format(W(5,1,10000000), big.mark=",", scientific=F)
