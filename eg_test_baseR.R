var1<-rnorm(mean = 0, sd = 1, 1000)
var2<-rnorm(mean = 17, sd = 4, 1000)
targ<-sample(2,1000, prob = c(0.8,0.2), replace = TRUE)-1

df<-data.frame(var1 = rnorm(mean = 0, sd = 1, 1000)
               , var2 = rnorm(mean = 17, sd = 4, 1000)
               , targ = sample(2,1000, prob = c(0.8,0.2), replace = TRUE)-1)

lm1<-glm(targ~var1+var2, data = df, family = "gaussian")