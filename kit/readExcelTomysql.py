#导入需要使用到的数据模块
import pandas as pd
import pymysql
import numpy as np

#读入数据
filepath ="C:\\Users\\dailiuyang\\Desktop\\电子档案EDocument\\在职人员工号20191231.xlsx"
#指定 读取sheet页 
sheetName='20191231数据'
data = pd.read_excel(filepath,sheet_name=sheetName)

#建立数据库连接
db = pymysql.connect('10.27.15.237','root','xxxx','xxxx')
cursor = db.cursor()

query=""" update xxx set study_minute=%s,townsman_count=%s,familyday=%s,hackerMatch=%s,OscarGoldenCup=%s,buy_house=%s,Recommend_count=%s,trip_date=%s WHERE usercode=%s"""
#迭代读取每行数据
#values中元素有个类型的强制转换，否则会出错的
#应该会有其他更合适的方式，可以进一步了解
print("===============>记录数:"+str(len(data)) +"<=================")
for r in range(0, len(data)):
    userName = data.ix[r,0]
    userCode = data.ix[r,1]
    joinDate = data.ix[r,2]
    
    hometwon=data.ix[r,4]
    if np.isnan(hometwon):
        hometwon=0
    neitui=data.ix[r,5]    
    if np.isnan(neitui):    
        neitui=0
    shengri=data.ix[r,3]
    anjia=data.ix[r,6]
    if anjia=='是':    
        anjia=1
    else:
        anjia=0    
    send_email=data.ix[r,7]
    receive_email=data.ix[r,8]
    
    study=data.ix[r,9]
    if np.isnan(study):
        study=0
    qinzi=data.ix[r,10]
    if np.isnan(qinzi):
        qinzi=0
    hacker=data.ix[r,11]
    if np.isnan(hacker):
        hacker=0
    oscar=data.ix[r,12]
    if np.isnan(oscar):
        oscar=0
    tuanjian=data.ix[r,13]
    
    values = ( int(study),int(hometwon),int(qinzi),int(hacker),int(oscar),int(anjia),int(neitui),str(tuanjian),str(userCode))
    print(r,values)
    cursor.execute(query, values)

cursor.close()
db.commit()
db.close()    

#关闭游标，提交，关闭数据库连接
#如果没有这些关闭操作，执行后在数据库中查看不到数据 


