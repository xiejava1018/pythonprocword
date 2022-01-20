# -*- coding: utf-8 -*-
"""
    :author: XieJava
    :url: http://ishareread.com
    :copyright: © 2021 XieJava <xiejava@ishareread.com>
    :license: MIT, see LICENSE for more details.
"""
import os
import pandas as pd
from docx import Document

data=[]
#读word的docx评议表文件，并读取word中的表格数据
def procdoc(docfilepath):
    document=Document(docfilepath)
    tables=document.tables
    table=tables[0]
    for i in range(1,len(table.rows)):
        id=int(table.cell(i,0).text)
        name=table.cell(i,1).text
        excellent=0
        if table.cell(i,2).text!='' and table.cell(i,2).text is not None:
            excellent=1
        competent = 0
        if table.cell(i, 3).text!='' and table.cell(i, 3).text is not None:
            competent=1
        basicacompetent=0
        if table.cell(i, 4).text!='' and table.cell(i, 4).text is not None:
            basicacompetent=1
        notcompetent = 0
        if table.cell(i, 5).text!='' and table.cell(i, 5).text is not None:
            notcompetent=1
        dontunderstand =0
        if table.cell(i, 6).text!='' and table.cell(i, 6).text is not None:
            dontunderstand=1
        appraisedata=[id,name,excellent,competent,basicacompetent,notcompetent,dontunderstand]
        data.append(appraisedata)

#读取评议表的目录，并处理目录中的docx文件，根据评议表计算评分，写入汇总表。
def readfile(filepah):
    files=os.listdir(filepah)
    for file in files:
        if file.find('.docx')>0:
            docfilepah=filepah+file
            procdoc(docfilepah)
    df = pd.DataFrame(data,columns=['序号','姓名','优秀','称职','基本称职','不称职','不了解'])
    print(df)
    df=df.groupby(['序号','姓名']).sum()
    df['票数'] = df.apply(lambda x: x.sum(), axis=1)
    df['计分'] = (df['优秀']*95+df['称职']*85+df['基本称职']*75+df['不称职']*65+df['不了解']*0)/len(df)
    df['评价']=df['计分'].map(getscore)
    print(df)
    write2excle('民主评议\\民主评议表汇总.xlsx',df)

#根据评分规则计算评级
def getscore(x):
    if x>=95:
        score='优秀'
    elif x>=80 and x<95:
        score='称职'
    elif x>=75 and x<80:
        score='基本称职'
    elif x<75:
        score='不称职'
    return score

#将汇总计算好的数据写入Excel
def write2excle(exclefile,dataframe):
    writer = pd.ExcelWriter(exclefile)
    dataframe.to_excel(writer)
    writer.save()
    print('输出成功')

if __name__ == '__main__':
    readfile('民主评议\\')