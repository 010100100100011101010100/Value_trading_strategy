#flow -> importing all modules -> reading stock -> creating dataframe and calling batch api -> filling up missing data -> calculating percentiles and rv score -> keeping top 50 stocks and calculting number of shares to buy -> excel manipulation

#importing stocks
import numpy as np
import pandas as pd
import math as h
from scipy.stats import percentileofscore as per
import requests
from statistics import mean
from secrets import IEX_CLOUD_API_TOKEN
import xlsxwriter 

#reading stock
stock=pd.read_csv('sp_500_stocks.csv')

#creating dataframe and calling api
def chunk(l,n):
    for i in range(0,len(l),n):
        yield l[i:i+n]
m=['Ticker','Price','number of shares','pe ratio','pe percentile','ps ratio','ps percentile','pb ratio','pb percentile','ebitda ratio','ebitda percentile','evgp ratio','evgp percentile','rv score']
final_dataframe=pd.DataFrame(columns=m)
symbol_group=list(chunk(stock,100))
symbol_str=[]
for s in range(len(symbol_group)):
    symbol_str.append(','.join(symbol_group[s]))
for si in symbol_str:
    batch=f'https://sandbox.iexapis.com/stable//stock/market/batch&symbols={s}&types=quote,advanced-stats&token={IEX_CLOUD_API_TOKEN}&'
    data=requests.get(batch).json()
    for s in si.split(':'):
        ev=data[s]['advanced-stats']['enterpriseValue']
        ebitda=data[s]['advanced-stats']['EBITDA']
        gp=data[s]['advanced-stats']['grossProfit']
        try:
            ev_ebitda=ev/ebitda
        except TypeError :
            ev_ebitda=np.NaN
        try:
            ev_gp=ev/gp
        except TypeError :
            ev_gp=np.NaN
        final_dataframe=final_dataframe.append(pd.Series([
            s,
            data[s]['quote']['latestPrice'],
            'na',
            data[s]['advanced-stats']['forwardPERatio'],
            'na',
            data[s]['advanced-stats']['priceToSales'],
            'na',
            data[s]['advanced-stats']['priceToBook'],
            'na',
            ev_ebitda,
            'na',
            ev_gp,
            'na',
            'na'
            ],index=m),ignore_index=True)
final_dataframe

#looking at the missing data from the 5 types of ratio columns
final_dataframe[final_dataframe.isnull().any(axis=1)]
#filling the empty column blocks with the mean value of the present column values
col=['pe ratio','ps ratio','pb ratio','ebitda ratio','evgp ratio']
for c in col:
    final_dataframe[c].fillna(final_dataframe[c].mean(),inplace=True)

#calculating the percentile and rv score
metrics={
    'pe ratio':'pe percentile',
    'ps ratio':'ps percentile',
    'pb ratio':'pb percentile',
    'ebitda ratio':'ebitda percentile',
    'evgp ratio':'evgp percentile'
}
#percentile calculated
for i in final_dataframe.index:
    for m in metrics.keys():
        final_dataframe.loc[i,metrics[m]] = per( final_dataframe[m],final_dataframe.loc[i,m])/100
#rv score
for i in final_dataframe.index:
    val=[]
    for m in metrics.keys():
        val.append(final_dataframe.loc[i,metrics[m]])
    final_dataframe.loc[i,'rv score']=mean(val)
final_dataframe

#keeping top 50 value stocks based on rv score
final_dataframe.sort_values('rv score',ascending=True,inplace=True)
final_dataframe=final_dataframe[:50]
final_dataframe.reset_index(drop=True,inplace=True)

#portfolio and position_size
portfolio=float(input())
position_size=portfolio/len(final_dataframe.index)

#calculating the number of shares for these top 50 value shares
for i in final_dataframe.index:
    final_dataframe.loc[i,'number of shares']=h.floor(position_size/final_dataframe.loc[i,'Price'])

#excel manipulation
writer=pd.ExcelWriter('value tradin.xlsx',engine='xlsxwriter')
final_dataframe.to_excel(writer,sheet_name='value trading',index=False)

font='#ffffff'
bg='#0a0a23'

string=writer.book.add_format({
    'bg_color':bg,
    'font_color':font,
    'border':1
})

integer=writer.book.add_format({
    'num_format':'0',
    'bg_color':bg,
    'font_color':font,
    'border':1
})

dollar=writer.book.add_format({
    'num_format':'$0.0',
    'bg_color':bg,
    'font_color':font,
    'border':1
})

percentage=writer.book.add_format({
    'num_format':'0.0%',
    'bg_color':bg,
    'font_color':font,
    'border':1
})

floa=writer.book.add_format({
    'num_format':'0.0',
    'bg_color':bg,
    'font_color':font,
    'border':1
})

#creating column format
c={
    'A':['Ticker',string],
    'B':['Price',dollar],
    'C':['number of shares',integer],
    'D':['pe ratio',floa],
    'E':['pe percentile',percentage],
    'F':['ps ratio',floa],
    'G':['ps percentile',percentage],
    'H':['pb ratio',floa],
    'I':['pb percentile',percentage],
    'J':['ebitda ratio',floa],
    'K':['ebitda percentile',percentage],
    'L':['evgp ratio',floa],
    'M':['evgp percentile',percentage],
    'N':['rv score',floa]

}

#writing and setting column formats
for w in c.keys():
    writer.sheets['value trading'].set_column(f'{w}:{w}',25,c[w][1])
    writer.sheets['value trading'].write(f'{w}1',c[w][0],string)
writer.save()

    


                      
