import math
import eikon as ek  # the Eikon Python wrapper package
ek.set_app_key('f46a38d4f9424f59b900d6d7b34d6fbe2ea681e5')
import numpy as np  # NumPy
import pandas as pd  # pandas
import cufflinks as cf  # Cufflinks
import configparser as cp
import scipy.optimize as sco  # optimization routines
cf.set_config_file(offline=True)  # set the plotting mode to offline
import os
path = os.getcwd()

#input
df_a = pd.read_csv(path/'agency.csv')
df_ut = pd.read_csv(path/'Tbill.csv')
df = pd.concat([df_a,df_ut])
df['CUSIP']=df['CUSIP'].str.strip("'")

#get_Dur_data,大きすぎて分割
df_Maturity_0,err = ek.get_data(list(df['CUSIP'][:5000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')
df_Maturity_1,err = ek.get_data(list(df['CUSIP'][5000:10000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')
df_Maturity_2,err = ek.get_data(list(df['CUSIP'][10000:15000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')
df_Maturity_3,err = ek.get_data(list(df['CUSIP'][15000:20000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_4,err = ek.get_data(list(df['CUSIP'][20000:25000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_5,err = ek.get_data(list(df['CUSIP'][25000:30000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_6,err = ek.get_data(list(df['CUSIP'][30000:35000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_7,err = ek.get_data(list(df['CUSIP'][35000:40000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_8,err = ek.get_data(list(df['CUSIP'][40000:45000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_9,err = ek.get_data(list(df['CUSIP'][45000:50000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_10,err = ek.get_data(list(df['CUSIP'][50000:55000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_11,err = ek.get_data(list(df['CUSIP'][55000:60000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
df_Maturity_12,err = ek.get_data(list(df['CUSIP'][60000:65000]),  # the RICs
                         fields='TR.FiMaturityCorpModDuration')  # the required fields
#agg
df_Maturity_all =pd.concat([df_Maturity_0,df_Maturity_1,df_Maturity_2,df_Maturity_3,\
                            df_Maturity_4,df_Maturity_5,df_Maturity_6,df_Maturity_7,\
                            df_Maturity_8,df_Maturity_9,df_Maturity_10,df_Maturity_11,\
                            df_Maturity_12])
#agg2
df_matome = pd.merge(df,df_Maturity_all,left_on='CUSIP',right_on='Instrument')
#select
df_kani = df_matome[['Security Type','Modified Duration','Current Face Value','Par Value']]
df_kani['Value'] = df_kani.sum(axis=1,skipna=True)
df_kani['label'] = pd.cut(df_kani['Modified Duration'],bins=[-0.00001,0.6,1,5,10,50],labels=['-6M','6M-1Y','1Y-5Y','5Y-10Y','10Y-'])
df_kani = df_matome[['Security Type','Modified Duration','Value','label']]
df_kani.pivot_table(index="label", columns= "Security Type", values="Value")

#output_kani_ver
df_matome.to_excel(path/'out.xlsx')
#output_pivot
df_kani.pivot_table(index="label", columns= "Security Type", values="Value").to_excel(path/'pivot.xlsx')