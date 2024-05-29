#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import sqlite3



conn = sqlite3.connect('C:/Users/tan7080/Documents/FAST/no_data.db')
print('done')


# In[2]:


df_prod = pd.read_csv('prod_ofm.csv')
df_prod = df_prod.rename(columns={'net_oil':'OIL','WaterProduced':'WATER','gas':'GAS','run_time':'RUNTIME'})
df_prod


# In[3]:


df_fluid = pd.read_csv('fluid_levelofm.csv')
df_fluid = df_fluid.rename(columns={'fluid_level':'fluid_depth','pi_x':'pip'})
df_fluid


# In[4]:


df_test = pd.read_csv('test_ofm.csv')
df_test = df_test.rename(columns={'date':'Date','gross_test':'GrossTest','net_oil_test':'NetTest','wc-Test':'WcTest'})
df_test


# In[5]:


df_inj = pd.read_csv('C:/Users/tan7080/Desktop/production reports/my_update/inj.csv')
df_inj = df_inj.iloc[:,1:]
df_inj = df_inj.rename(columns={'act_inj':'inj_daily'})
df_inj


# In[6]:


df_prod.to_sql('daily_production',conn,if_exists='append',index = False)
print('production done')


# In[7]:


df_inj.to_sql('daily_injection',conn,if_exists='append',index = False)
print('injection done')


# In[8]:


df_test.to_sql('well_test',conn,if_exists='append',index = False)
print('test done')


# In[9]:


df_fluid.to_sql('fluid_level',conn,if_exists='append',index = False)
print('fluid done')


# In[10]:


df_event = pd.read_csv('events_ofm.csv')
df_event = df_event.rename(columns={'events':'event'})


# In[11]:


df_event


# In[12]:


df_event.to_sql('remarks',conn,if_exists='append',index = False)
print('events done')


# In[ ]:


df_dpr = pd.read_csv('C:/Users/tan7080/Desktop/production reports/my_update/dpr.csv')
df_dpr = df_dpr.iloc[:,1:]


# In[ ]:


df_dpr.to_sql('dpr_tmp',conn,if_exists='append',index = False)
print('dpr done')


# In[13]:


conn.close()

df_dpr.to_csv('summary is ready.csv')


# In[ ]:




