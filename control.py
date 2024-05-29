#!/usr/bin/env python
# coding: utf-8

# In[1]:


# Import the necessary modules
import subprocess

# Run the scripts using subprocess module
subprocess.run(['python', 'ofm_update_pressure.py'])
print('done_fluid_level')

subprocess.run(['python', 'ofm_update_production.py'])
print('done_prod')

subprocess.run(['python', 'ofm_update_test.py'])
print('done_test')

subprocess.run(['python', 'ofm_update_events.py'])
print('done_events')


# In[ ]:




