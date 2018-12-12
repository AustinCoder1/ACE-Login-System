#!/usr/bin/env python
# coding: utf-8

# In[4]:


from my_module.Security import *
import string

x = securityEncryption()
assert x.hashPassword("teststring") == "d67c5cbf5b01c9f91932e3b8def5e5f8"
assert isinstance(x.hashPassword("teststring"), str)
assert callable(x.hashPassword)


# In[5]:


from my_module.ExcelFunctions import *

x = excelFunctions()
assert isinstance(x.getDataColumn(), int)
assert callable(x.getDataColumn)
assert callable(x.registerFunction)
assert callable(x.loginFunction)


# In[ ]:




