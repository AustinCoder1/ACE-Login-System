#!/usr/bin/env python
# coding: utf-8

# In[1]:


import hashlib
import os, sys

class securityEncryption():
    def hashPassword(self, password):
        md5 = hashlib.md5()
        md5.update(bytes(password, 'utf-8'))
        return md5.hexdigest()

