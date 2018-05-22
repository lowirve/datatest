# -*- coding: utf-8 -*-
"""
Created on Mon May 21 11:23:03 2018

@author: XuBo
"""
from __future__ import division, print_function

import xlrd, os

import pandas as pd
import numpy as np

name = '4692318-1334500'

directory = r'C:\Users\xub\Desktop'

directory = os.path.join(directory, name+'.xlsx')

data = pd.read_excel(directory, sheet_name=0)

