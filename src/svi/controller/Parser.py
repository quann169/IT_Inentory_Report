'''
Created on Nov 11, 2019

@author: phuongtruong
'''

from svi.view.Enum import DB_TAG, DB_FILE
from svi.model import ParsingHTML_Table

from svi.model.ParsingHTML_Table import ParsingHTML_Table

#get table

parsing_TABLE = ParsingHTML_Table(DB_TAG.TABLE)
parsing_TABLE.getalltable()
parsing_TABLE.table  
    