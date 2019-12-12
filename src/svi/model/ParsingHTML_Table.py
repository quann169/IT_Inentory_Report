'''
Created on Nov 12, 2019

@author: phuongtruong
'''
from svi.view.Enum import DB_TAG

class ParsingHTML_Table():
    '''
    classdocs
    '''
    def __init__(self, table):
        self.table = table
    def getalltable(self):
        self.table =(DB_TAG.TABLE)   