'''
Created on Nov 12, 2019

@author: phuongtruong
'''
from svi.view.Enum import DB_TAG
class ParsingHTML_TR(object):
    '''
    classdocs
    '''

    def __init__(self, tr):
        self.tr = tr
    def getalltr(self):
        self.tr =(DB_TAG.TR)     