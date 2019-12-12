'''
Created on Nov 12, 2019

@author: phuongtruong
'''
from svi.view.Enum import DB_TAG
class ParsingHTML_TD(object):
    '''
    classdocs
    '''

    def __init__(self, td):
        self.td = td
    def getalltd(self):
        self.td =(DB_TAG.TD)     