from smartcap import SmartCap, redcap
import logging
import os
import sys


cwd = os.path.dirname(__file__)
logpath = os.path.join(cwd, 'log.txt')
logging.basicConfig(filemode = 'w', format='%(levelname)s:%(asctime)s:%(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename= logpath, level=logging.INFO)
logging.info('\n')
logging.info('*****************Program Start**********************')
