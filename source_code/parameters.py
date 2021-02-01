##### PARAMETERS SCRIPT

## IMPORTS FOR ALL SCRIPTS
import tkinter as t
from tkinter import ttk, filedialog, messagebox
from ttkthemes import themed_tk as tk
from icon_bitmap import ICON
from win32com.client import Dispatch
import pyodbc as db
import openpyxl as xl
from openpyxl.styles import Alignment
import sqlite3 as lite
import os
from report import *
from pdf import *



PROGRAM = 'Amabilis'
VERSION = ' v1.0'

#### MAIN PROGRAM FORMATTING
BLACK = "#000000"
GREY = "#D7D7D7"
RED = "#FF0000"
DARKRED = "#A10C0C"

FORESTGREEN = "#418A52"

WIDTH = 1500
HEIGHT = 750

SCREEN_POS_RIGHT = 5
SCREEN_POS_DOWN = 5

#'Calibri'
#'Candara'

font_name = 'Calibri'

font8Cb = (font_name, "8", "bold")
font10Cb = (font_name, "10", "bold")
font11Cb = (font_name, "11", "bold")
font12Cb = (font_name, "12", "bold")
font14Cb = (font_name, "14", "bold")
font18Cb = (font_name, "18", "bold")
font36Cb = (font_name, "36", "bold")


METRIC_LIST = ["STAND", "PLOT #", "TREE #", "SPECIES", "DBH", "TOTAL HGT", "PLOT FACTOR"]


METRIC_DICT = {"STAND": ["Enter the name of the stand for example: 'EX1'."],
               "PLOT #": ["Enter the plot number within the stand."],
               "TREE #": ["Enter the tree number within the plot."],
               "SPECIES": "",
               "DBH": ["Enter the DBH of the tree."],
               "TOTAL HGT": ["Enter the total height of the tree.", "Note that only one height is needed per stand",
                             "but more heights will produce more reliable data."],
               "PLOT FACTOR": ["Variable-Radius Plots:",  "\tEnter Basal Area Factor", "",
                               "Fixed-Radius Plots:", "\tEnter the negative-inverse of", "\tthe radius (1/30th ac = -30)."]}




## ACCESS DATA TYPES
vc = 'varchar'
dbl = 'double'
sht = 'short'
lt = 'longtext'

## SQLITE3 DATA TYPES
tt = 'TEXT'
it = 'INTEGER'
rl = 'REAL'

## EXCEL SHEET NAMES
FVS_G = 'FVS_GroupAddFilesAndKeywords'
FVS_S = 'FVS_StandInit'
FVS_T = 'FVS_TreeInit'


## FVS TABLES AND COLUMN VALUES
FVS_KEYWORDS = "Database\r\nDSNIn\r\n{FILLER_DB_NAME}\r\nStandSQL\r\nSELECT *\r\nFROM FVS_StandInit\r\nWHERE Stand_ID = '%StandID%'\r\nEndSQL\r\nTreeSQL\r\nSELECT *\r\nFROM FVS_TreeInit\r\nWHERE Stand_ID = '%StandID%'\r\nEndSQL\r\nEND"


EXCEL_KEYWORDS = "Database\nDSNIn\n{FILLER_DB_NAME}\nStandSQL\nSELECT * FROM FVS_StandInit WHERE Stand_ID =\n'%StandID%'\nEndSQL\nTreeSQL\nSELECT * FROM FVS_TreeInit WHERE Stand_ID =\n'%StandID%'\nEndSQL\nEND"



FVS_GROUPS = {FVS_G: [['Groups', vc, tt, 'All_Stands'],
                      ['Addfiles', vc, tt, ''],
                      ['FVSKeywords', lt, tt, FVS_KEYWORDS]]}


FVS_STANDS = {FVS_S: [['Stand_ID', vc, tt], ['Variant', vc, tt], ['Inv_Year', sht, it],
                      ['Groups', vc, tt], ['Latitude', dbl, rl], ['Longitude', dbl, rl],
                      ['Region', sht, it], ['Forest', sht, it], ['Age', sht, it], ['ElevFt', sht, it],
                      ['Basal_Area_Factor', dbl, rl], ['Brk_DBH', dbl, rl], ['Num_Plots', sht, it],
                      ['Site_Species', vc, tt], ['Site_Index', sht, it]]}


FVS_TREES = {FVS_T: [['Stand_ID', vc, tt], ['StandPlot_ID', vc, tt], ['Plot_ID', sht, it],
                     ['Tree_ID', sht, it], ['Tree_Count', sht, it], ['History', sht, it],
                     ['Species', vc, tt], ['DBH', dbl, rl], ['Ht', dbl, rl]]}


STANDS_INPUT = [['Variant', 'Forest Code', 'Region Code'],
                ['Inventory Year', 'Latitude', 'Longitude', 'Age', 'Elevation', 'Site Species', 'Site Index']]



## FOREST CODES FOR EACH VARIANT
VARIANTS_LOCS = {'AK': ['Alaska', [1002, 1003, 1004, 1005]],
                 'BM': ['Blue Mountains', [604, 607, 614, 616, 619]],
                 'CA': ['Inland California', [505, 506, 508, 511, 514, 518, 610, 611]],
                 'CI': ['Central Idaho', [117, 402, 406, 412, 413, 414]],
                 'CR': ['Central Rockies', [202, 203, 204, 206, 207, 209, 210, 212, 213, 214, 215, 301, 302, 303, 304,
                                            305, 306, 307, 308, 309, 310, 312]],
                 'CS': ['Central States', [905, 908, 912]],
                 'EC': ['East Cascades', [603, 606, 608, 613, 617, 699]],
                 'EM': ['Eastern Montana', [102, 108, 109, 111, 112, 115]],
                 'IE': ['Inland Empire', [103, 104, 105, 106, 110, 113, 114, 116, 117, 118, 621]],
                 'LS': ['Lake States', [902, 903, 904, 906, 907, 909, 910, 913, 924]],
                 'NC': ['Klamath Mountains', [505, 510, 514, 611]],
                 'NE': ['Northeast', [914, 919, 920, 921]],
                 'PN': ['Pacific Northwest', [609, 612]],
                 'SN': ['Southern', [80101, 80103, 80104, 80105, 80106, 80107,
                                     80211, 80212, 80213, 80214, 80215, 80216, 80217,
                                     80301, 80302, 80304, 80305, 80306, 80307, 80308,
                                     80401, 80402, 80403, 80404, 80405, 80506,
                                     80501, 80502, 80504, 80505, 80506,
                                     80601, 80602, 80603, 80604, 80605,
                                     80701, 80702, 80704, 80705, 80706, 80717,
                                     80801, 80802, 80803, 80804, 80805, 80806,
                                     80811, 80812, 80813, 80814, 80815, 80816,
                                     80901, 80902, 80903, 80904, 80905, 80906,
                                     80907, 80908, 80909, 80910, 80911, 80912,
                                     81001, 81002, 81003, 81004, 81005, 81006, 81007,
                                     81102, 81103, 81105, 81106, 81107, 81108, 81109, 81110, 81111,
                                     81201, 81202, 81203, 81205,
                                     81301, 81303, 81304, 81307, 81308, 905, 908]],
                 'SO': ['SC Oregon NE California', [505, 506, 509, 511, 601, 602, 620]],
                 'TT': ['Tetons', [403, 405, 415, 416]],
                 'UT': ['Utah', [401, 407, 408, 410, 418, 419]],
                 'WC': ['Western Cascades', [603, 605, 606, 610, 615, 618]],
                 'WS': ['Western Sierra Nevada', [503, 511, 513, 515, 516, 517]]}


## DEFAULT LAT AND LONG FOR EACH VARIANT (Chosen from center of Variant geographic region)
LAT_LONG_DEFAULT = {'AK': (65.3488, -150.9184),
                    'BM': (44.9964, -118.3989),
                    'CA': (39.9440, -122.2660),
                    'CI': (43.7716, -115.1908),
                    'CR': (36.7042, -106.0942),
                    'CS': (39.6401, -90.5375),
                    'EC': (47.7250, -120.2885),
                    'EM': (47.6658, -105.1274),
                    'IE': (47.7250, -116.1576),
                    'LS': (45.4606, -90.4496),
                    'NC': (41.9355, -123.7602),
                    'NE': (41.7718, -75.0688),
                    'PN': (45.6018, -123.4371),
                    'SN': (34.5446, -87.9665),
                    'SO': (42.8761, -119.3443),
                    'TT': (42.4537, -110.5844),
                    'UT': (39.1391, -114.8926),
                    'WC': (45.9024, -121.4266),
                    'WS': (37.3916, -118.6263)}

                                                  





        
















    

