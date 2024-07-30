from enum import Enum 

REPORT_SHEET_NAME = 'Scheduled-Report-Darwin---Detai'

class REPORT_HEADERS(Enum):
    NetBios = 4
    Title = 9
    Severity = 12
    Solution = 36
    Results = 37

class LAPTOP_HEADERS(Enum):
    Vulnerability = 1

class INTUNE_HEADERS(Enum):
    DeviceName = 2
    Category = 25
    UserName = 27

class RESULT_HEADERS(Enum):
    Country = 1
    UserName = 2
    LaptopName = 3
    Title = 4
    Severity = 5
    Solution = 6
    Results = 7

class ERR_MSG(Enum):
    Intune_ws_err = 'Unable to find Intune worksheet in intune.xlsx'
    Main_ws_err = 'Unable to find Main worksheet in report.xlsx'
    Laptop_ws_err = 'Unable to find Laptop worksheet in report.xlsx!'