"""
    This file includes some constants used in other modules
"""

""" Errors that may happen in ptu files """

COMMENT_ERROR = 0
DEFINE_STUB_ERROR = 1
INITIALIZATION_ERROR = 2
ENVIRONMENT_ERROR = 3
SERVICE_ERROR = 4
TEST_ERROR = 5
ELEMENT_ERROR = 6
IF_ERROR = 7

""" Number of columns in some sheets used for exporting data to excel format """

ENVIRONMENT_WIDTH = 4
SERVICE_WIDTH = 2
ELEMENT_WIDTH = 5

""" Defining width for columns in test sheets for better reading and showing data """

TEST_SHEET_COLUMNS_WIDTH = [50, 15, 80, 40, 40]