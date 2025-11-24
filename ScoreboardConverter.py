"""

    Authors: Gregory Mejia
    Date: 11/24/2025

"""

# -- Libraries -- #

import openpyxl

# -- Functions -- #


def forceSpreadsheetFileExtension(string: str) -> str:
    """
        Force parameter 'string' to have .xlsx at the end of its name.
        Used to ensure a valid file type from the file extension.
    """
    try:
        dot_pos: int = string.index(".")
        if (string[dot_pos:-1] == ".xlsx"):
            return string
        else:
            return string[0:dot_pos] + "xlsx"
    except ValueError:
        return string + ".xlsx"


def createSpreadSheet(file_name: str = "untitled.xlsx"):
    """
        Creates a new spreadsheet with the given file_name.
        Styled to be like the CyberPatriot's default sheet.
    """
    file_name = forceSpreadsheetFileExtension(file_name)
