"""

    Authors: Gregory Mejia
    Date: 11/24/2025

"""

# -- Libraries -- #

import openpyxl
from typing import cast

# -- Functions -- #


def forceSpreadSheetFileExtension(string: str) -> str:
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


def createSpreadSheet(sheet_title: str, file_name: str = "untitled.xlsx"):
    """
        Creates a new spreadsheet with the given file_name.
        Styled to be like the CyberPatriot's default sheet.
    """
    # Ignore all the MemberAccess and Subscript warnings from openpyxl
    # Create a new workbook and set the name to a valid version
    file_name = forceSpreadSheetFileExtension(file_name)
    workbook = openpyxl.Workbook()

    # Setting the main spreadsheet sheet to the main one and giving it a title
    sheet = workbook.active
    sheet.title = sheet_title

    # Headers 'n Stuff (boringg...)
    sheet["A10"] = "Team Number"
    sheet["B10"] = "Location"
    sheet["C10"] = "Division"
    sheet["D10"] = "Image Score"
    sheet["E10"] = "Adjustment"
    sheet["F10"] = "Quiz"
    sheet["G10"] = "PT"
    sheet["H10"] = "Total"

    # Save to the file to ensure every change is reflected in a new file
    workbook.save(file_name)
