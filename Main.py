"""

    Authors: Gregory Mejia
    Date: 11/24/2025

"""

# -- Libraries -- #

import openpyxl  # type: ignore
from typing import TypeAlias

# -- Variables -- #

ShtEry: TypeAlias = tuple[str, str, str, int, int, int, float, float]

r1_score_path = 'CP18_RD1_Final_Scores.xlsx'
r1_scores = openpyxl.load_workbook(r1_score_path).active

# -- Functions -- #


def getDivisionScores(sheet, division: str,
                      state: str | None = None) -> list[ShtEry]:
    """
        Retrieve a list of all scores in a division and state (if providied)
        Used for filtering a spreadsheet provided by CyberPatriot
    """
    # Create a master list to store all the scores
    master_list = []

    # Loop over all the scores and only add if they're in the div or state
    for row in sheet.iter_rows(values_only=True):  # type: ignore
        if (state and (row[1] != state)):
            continue
        if (row[2] != division):
            continue

        # Add to master_list because it's in both
        master_list.append(row)
    return master_list


# -- Execution -- #

print(getDivisionScores(r1_scores, "Open"))
print("finish execution")
