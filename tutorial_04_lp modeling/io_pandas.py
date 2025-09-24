"""
input and output data from Excel file using pandas.
"""

import pandas as pd
import constant


def readDataPandas() -> tuple[dict, dict, dict]:
    """
    Read production planning data from an Excel file using pandas.

    Returns
    -------
    productProfits : dict
        The profit per batch for each product.
    plantProductHours : nested dict
        The hours required to produce one batch of each product at each plant.
    plantAvailableHours : dict
        The available hours at each plant.
    """
    # Read data from the Excel sheet into a pandas DataFrame.
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # 1. Read product profits from the inputSheet using a loop.
    productProfits = {}
    # We define a helper dictionary to map product names to their column numbers.
    productColumns = {
        constant.DOORS: constant.DOORS_COL,
        constant.WINDOWS: constant.WINDOWS_COL,
    }
    # Read data from the inputSheet and create dictionaries
    for productName, colIndex in productColumns.items():
        profitValue = data.at[constant.PROFIT_ROW - 1, colIndex - 1]
        productProfits[productName] = profitValue

    plantAvailableHours = {}

    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.HOURS_START_ROW - 1 + i
        colIndex = constant.HOURS_COL - 1
        hours = data.at[rowIndex, colIndex]
        plantAvailableHours[plantName] = hours

    plantProductHours = {}
    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.HOURS_START_ROW - 1 + i
        hoursPerProduct = {}
        for productName, colIndex in productColumns.items():
            hours = data.at[rowIndex, colIndex - 1]
            hoursPerProduct[productName] = hours
        plantProductHours[plantName] = hoursPerProduct

    return productProfits, plantProductHours, plantAvailableHours


def writeDataPandas(productsSolution, totalProfit) -> None:
    """
    Write the solution back to the Excel file using pandas.

    Parameters
    ----------
    productsSolution : dict
        The optimal solutions.
    totalProfit : float
        The optimal objective function value.
    """
    # Create a pandas DataFrame from the solution.
    outputData = pd.DataFrame(
        {
            constant.DOORS: [productsSolution.get(constant.DOORS, 0)],
            constant.WINDOWS: [productsSolution.get(constant.WINDOWS, 0)],
            "Total Profit": [totalProfit],
        }
    )

    # Write the solution DataFrame to a specific sheet.
    outputData.to_excel(
        constant.DATA_PATH,
        sheet_name=constant.WRITE_SHEET_NAME_PANDAS,
        index=False,
    )
    