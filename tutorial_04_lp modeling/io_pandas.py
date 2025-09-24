"""
Reading from and writing to an Excel file using pandas.
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

    # Read data from the inputSheet.
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # Create dictionaries based on start coordinates and list lengths.
    productProfits = {}
    startRow = constant.INPUT_PROFIT_START_ROW - 1
    startCol = constant.INPUT_PROFIT_START_COL - 1
    for j, productName in enumerate(constant.PRODUCT_NAMES):
        profitValue = data.iloc[startRow, startCol + j]
        productProfits[productName] = profitValue

    plantAvailableHours = {}
    startRow = constant.INPUT_HOURS_AVAILABLE_START_ROW - 1
    startCol = constant.INPUT_HOURS_AVAILABLE_START_COL - 1
    for i, plantName in enumerate(constant.PLANT_NAMES):
        hours = data.iloc[startRow + i, startCol]
        plantAvailableHours[plantName] = hours

    plantProductHours = {}
    startRow = constant.INPUT_HOURS_START_ROW - 1
    startCol = constant.INPUT_HOURS_START_COL - 1
    for i, plantName in enumerate(constant.PLANT_NAMES):
        hoursPerProduct = {}
        for j, productName in enumerate(constant.PRODUCT_NAMES):
            hours = data.iloc[startRow + i, startCol + j]
            hoursPerProduct[productName] = hours
        plantProductHours[plantName] = hoursPerProduct

    return productProfits, plantProductHours, plantAvailableHours


def writeDataPandas(soln, objVal) -> None:
    """
    Write the solution back to the original Excel sheet using pandas.

    Parameters
    ----------
    soln : dict
        The optimal solutions (key is the full variable name).
    objVal : float
        The optimal objective function value.
    """
    # Read data from the inputSheet.
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # Write the batches produced solutions.
    startRow = constant.OUTPUT_BATCHES_PRODUCED_START_ROW - 1
    startCol = constant.OUTPUT_BATCHES_PRODUCED_START_COL - 1
    for j, product in enumerate(constant.PRODUCT_NAMES):
        varName = product
        value = soln.get(product, 0)
        data.loc[startRow, startCol + j] = value

    # Write the total profit solution.
    profitRow = constant.OUTPUT_PROFIT_START_ROW - 1
    profitCol = constant.OUTPUT_PROFIT_START_COL - 1
    data.loc[profitRow, profitCol] = objVal

    # Save to Excel file.
    data.to_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, index=False, header=False)

