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
    
    # Read data from the inputSheet
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # Create dictionaries
    productProfits = {}
    for productName in constant.PRODUCT_NAMES:
        profitColConst = getattr(constant, f"INPUT_PROFIT_{productName.upper()}_COL")
        profitValue = data.iloc[constant.INPUT_PROFIT_ROW - 1, profitColConst - 1]
        productProfits[productName] = profitValue

    plantAvailableHours = {}
    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.INPUT_HOURS_AVAILABLE_START_ROW - 1 + i
        colIndex = constant.INPUT_HOURS_AVAILABLE_COL - 1
        hours = data.iloc[rowIndex, colIndex]
        plantAvailableHours[plantName] = hours

    plantProductHours = {}
    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.INPUT_HOURS_START_ROW - 1 + i
        hoursPerProduct = {}
        for productName in constant.PRODUCT_NAMES:
            hourColConst = getattr(constant, f"INPUT_HOURS_{productName.upper()}_COL")
            hours = data.iloc[rowIndex, hourColConst - 1]
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
    # Read data from the inputSheet
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # Build the outputValues dictionary
    outputValues = {}
    for product in constant.PRODUCT_NAMES:
        varName = f'{constant.MODEL_VAR_NAME_PREFIX}{product}'
        outputColConst = getattr(constant, f"OUTPUT_{product.upper()}_COL")
        outputValues[outputColConst] = soln.get(varName, 0)

    # Add the total profit to the dictionary.
    outputValues[constant.OUTPUT_PROFIT_COL] = objVal

    rowIndex = constant.OUTPUT_ROW - 1

    for col_constant, value in outputValues.items():
        colIndex = col_constant - 1
        data.loc[rowIndex, colIndex] = value

    data.to_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, index=False, header=False)