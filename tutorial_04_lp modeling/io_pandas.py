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

    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    productProfitCols = {
        'Doors': constant.INPUT_PROFIT_DOORS_COL,
        'Windows': constant.INPUT_PROFIT_WINDOWS_COL
    }
    productHourCols = {
        'Doors': constant.INPUT_HOURS_DOORS_COL,
        'Windows': constant.INPUT_HOURS_WINDOWS_COL
    }

    productProfits = {}
    for productName in constant.PRODUCT_NAMES:
        colIndex = productProfitCols[productName]
        profitValue = data.at[constant.INPUT_PROFIT_ROW - 1, colIndex - 1]
        productProfits[productName] = profitValue

    plantAvailableHours = {}
    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.INPUT_HOURS_AVAILABLE_START_ROW - 1 + i
        colIndex = constant.INPUT_HOURS_AVAILABLE_COL - 1
        hours = data.at[rowIndex, colIndex]
        plantAvailableHours[plantName] = hours

    plantProductHours = {}
    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.INPUT_HOURS_START_ROW - 1 + i
        hoursPerProduct = {}
        for productName in constant.PRODUCT_NAMES:
            colIndex = productHourCols[productName]
            hours = data.at[rowIndex, colIndex - 1]
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
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    doorsVarName = f'{constant.MODEL_VAR_NAME_PREFIX}{constant.PRODUCT_NAMES[0]}'
    windowsVarName = f'{constant.MODEL_VAR_NAME_PREFIX}{constant.PRODUCT_NAMES[1]}'

    outputValues = {
        constant.OUTPUT_DOORS_COL: soln.get(doorsVarName, 0),
        constant.OUTPUT_WINDOWS_COL: soln.get(windowsVarName, 0),
        constant.OUTPUT_PROFIT_COL: objVal
    }

    rowIndex = constant.OUTPUT_ROW - 1

    for col_constant, value in outputValues.items():
        colIndex = col_constant - 1
        data.loc[rowIndex, colIndex] = value

    data.to_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, index=False, header=False)