import pandas as pd
import constant


def readData() -> tuple[dict, dict, dict]:
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
        'Doors': constant.INPUT_DOORS_PROFIT_COL,
        'Windows': constant.INPUT_WINDOWS_PROFIT_COL
    }
    productHourCols = {
        'Doors': constant.INPUT_DOORS_HOURS_COL,
        'Windows': constant.INPUT_WINDOWS_HOURS_COL
    }

    # Read data from the Excel sheet into a pandas DataFrame.
    productProfits = {}
    for productName in constant.PRODUCT_NAMES:
        colIndex = productProfitCols[productName]
        profitValue = data.at[constant.INPUT_PROFIT_START_ROW - 1, colIndex - 1]
        productProfits[productName] = profitValue

    plantAvailableHours = {}
    for i, plantName in enumerate(constant.PLANT_NAMES):
        rowIndex = constant.INPUT_HOURS_AVAILABLE_START_ROW - 1 + i
        colIndex = constant.INPUT_HOURS_AVAILABLE_START_COL - 1
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


def writeData(soln, objVal) -> None:
    """
    Write the solution back to the original Excel sheet using pandas.

    Parameters
    ----------
    soln : dict
        The optimal solutions (key is the full variable name).
    objVal : float
        The optimal objective function value.
    """
    # read the entire sheet
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # Note: We construct the full variable name to look up the value in the solution dictionary.
    doorsVarName = f'{constant.MODEL_VAR_NAME_PREFIX}{constant.PRODUCT_NAMES[0]}'
    windowsVarName = f'{constant.MODEL_VAR_NAME_PREFIX}{constant.PRODUCT_NAMES[1]}'
    
    outputValues = {
        constant.OUTPUT_DOORS_COL: soln.get(doorsVarName, 0),
        constant.OUTPUT_WINDOWS_COL: soln.get(windowsVarName, 0),
        constant.OUTPUT_PROFIT_COL: objVal
    }

    rowIndex = constant.OUTPUT_ROW - 1

    # Loop through the prepared dictionary and update each cell in the DataFrame.
    for col_constant, value in outputValues.items():
        colIndex = col_constant - 1
        data.loc[rowIndex, colIndex] = value

    # Write DataFrame back to the Excel sheet.
    data.to_excel(
        constant.DATA_PATH,
        sheet_name=constant.SHEET_NAME,
        index=False,
        header=False
    )

