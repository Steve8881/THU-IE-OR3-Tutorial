from openpyxl import load_workbook
import constant


def readData():
    """
    Given an Excel file, read production planning data from Excel file using openpyxl.
    
    Returns
    -------
    productProfits : dict
        The profit per batch for each product
        - Keys: product names (`Doors`, `Windows`)
        - Values: profit per batch
    plantProductHours : nested dict
        The hours required to produce one batch of each product at each plant
        - Keys: plant names (`Plant 1`, `Plant 2`, `Plant 3`)
        - Values: dict
            - Keys: product names (`Doors` and `Windows`)
            - Values: hours required
    plantAvailableHours : dict
        The available hours at each plant
        - Keys: plant names (`Plant 1`, `Plant 2`, `Plant 3`)
        - Values: available hours
    """

    # Load a workbook from DATA_PATH
    inputBook = load_workbook(constant.DATA_PATH)

    # Find the sheet in which you want to write the solution
    inputSheet = inputBook[constant.SHEET_NAME]

    # Read data from the inputSheet and create dictionaries
    productProfits = {constant.DOORS: inputSheet.cell(constant.PROFIT_ROW, constant.DOORS_COL).value,
                      constant.WINDOWS: inputSheet.cell(constant.PROFIT_ROW, constant.WINDOWS_COL).value
                      }

    plantProductHours = {
        plantName: {
            constant.DOORS: inputSheet.cell(constant.HOURS_START_ROW + i, constant.DOORS_COL).value,
            constant.WINDOWS: inputSheet.cell(constant.HOURS_START_ROW + i, constant.WINDOWS_COL).value
        }
        for i, plantName in enumerate(constant.PLANT_NAMES)
    }

    plantAvailableHours = {
        plantName: inputSheet.cell(constant.HOURS_START_ROW + i, constant.HOURS_COL).value
        for i, plantName in enumerate(constant.PLANT_NAMES)
    }

    return productProfits, plantProductHours, plantAvailableHours


def writeData(soln, objVal):
    """
    Write the solution back to the Excel file using openpyxl.

    Parameters
    ----------
    soln : dict
        The optimal solutions
        - Keys: variable names
        - Values: optimal decision variable values
    objVal : float
        The optimal objective function value
    """
    # Load a workbook from DATA_PATH
    outputBook = load_workbook(constant.DATA_PATH)

    # Find the sheet to write the solution
    outputSheet = outputBook[constant.SHEET_NAME]

    # Write the optimal solutions to the outputSheet
    outputSheet.cell(constant.OUTPUT_ROW, constant.OUTPUT_DOORS_COL, soln[constant.DOORS])
    outputSheet.cell(constant.OUTPUT_ROW, constant.OUTPUT_WINDOWS_COL, soln[constant.WINDOWS])
    outputSheet.cell(constant.OUTPUT_ROW, constant.OUTPUT_PROFIT_COL, objVal)

    # Save the workbook
    outputBook.save(constant.DATA_PATH)
