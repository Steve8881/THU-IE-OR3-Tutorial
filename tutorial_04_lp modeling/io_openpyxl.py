from openpyxl import load_workbook
import constant as C


def generateDataOpenpyxl():
    """
    Given an excel file, read production planning data from Excel file using openpyxl.
    """
    # Load a workbook from DATA_PATH
    inputBook = load_workbook(C.DATA_PATH)

    # Find the sheet in which you want to write the solution
    inputSheet = inputBook[C.SHEET_NAME]

    # Read data from the inputSheet and create dictionaries
    productProfits = {C.DOORS: inputSheet.cell(C.PROFIT_ROW, C.DOORS_COL).value,
                      C.WINDOWS: inputSheet.cell(C.PROFIT_ROW, C.WINDOWS_COL).value
                      }

    plantProductHours = {
        plantName: {
            C.DOORS: inputSheet.cell(C.HOURS_START_ROW + i, C.DOORS_COL).value,
            C.WINDOWS: inputSheet.cell(C.HOURS_START_ROW + i, C.WINDOWS_COL).value
        }
        for i, plantName in enumerate(C.PLANT_NAMES)
    }

    plantAvailableHours = {
        plant_name: inputSheet.cell(C.HOURS_START_ROW + i, C.HOURS_COL).value
        for i, plant_name in enumerate(C.PLANT_NAMES)
    }

    return productProfits, plantProductHours, plantAvailableHours


def writeDataOpenpyxl(solution, optimalObjectiveValue):
    """
    Write the solution back to the Excel file using openpyxl.
    """
    # Load a workbook from DATA_PATH
    outputBook = load_workbook(C.DATA_PATH)

    # Find the sheet to write the solution
    outputSheet = outputBook[C.SHEET_NAME]

    # Write the optimal solutions to the outputSheet
    outputSheet.cell(C.OUTPUT_ROW, C.OUTPUT_DOORS_COL, solution[C.DOORS])
    outputSheet.cell(C.OUTPUT_ROW, C.OUTPUT_WINDOWS_COL, solution[C.WINDOWS])
    outputSheet.cell(C.OUTPUT_ROW, C.OUTPUT_PROFIT_COL, optimalObjectiveValue)

    # Save the workbook
    outputBook.save(C.DATA_PATH)
