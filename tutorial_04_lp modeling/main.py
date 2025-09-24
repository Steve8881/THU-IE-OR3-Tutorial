"""
Wyndor Production Planning Optimization Main Program
"""

from io_openpyxl import readDataOpenpyxl, writeDataOpenpyxl
from io_pandas import readDataPandas, writeDataPandas
from model import formulateModel, solveModel, getOptimalDualVariableValues, saveModel, readModel

import constant


def main():
    """
    Main function
    """

    #
    # Step 1: Formulate the LP model
    #

    # Option 1) Read data from the Excel file using openpyxl, then formulate the LP model

    # Read data
    productProfits, plantProductHours, plantAvailableHours = readDataOpenpyxl()

    # Formulate the LP model
    model = formulateModel(productProfits, plantProductHours, plantAvailableHours)

    # Option 2) Read data from the Excel file using pandas, then formulate the LP model

    # Read data
    # productProfits, plantProductHours, plantAvailableHours = readDataPandas()

    # Formulate the LP model
    # model = formulateModel(productProfits, plantProductHours, plantAvailableHours)

    # Option 3) Directly load the LP from the MPS or LP file

    # model = readModel(constant.MPS_PATH)
    # model = readModel(constant.LP_PATH)

    #
    # Step 2: Save the LP model to an LP or an MPS file (optional)
    #

    # Save the model
    # saveModel(model, constant.MPS_PATH)
    # saveModel(model, constant.LP_PATH)

    #
    # Step 3: Solve the LP model
    #

    soln, objVal = solveModel(model)

    #
    # Step 4: Extract dual variables (optional)
    #

    # dualVariables = getOptimalDualVariableValues(model)

    #
    # Step 5: Write the solution to the Excel file
    #

    # Option 1) Use openpyxl
    writeDataOpenpyxl(soln, objVal)

    # Option 2) Use pandas
    # writeDataPandas(soln, objVal)


if __name__ == "__main__":
    main()
