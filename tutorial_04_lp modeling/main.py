"""
Wyndor Production Planning Optimization Main Program
"""

from io_openpyxl import readDataOpenpyxl, writeDataOpenpyxl
from io_pandas import readDataPandas, writeDataPandas
from model import formulateModel, solveModel, getOptimalDualVariableValues


def main():
    """
    Main function: Execute linear programming solution process
    """
    print("Wyndor Production Planning Optimization")

    # Read data from Excel file
    # Two methods available: from io_openpyxl import readDataOpenpyxl OR from io_pandas import readDataPandas
    # You can modify the "from ··· import ··" line to try different data reading methods
    productsData, productsPlantsData, plantsData = readDataOpenpyxl()

    # Build model
    model = formulateModel(productsData, productsPlantsData, plantsData)

    # Solve model
    soln, objValue = solveModel(model)

    # Extract dual variables if use this code
    # dualVariables = getOptimalDualVariableValues(model)

    # Write results back to Excel
    writeDataOpenpyxl(soln, objValue)


if __name__ == "__main__":
    main()
