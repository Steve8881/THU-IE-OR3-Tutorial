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

    # Read data
    # Two methods available: readDataOpenpyxl() and readDataPandas()
    productProfits, plantProductHours, plantAvailableHours = readDataOpenpyxl()
    # productProfits, plantProductHours, plantAvailableHours = readDataPandas()

    # Build model
    model = formulateModel(productProfits, plantProductHours, plantAvailableHours)

    # Solve model
    soln, objVal = solveModel(model)

    # Extract dual variables
    # dualVariables = getOptimalDualVariableValues(model)

    # Write results back to Excel
    # Two methods available: writeDataOpenpyxl() and writeDataPandas()
    writeDataOpenpyxl(soln, objVal)
    # writeDataPandas(soln, objVal)


if __name__ == "__main__":
    main()
