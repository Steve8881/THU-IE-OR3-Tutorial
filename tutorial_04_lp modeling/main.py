"""
Wyndor Production Planning Optimization Main Program
"""

from io_openpyxl import readDataOpenpyxl, writeDataOpenpyxl
from io_pandas import readDataPandas, writeDataPandas
from model import formulateModel, solveModel, getOptimalDualVariableValues
DATA_PATH = 'wyndor.xlsx'
def main():
    """
    Main function: Execute linear programming solution process
    """
    print("Wyndor Production Planning Optimization")
    
    # Read data from Excel file
    # Two methods available: from io_openpyxl import readDataOpenpyxl OR from io_pandas import readDataPandas
    # You can modify the "from ··· import ··" line to try different data reading methods
    productsData, productsPlantsData, plantsData = readDataOpenpyxl(DATA_PATH)

    # Build model
    model = formulateModel(productsData, productsPlantsData, plantsData)
    
    # Solve model
    solutionResults, objValue = solveModel(model)
    
    # Extract dual variables
    dualVariables = getOptimalDualVariableValues(model)

    # Write results back to Excel
    writeDataOpenpyxl(solutionResults, objValue)

if __name__ == "__main__":
    main()
