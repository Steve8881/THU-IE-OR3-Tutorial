"""
Wyndor Production Planning Optimization Main Program for only I/O operations
"""

from io_openpyxl import readDataOpenpyxl, writeDataOpenpyxl
from io_pandas import readDataPandas, writeDataPandas


def main():
    """
    Main function: Execute linear programming solution process
    """
    print("Wyndor Production Planning Optimization")

    # Read data
    # Two methods available: readDataOpenpyxl() and readDataPandas()
    # productProfits, plantProductHours, plantAvailableHours = readDataOpenpyxl()
    productProfits, plantProductHours, plantAvailableHours = readDataPandas()

    # Write results back to Excel
    # Two methods available: writeDataOpenpyxl() and writeDataPandas()
    # Skip modeling and solving steps
    soln = {"Product 1": 2, "Product 2": 6}  # Example solution
    objVal = 36000  # Example objective value
    
    # writeDataOpenpyxl(soln, objVal)
    writeDataPandas(soln, objVal)


if __name__ == "__main__":
    main()
