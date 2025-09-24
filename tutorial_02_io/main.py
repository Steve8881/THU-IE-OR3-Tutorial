"""
Wyndor Production Planning Optimization Main Program for only I/O operations
"""

from io_openpyxl import readDataOpenpyxl, writeDataOpenpyxl
from io_pandas import readDataPandas, writeDataPandas


def main():
    """
    Main function
    """

    #
    # Step 1: Read data from the Excel file
    #

    # Option 1) Use openpyxl
    productProfits, plantProductHours, plantAvailableHours = readDataOpenpyxl()

    # Option 2) Use pandas
    # productProfits, plantProductHours, plantAvailableHours = readDataPandas()

    print("Profit per batch:")
    print(productProfits, "\n")

    print('Hours used per batch produced at each plant:')
    print(plantProductHours, "\n")

    print('Available hours at each plant:')
    print(plantAvailableHours, "\n")

    #
    # Step 2: Write the solution to the Excel file
    #

    # An arbitrary solution and its objective function value
    soln = {"Doors": 0, "Windows": 0}
    objVal = 0

    # Option 1) Use openpyxl
    writeDataOpenpyxl(soln, objVal)

    # Option 2) Use pandas
    # writeDataPandas(soln, objVal)


if __name__ == "__main__":
    main()
