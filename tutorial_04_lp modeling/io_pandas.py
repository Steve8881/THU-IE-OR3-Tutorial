
# io_pandas.py


import pandas as pd
import constant


def readData():
    """
    Given an Excel file, read production planning data from Excel file using pandas.

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
    # Read data from the Excel sheet into a pandas DataFrame.
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # Read data from the inputSheet and create dictionaries
    product_profit = {
        constant.DOORS: data.at[constant.PROFIT_ROW - 1, constant.DOORS_COL - 1],
        constant.WINDOWS: data.at[constant.PROFIT_ROW - 1, constant.WINDOWS_COL - 1],
    }

    plant_product_hour = {
        plant_name: {
            constant.DOORS: data.at[constant.HOURS_START_ROW - 1 + i, constant.DOORS_COL - 1],
            constant.WINDOWS: data.at[constant.HOURS_START_ROW - 1 + i, constant.WINDOWS_COL - 1],
        }
        for i, plant_name in enumerate(constant.PLANT_NAMES)
    }

    plant_available_hour = {
        plant_name: data.at[constant.HOURS_START_ROW - 1 + i, constant.HOURS_COL - 1]
        for i, plant_name in enumerate(constant.PLANT_NAMES)
    }

    return product_profit, plant_product_hour, plant_available_hour


def writeData(products_solution, total_profit):
    """
    Write the solution back to the Excel file using pandas.

    Parameters
    ----------
    soln : dict
        The optimal solutions
        - Keys: variable names
        - Values: optimal decision variable values
    objVal : float
        The optimal objective function value
    """
    # Create a pandas DataFrame from the solution.
    output_data = pd.DataFrame(
        {
            constant.DOORS: [products_solution.get(constant.DOORS, 0)],
            constant.WINDOWS: [products_solution.get(constant.WINDOWS, 0)],
            "Total Profit": [total_profit],
        }
    )

    # Write the solution to the specified sheet.
    with pd.ExcelWriter(
        constant.DATA_PATH,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        output_data.to_excel(
            writer, sheet_name=constant.WRITE_SHEET_NAME_PANDAS, index=False
        )




