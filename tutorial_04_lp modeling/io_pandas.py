
# io_pandas.py

import pandas as pd
import constant


# Reads the production planning data from the Excel file.
def readData():
    """
    tuple of dict or None
        If successful, returns a tuple of three dictionaries:
        - product_profit: Profit per batch for each product.
          e.g., {"Doors": 3000, "Windows": 5000}
        - plant_product_hour: Hours required by each product at each plant.
          e.g., {"Plant 1": {"Doors": 1, "Windows": 0}, "Plant 2": {"Doors": 0, "Windows": 2},"Plant 3": {"Doors": 3, "Windows": 2}}
        - plant_available_hour: Total hours available at each plant.
          e.g., {"Plant 1": 4, "Plant 2": 12, "Plant 2": 18}
        If the file is not found or an error occurs, returns (None, None, None).
    """
    # Read the entire Excel sheet into a pandas DataFrame.
    # header=None ensures that the first row is not treated as a header.
    data = pd.read_excel(constant.DATA_PATH, sheet_name=constant.SHEET_NAME, header=None)

    # 1. Extract product profits directly from specific cells.
    product_profit = {
        constant.DOORS: data.at[constant.PROFIT_ROW - 1, constant.DOORS_COL - 1],
        constant.WINDOWS: data.at[constant.PROFIT_ROW - 1, constant.WINDOWS_COL - 1],
    }

    # 2. Use a dictionary comprehension to get the available hours for each plant.
    # This is a concise way to create a dictionary from a loop.
    plant_available_hour = {
        plant_name: data.at[constant.HOURS_START_ROW - 1 + i, constant.HOURS_COL - 1]
        for i, plant_name in enumerate(constant.PLANT_NAMES)
    }

    # 3. Use a nested dictionary comprehension for the plant and product hours.
    plant_product_hour = {
        plant_name: {
            constant.DOORS: data.at[constant.HOURS_START_ROW - 1 + i, constant.DOORS_COL - 1],
            constant.WINDOWS: data.at[constant.HOURS_START_ROW - 1 + i, constant.WINDOWS_COL - 1],
        }
        for i, plant_name in enumerate(constant.PLANT_NAMES)
    }

    print("Data read successfully using pandas.")
    return product_profit, plant_product_hour, plant_available_hour


# Writes the optimal solution to a new sheet in the Excel file.
def writeData(products_solution, total_profit):
    """
    Args:
        products_solution (dict): The optimal amount for each product.
        total_profit (float): The final optimal total profit.
    """
    # Create a pandas DataFrame from the solution. This is the standard way
    # to prepare tabular data for writing with pandas.
    output_data = pd.DataFrame(
        {
            constant.DOORS: [products_solution.get(constant.DOORS, 0)],
            constant.WINDOWS: [products_solution.get(constant.WINDOWS, 0)],
            "Total Profit": [total_profit],
        }
    )

    # Use pd.ExcelWriter to add or replace a sheet without affecting others.
    # This is the standard and safest way to write with pandas.
    with pd.ExcelWriter(
        constant.DATA_PATH,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        output_data.to_excel(
            writer, sheet_name=constant.WRITE_SHEET_NAME_PANDAS, index=False
        )

    print(
        f"Solution successfully written to sheet '{constant.WRITE_SHEET_NAME_PANDAS}' "
        f"in '{constant.DATA_PATH}'."
    )




