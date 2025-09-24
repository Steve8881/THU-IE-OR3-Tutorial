# io_pandas.py


import pandas as pd
import constant as C


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
    try:
        # header=None tells pandas that our file doesn't have a header row,
        # so it will use default integer indices for columns.
        data = pd.read_excel(C.DATA_PATH, sheet_name=C.SHEET_NAME, header=None)
    except FileNotFoundError:
        # This block runs if the python script cannot find 'wyndor.xlsx'.
        print(f"Error: The file {C.DATA_PATH} was not found.")
        return None, None, None
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None, None, None

    # 1. Extract Product Profits
    # We use .at[row, col] for fast and precise access to a single value.
    # programming indices start at 0, so we subtract 1.
    product_profit = {
        C.DOORS: data.at[C.PROFIT_ROW - 1, C.DOORS_COL - 1],
        C.WINDOWS: data.at[C.PROFIT_ROW - 1, C.WINDOWS_COL - 1],
    }

    # 2. Extract Available Hours Per Plant
    plant_available_hour = {}  
    
    # We loop through the list of plant names defined in constant.py.
    # The `enumerate` function gives us both the index (0, 1, 2, ...) and the value.
    for i, plant_name in enumerate(C.PLANT_NAMES):
        # Calculate the correct row in the Excel sheet.
        row_index = C.HOURS_START_ROW - 1 + i
        # Get the value from the specific cell for available hours.
        hours = data.at[row_index, C.HOURS_COL - 1]
        # Add the new key-value pair to our dictionary.
        plant_available_hour[plant_name] = hours

    # 3. Extract Hour Requirements (Plant-Centric Structure)
    # We will build a nested dictionary: {plant: {product: hours}}.
    plant_product_hour = {} # Start with an empty dictionary.

    # We loop through each plant name, just like we did before.
    for i, plant_name in enumerate(C.PLANT_NAMES):
        # Calculate the current row index.
        row_index = C.HOURS_START_ROW - 1 + i
        
        # For each plant, we need to get the hours for Doors and Windows.
        hours_for_doors = data.at[row_index, C.DOORS_COL - 1]
        hours_for_windows = data.at[row_index, C.WINDOWS_COL - 1]
        
        # Create the inner dictionary for this specific plant.
        product_hours = {
            C.DOORS: hours_for_doors,
            C.WINDOWS: hours_for_windows,
        }
        
        # Add the inner dictionary to our main dictionary, using the plant name as the key.
        plant_product_hour[plant_name] = product_hours

    print("Data read successfully (io_pandas) with plant-centric structure.")
    return product_profit, plant_product_hour, plant_available_hour


# Writes the optimal production plan to a new sheet in the Excel file.
def writeData(products_solution, total_profit):
    """Args:
        products_solution (dict): A dictionary with the optimal production amount for each product.
            e.g., {"Doors": 2.0, "Windows": 6.0}
        total_profit (float): The final, optimal total profit calculated by the model.
    """
    try:
        # Create a pandas DataFrame from the solution data. A DataFrame is like a table.
        output_data = pd.DataFrame(
            {
                C.DOORS: [products_solution.get(C.DOORS, 0)],
                C.WINDOWS: [products_solution.get(C.WINDOWS, 0)],
                "Total Profit": [total_profit],
            }
        )

        # To add a new sheet without deleting existing ones, we use ExcelWriter.
        with pd.ExcelWriter(
            C.DATA_PATH,
            engine="openpyxl",  # Specifies the library to use for writing.
            mode="a",            # "a" stands for "append" mode.
            if_sheet_exists="replace",  # If the sheet already exists, replace it.
        ) as writer:
            # Write our DataFrame to the specified sheet.
            # index=False prevents pandas from writing the DataFrame's row index (0)
            # into the first column of the Excel sheet.
            output_data.to_excel(
                writer, sheet_name=C.WRITE_SHEET_NAME_PANDAS, index=False
            )
        
        print(
            f"Solution successfully written to sheet '{C.WRITE_SHEET_NAME_PANDAS}' "
            f"in '{C.DATA_PATH}'."
        )
    except FileNotFoundError:
        print(
            f"Error: The file {C.DATA_PATH} was not found. "
            "Could not write the solution."
        )
    except Exception as e:
        print(f"An error occurred while writing to the Excel file: {e}")

        