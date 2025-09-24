# io_pandas.py



import pandas as pd
import constant as C

# Reads production planning data from the Excel file.
def generate_data():

    try:
        data = pd.read_excel(C.DATA_PATH, sheet_name=C.SHEET_NAME, header=None)
    except FileNotFoundError:
        print(f"Error: The file {C.DATA_PATH} was not found.")
        return None, None, None
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None, None, None

    # 1. Extract product profits from specified cells
    product_profit = {
        C.DOORS: data.at[C.PROFIT_ROW - 1, C.DOORS_COL - 1],
        C.WINDOWS: data.at[C.PROFIT_ROW - 1, C.WINDOWS_COL - 1],
    }

    # 2. Extract available hours per plant from a column range
    hour_rows = slice(C.HOURS_START_ROW - 1, C.HOURS_END_ROW)
    available_hours_series = data.loc[hour_rows, C.HOURS_COL - 1]
    plant_available_hour = dict(zip(C.PLANT_NAMES, available_hours_series))

    # 3. Extract hour requirements for each product at each plant
    product_plant_hour = {
        C.DOORS: dict(
            zip(C.PLANT_NAMES, data.loc[hour_rows, C.DOORS_COL - 1])
        ),
        C.WINDOWS: dict(
            zip(C.PLANT_NAMES, data.loc[hour_rows, C.WINDOWS_COL - 1])
        ),
    }

    print("Data read successfully (io_pandas).")
    return product_profit, product_plant_hour, plant_available_hour

# Writes the optimal production plan to the Excel file.
def write_data(products_solution, total_profit):
   
    try:
        # Create a DataFrame to structure the output data
        output_data = pd.DataFrame(
            {
                C.DOORS: [products_solution.get(C.DOORS, 0)],
                C.WINDOWS: [products_solution.get(C.WINDOWS, 0)],
                "Total Profit": [total_profit],
            }
        )

        # Use ExcelWriter in append mode to add/replace a sheet without
        # disturbing other sheets in the workbook
        with pd.ExcelWriter(
            C.DATA_PATH,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            output_data.to_excel(
                writer, sheet_name=C.WRITE_SHEET_NAME_PANDAS, index=False
            )
        print(
            f"Solution written to sheet '{C.WRITE_SHEET_NAME_PANDAS}' "
            f"in '{C.DATA_PATH}'."
        )
    except FileNotFoundError:
        print(
            f"Error: The file {C.DATA_PATH} was not found. "
            "Could not write the solution."
        )
    except Exception as e:
        print(f"An error occurred while writing to the Excel file: {e}")
