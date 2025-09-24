import pandas as pd
from constant import DOORS, WINDOWS, PLANT_NAMES


def print_results(solution_results, dual_variables, products_data, plants_data):
    """
    Print solution results and dual variables
    Using pandas for data processing and formatted output
    
    Args:
        solution_results: Solution results dictionary
        dual_variables: Dual variables dictionary
        products_data: Product data (dict)
        plants_data: Plant data (dict)
    """
    # Convert dictionary data to pandas DataFrame for processing
    products_df = pd.DataFrame([products_data])
    plants_df = pd.DataFrame.from_dict(plants_data, orient='index', columns=['Available Hours'])
    
    # Print optimal solution
    print("\n【Optimal Solution】")
    print(f"Doors production quantity: {solution_results[DOORS]:.2f}")
    print(f"Windows production quantity: {solution_results[WINDOWS]:.2f}")
    print(f"Maximum profit: ${solution_results['objective_value']:.2f}")
    
    # Use pandas to format product information output
    print("\n【Product Profit Information】")
    print(products_df.to_string(index=False))
    
    # Use pandas to format plant information output
    print("\n【Plant Resource Information】")
    print(plants_df.to_string())
    
    # Print dual variables (shadow prices)
    print("\n【Dual Variables (Shadow Prices)】")
    for i, plant_name in enumerate(PLANT_NAMES):
        dual_value = dual_variables.get(f'plant_{i+1}', 0)
        print(f"{plant_name}: ${dual_value:.2f}")


def read_data_pandas():
    """
    Read production planning data from Excel file using pandas
    
    Returns:
        tuple: Tuple containing three dictionaries
            - products_profits: Product profit data (dict)
            - products_plants_hours: Product hours requirement data for each plant (dict)
            - plants_available_hours: Available hours data for each plant (dict)
    """
    from constant import (
        DATA_PATH, SHEET_NAME,
        PROFIT_ROW, DOORS_COL, WINDOWS_COL,
        HOURS_START_ROW, HOURS_END_ROW, HOURS_COL,
        DOORS, WINDOWS, PLANT_NAMES
    )
    
    # Read Excel file using pandas
    df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME, header=None)
    
    # Read product profit data
    products_profits = {
        DOORS: df.iloc[PROFIT_ROW-1, DOORS_COL-1],
        WINDOWS: df.iloc[PROFIT_ROW-1, WINDOWS_COL-1]
    }
    
    # Read product hours requirement data for each plant
    products_plants_hours = {}
    for i, plant_name in enumerate(PLANT_NAMES):
        row = HOURS_START_ROW - 1 + i
        products_plants_hours[plant_name] = {
            DOORS: df.iloc[row, DOORS_COL-1],
            WINDOWS: df.iloc[row, WINDOWS_COL-1]
        }
    
    # Read available hours data for each plant
    plants_available_hours = {}
    for i, plant_name in enumerate(PLANT_NAMES):
        row = HOURS_START_ROW - 1 + i
        plants_available_hours[plant_name] = df.iloc[row, HOURS_COL-1]
    
    return products_profits, products_plants_hours, plants_available_hours