"""
Constants definition module
"""

# Data file path and worksheet name
DATA_PATH = "wyndor.xlsx"
SHEET_NAME = "Data"

# Profit per batch data location (Doors & Windows)
INPUT_PROFIT_ROW = 4
INPUT_PROFIT_DOORS_COL = 2
INPUT_PROFIT_WINDOWS_COL = 3

# Hours used per batch produced data location (3 plants)
INPUT_HOURS_START_ROW = 7
INPUT_HOURS_DOORS_COL = 2
INPUT_HOURS_WINDOWS_COL = 3

# Hours available data location (3 plants)
INPUT_HOURS_AVAILABLE_START_ROW = 7
INPUT_HOURS_AVAILABLE_COL = 6


# Output result location
OUTPUT_ROW = 12
OUTPUT_DOORS_COL = 2
OUTPUT_WINDOWS_COL = 3
OUTPUT_PROFIT_COL = 6

# Product names list
PRODUCT_NAMES = ["Doors", "Windows"] 


# Plant names list
PLANT_NAMES = ["Plant 1", "Plant 2", "Plant 3"]

#
# Model related constants
#

# Model name
MODEL_NAME = "Wyndor"

# Decision variable name prefix
MODEL_VAR_NAME_PREFIX = "Var_BatchProduced_"

# Constraint name prefix
MODEL_CONSTR_NAME_PREFIX = "Constr_AvailableHours_"
