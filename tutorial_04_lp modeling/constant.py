"""
Constants definition module
"""

# Data file path and worksheet name
DATA_PATH = 'wyndor.xlsx'
SHEET_NAME = 'Data'

# Profit data location
PROFIT_ROW = 4
DOORS_COL = 2
WINDOWS_COL = 3

# Hours requirement data location
HOURS_START_ROW = 7
HOURS_END_ROW = 9
HOURS_COL = 6

# Output result location
OUTPUT_ROW = 12
OUTPUT_DOORS_COL = 2
OUTPUT_WINDOWS_COL = 3
OUTPUT_PROFIT_COL = 6

# Product names
DOORS = 'Doors'
WINDOWS = 'Windows'

# Data labels
PROFIT_PER_BATCH = 'Profit per batch'
HOURS_AVAILABLE = 'Hours available'

# Plant names list
PLANT_NAMES = ['Plant 1', 'Plant 2', 'Plant 3']

# Model name
MODEL_NAME = "Wyndor"

# Decision variable name prefix
VAR_NAME_PREFIX = "Var_BatchProduced_"

# Constraint name prefix
CONSTR_NAME_PREFIX = "Constr_AvailableHours_"
