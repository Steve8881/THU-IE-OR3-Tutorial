"""
Constants
"""

# Data file path and worksheet name
DATA_PATH = "./data/wyndor.xlsx"
SHEET_NAME = "Data"

LP_PATH = "./instance/wyndor.lp"
MPS_PATH = "./instance/wyndor.mps"

# Profit per batch data location (Doors & Windows)
INPUT_PROFIT_START_ROW = 4
INPUT_PROFIT_START_COL = 2

# Hours used per batch produced data location (3 plants)
INPUT_HOURS_START_ROW = 7
INPUT_HOURS_START_COL = 2

# Hours available data location (3 plants)
INPUT_HOURS_AVAILABLE_START_ROW = 7
INPUT_HOURS_AVAILABLE_START_COL = 6

# Output result location
OUTPUT_BATCHES_PRODUCED_START_ROW = 12
OUTPUT_BATCHES_PRODUCED_START_COL = 2

OUTPUT_PROFIT_START_ROW = 12
OUTPUT_PROFIT_START_COL = 6

# Product names list
PRODUCT_NAMES = ["Doors", "Windows"]

# Plant names list
PLANT_NAMES = ["Plant 1", "Plant 2", "Plant 3"]

# Model name
MODEL_NAME = "Wyndor"
