"""
Excel data reading module
Using openpyxl library to read Wyndor production planning problem data from Excel file

Author: Teaching Team
Version: 1.0
Date: 2024
"""

from openpyxl import load_workbook
from constant import (
    DATA_PATH, SHEET_NAME,
    PROFIT_ROW, DOORS_COL, WINDOWS_COL,
    HOURS_START_ROW, HOURS_END_ROW, HOURS_COL,
    DOORS, WINDOWS, PLANT_NAMES
)


def read_data_openpyxl():
    """
    Read production planning data from Excel file using openpyxl
    
    Returns:
        tuple: Tuple containing three dictionaries
            - products_profits: Product profit data (dict)
            - products_plants_hours: Product hours requirement data for each plant (dict)
            - plants_available_hours: Available hours data for each plant (dict)
    """
    # Load Excel workbook
    input_book = load_workbook(DATA_PATH)
    input_sheet = input_book[SHEET_NAME]
    
    # Read product profit data
    products_profits = _read_products_profits(input_sheet)
    
    # Read product hours requirement data for each plant
    products_plants_hours = _read_products_plants_hours(input_sheet)
    
    # Read available hours data for each plant
    plants_available_hours = _read_plants_available_hours(input_sheet)
    
    return products_profits, products_plants_hours, plants_available_hours


def _read_products_profits(input_sheet):
    """
    Read product profit data
    
    Args:
        input_sheet: openpyxl worksheet object
        
    Returns:
        dict: Product profit data
    """
    products_profits = {
        DOORS: input_sheet.cell(PROFIT_ROW, DOORS_COL).value,
        WINDOWS: input_sheet.cell(PROFIT_ROW, WINDOWS_COL).value
    }
    
    return products_profits


def _read_products_plants_hours(input_sheet):
    """
    Read product hours requirement data for each plant
    
    Args:
        input_sheet: openpyxl worksheet object
        
    Returns:
        dict: Product hours requirement data, format: {plant_name: {product: hours}}
    """
    products_plants_hours = {}
    
    for i, plant_name in enumerate(PLANT_NAMES):
        row = HOURS_START_ROW + i
        products_plants_hours[plant_name] = {
            DOORS: input_sheet.cell(row, DOORS_COL).value,
            WINDOWS: input_sheet.cell(row, WINDOWS_COL).value
        }
    
    return products_plants_hours


def _read_plants_available_hours(input_sheet):
    """
    Read available hours data for each plant
    
    Args:
        input_sheet: openpyxl worksheet object
        
    Returns:
        dict: Plant available hours data, format: {plant_name: hours}
    """
    plants_available_hours = {}
    
    for i, plant_name in enumerate(PLANT_NAMES):
        row = HOURS_START_ROW + i
        plants_available_hours[plant_name] = input_sheet.cell(row, HOURS_COL).value
    
    return plants_available_hours
