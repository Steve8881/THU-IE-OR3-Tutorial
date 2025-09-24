"""
Wyndor Production Planning Optimization Main Program
"""

from io_openpyxl import read_data_openpyxl
from io_pandas import read_data_pandas, print_results
from model import formulate_model, solve_model, extract_dual_variables


def main():
    """
    Main function: Execute linear programming solution process
    """
    print("Wyndor Production Planning Optimization")
    
  # Read data from Excel file
    # Two methods available: read_data_openpyxl() or read_data_pandas()
    # Students can modify this line to try different data reading approaches
    products_data, products_plants_data, plants_data = read_data_openpyxl()
    
    # Build model
    model = formulate_model(products_data, products_plants_data, plants_data)
    
    # Solve model
    solution_results = solve_model(model)
    
    # Extract dual variables
    dual_variables = extract_dual_variables(model)
    
    # Print results
    print_results(solution_results, dual_variables, products_data, plants_data)


if __name__ == "__main__":
    main()
