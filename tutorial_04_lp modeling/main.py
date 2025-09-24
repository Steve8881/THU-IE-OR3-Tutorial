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
    print("="*40)
    
    # Let students choose data reading method
    print("\nChoose data reading method:")
    print("1. OpenPyXL (direct Excel access)")
    print("2. Pandas (DataFrame-based)")
    
    while True:
        choice = input("Enter your choice (1 or 2): ").strip()
        if choice == "1":
            print("Using OpenPyXL to read data...")
            products_data, products_plants_data, plants_data = read_data_openpyxl()
            break
        elif choice == "2":
            print("Using Pandas to read data...")
            products_data, products_plants_data, plants_data = read_data_pandas()
            break
        else:
            print("Invalid choice. Please enter 1 or 2.")
    
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