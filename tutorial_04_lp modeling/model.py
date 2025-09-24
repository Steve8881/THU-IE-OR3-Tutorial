"""
Linear programming model construction and solving module
Using Gurobi solver to build and solve Wyndor production planning problem
"""

import gurobipy as grb
from constant import MODEL_NAME, DOORS, WINDOWS, PLANT_NAMES


def formulate_model(products_data, products_plants_data, plants_data):
    """
    Build linear programming model
    
    Args:
        products_data: Product profit data (dict)
        products_plants_data: Product hours requirement data (dict)
        plants_data: Plant available hours data (dict)
    
    Returns:
        grb.Model: Built Gurobi model
    """
    # Create model
    model = grb.Model(MODEL_NAME)
    
    # Create decision variables
    doors = model.addVar(name=DOORS, vtype=grb.GRB.CONTINUOUS, lb=0)
    windows = model.addVar(name=WINDOWS, vtype=grb.GRB.CONTINUOUS, lb=0)
    
    # Set objective function (maximize profit)
    profit_doors = products_data[DOORS]
    profit_windows = products_data[WINDOWS]
    model.setObjective(profit_doors * doors + profit_windows * windows, grb.GRB.MAXIMIZE)
    
    # Add constraints
    for i, plant_name in enumerate(PLANT_NAMES):
        # Get hours requirement
        hours_doors = products_plants_data[plant_name][DOORS]
        hours_windows = products_plants_data[plant_name][WINDOWS]
        # Get available hours
        available_hours = plants_data[plant_name]
        
        # Add constraint
        model.addConstr(
            hours_doors * doors + hours_windows * windows <= available_hours,
            name=f'plant_{i+1}'
        )
    
    return model


def solve_model(model):
    """
    Solve linear programming model
    
    Args:
        model: Gurobi model object
    
    Returns:
        dict: Solution results dictionary
    """
    # Solve model
    model.optimize()
    
    # Extract results
    results = {
        DOORS: model.getVarByName(DOORS).x,
        WINDOWS: model.getVarByName(WINDOWS).x,
        'objective_value': model.objVal
    }
    
    return results


def extract_dual_variables(model):
    """
    Extract dual variables (shadow prices)
    
    Args:
        model: Solved Gurobi model object
    
    Returns:
        dict: Dual variables dictionary
    """
    dual_variables = {}
    
    # Extract dual variables of constraints
    for constr in model.getConstrs():
        dual_variables[constr.constrName] = constr.pi
    
    return dual_variables