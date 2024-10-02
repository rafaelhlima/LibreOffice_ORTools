###############################################
# This file implements the solver component
###############################################

import sys
import unohelper
import uno
import time

from com.sun.star.sheet import XSolver, XSolverDescription
from com.sun.star.beans import XPropertySet, XPropertySetInfo, Property, PropertyValue
from com.sun.star.lang import XServiceInfo
from com.sun.star.table import CellAddress
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.uno.TypeClass import LONG


implementation_name = "org.libreoffice.comp.ORToolsLinear_Impl"
implementation_service = "org.libreoffice.comp.ORToolsLinear_Impl"

# List of supported properties
ortools_properties = {"NonNegative": "Assume variables as non-negative",
                      "Integer": "Assume variables as integer",
                      "Timeout": "Solving time limit (seconds)",
                      "RelativeGap": "Relative gap for optimality"}

uno_bool_type = uno.getTypeByName("boolean")
uno_long_type = uno.getTypeByName("long")
uno_double_type = uno.getTypeByName("double")

CONSTR_BINARY = uno.Enum("com.sun.star.sheet.SolverConstraintOperator", "BINARY")
CONSTR_INTEGER = uno.Enum("com.sun.star.sheet.SolverConstraintOperator", "INTEGER")
CONSTR_EQUAL = uno.Enum("com.sun.star.sheet.SolverConstraintOperator", "EQUAL")
CONSTR_LESS_EQUAL = uno.Enum("com.sun.star.sheet.SolverConstraintOperator", "LESS_EQUAL")
CONSTR_GREATER_EQUAL = uno.Enum("com.sun.star.sheet.SolverConstraintOperator", "GREATER_EQUAL")
CELLTYPE_VALUE = uno.Enum("com.sun.star.table.CellContentType", "VALUE")
CELLTYPE_FORMULA = uno.Enum("com.sun.star.table.CellContentType", "FORMULA")

MAIN_NODE = "ortools.Settings/EngineOptions"

class PropertySetInfo(unohelper.Base, XPropertySetInfo):

    def __init__(self, props):
        self.props = props

    def get_index(self, name):
        for i, prop in enumerate(self.props):
            if name == prop[0]:
                return i
        return None

    def getProperties(self):
        _props = []
        for prop in self.props:
            _props.append(Property(*prop))
        return tuple(_props)

    def getPropertyByName(self, name):
        i = self.get_index(name)
        if i is None:
            raise UnknownPropertyException("Unknown property: " + name, self)
        p = self.props[i]
        return Property(*p)

    def hasPropertyByName(self, name):
        return self.get_index(name) != None


class ORToolsSolver(unohelper.Base,
                    XSolver,
                    XSolverDescription,
                    XPropertySet,
                    XServiceInfo):

    def __init__(self, ctx):
        self.ctx = ctx
        self.Document = None
        self.Objective = CellAddress(0, 0, 0)
        self.Variables = list()
        self.Constraints = list()
        self.Maximize = True
        self.Success = False
        self.ResultValue = 0
        self.Solution = list()
        self.ComponentDescription = "OR-Tools for Linear Models"
        self.tempo_total_gv = 0
        self.tempo_total_sv = 0
        self.sv_count = 0
        self.gv_count = 0
        self.ORTOOLS_IMPORT_OK = True
        self.StatusDescription = ""
        self.ortools_path = ""
        self.ortools_engine = ""
        # Set-up engine
        self.setup_ortools()
        # Engine properties
        self.NonNegative = True
        self.Integer = False
        self.Timeout = 100
        self.RelativeGap = 0.01 # 1%
        # Used for XPropertySetInfo
        self.ortools_prop_info = (("NonNegative", -1, uno_bool_type, 0),
                                  ("Integer", -1, uno_bool_type, 0),
                                  ("Timeout", -1, uno_long_type, 0),
                                  ("RelativeGap", -1, uno_double_type, 0))

    # Read regisry configuration and attempt to import ortools
    def setup_ortools(self):
        ctx = uno.getComponentContext()
        smgr = ctx.getServiceManager()
        cp = smgr.createInstance("com.sun.star.configuration.ConfigurationProvider")
        node = PropertyValue("nodepath", 0, MAIN_NODE, 0)
        self.config_access = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
        self.ortools_engine = self.config_access.CurrentEngine
        self.ortools_path = self.config_access.Path
        # Check if import works with the provided path
        try:
            if self.ortools_path != "":
                sys.path.append(self.ortools_path)
            from ortools.linear_solver import pywraplp
        except:
            self.ORTOOLS_IMPORT_OK = False
            self.StatusDescription = "Error importing OR-Tools module"


    # Return the value at a given cell
    # Argument is a CellAddress struct
    def get_value(self, cell_address):
        xSheets = self.Document.getSheets()
        xSheet = xSheets.getByIndex(cell_address.Sheet)
        xCell = xSheet.getCellByPosition(cell_address.Column, cell_address.Row)
        fValue = xCell.getData()[0][0]
        return fValue
        # return self.Document.Sheets[cell_address.Sheet][cell_address.Row, cell_address.Column].getValue()

    # Return the type of the cell content
    def get_type(self, cell_address):
        return self.Document.Sheets[cell_address.Sheet][cell_address.Row, cell_address.Column].Type

    # Argument is a CellAddress struct; and value is numeric
    def set_value(self, cell_address, value):
        xSheets = self.Document.getSheets()
        xSheet = xSheets.getByIndex(cell_address.Sheet)
        xCell = xSheet.getCellByPosition(cell_address.Column, cell_address.Row)
        # Use setData instead of setValue because it is faster
        xCell.setData([[value]])

    # Returns a tuple with the cell address
    # This is needed because each CellAddress is different, even when pointing to the same cell
    def get_tuple(self, cell_address):
        return (cell_address.Sheet, cell_address.Row, cell_address.Column)

    # XSolver
    def setDocument(self, aDoc):
        self.Document = aDoc

    def getDocument(self):
        return self.Document

    def setObjective(self, objCell):
        self.Objective = objCell

    def getObjective(self):
        return objCell

    def setVariables(self, arrVariables):
        self.Variables = arrVariables

    def getVariables(self):
        return self.Variables

    def setConstraints(self, arrConstraints):
        self.Constraints = arrConstraints

    def getConstraints(self):
        return self.Constraints

    def setMaximize(self, bMaximize):
        self.Maximize = bMaximize

    def getMaximize(self):
        return self.Maximize

    def getSuccess(self):
        return self.Success

    def getResultValue(self):
        return self.ResultValue

    def getSolution(self):
        return self.Solution

    def solve(self):
        print("\n----------------------------")
        print("Initializing OR-Tools LibreOffice integration\n")
        # If an error occurred, update status and return
        if not self.ORTOOLS_IMPORT_OK:
            self.StatusDescription = "Error importing OR-Tools module"
            self.Success = False
            self.ResultValue = 0
            return
        else:
            from ortools.linear_solver import pywraplp

        # Lock updating the UI
        self.Document.addActionLock()
        self.Document.lockControllers()

        print("Selected engine:", self.ortools_engine)
        t_ini = time.time()
        print("Inferring linear model from sheet... ", end='')

        # Set all variable cells to zero
        for cell in self.Variables:
            self.set_value(cell, 0)

        # Dictionary containing the variable types
        # Innitially they are all floats or integers (this may change later while processing constraints)
        default_var_type = "float"
        if self.Integer:
            default_var_type = "integer"
        dic_var_types = dict()
        # This list has all the tuples in the same order as in self.Variables
        list_var_tuples = list()
        for cell in self.Variables:
            cell_tuple = self.get_tuple(cell)
            dic_var_types[cell_tuple] = default_var_type
            list_var_tuples.append(cell_tuple)

        # Extract the coefficients of the objective function
        obj_coefficients = list()
        for cell in self.Variables:
            # Get current value of the objective function
            obj_0 = self.get_value(self.Objective)
            # Increase the value of the cell to 1 and check the new objective value
            self.set_value(cell, 1)
            obj_1 = self.get_value(self.Objective)
            obj_coefficients.append(obj_1 - obj_0)
            # Restore cell value to zero
            self.set_value(cell, 0)

        # Extract the coefficients of all constraints / variable types
        constr_coefficients = list()
        for constraint in self.Constraints:
            # Constraints of type "boolean" or "integer" don't generate coefficients
            # They rather update the variable type
            c_type = constraint.Operator
            cell_tuple = self.get_tuple(constraint.Left)
            if c_type == CONSTR_BINARY:
                if cell_tuple in dic_var_types:
                    dic_var_types[cell_tuple] = "binary"
                    continue
            if c_type == CONSTR_INTEGER:
                if cell_tuple in dic_var_types:
                    dic_var_types[cell_tuple] = "integer"
                    continue
            # If it reaches here, then we need to extract coefficients
            coeff_line = list()
            # The left side must be a CellAddress
            left_0 = self.get_value(constraint.Left)
            # The right side may be a CellAddress or a double value
            right_0 = 0
            if isinstance(constraint.Right, CellAddress):
                right_0 = self.get_value(constraint.Right)
            # Check if right_1 is different than right_0 at least once
            # If so, then the right value is a constant, even when it is a CellAddress
            is_right_constant = True
            for cell in self.Variables:
                self.set_value(cell, 1)
                left_1 = self.get_value(constraint.Left)
                right_1 = 0
                if isinstance(constraint.Right, CellAddress):
                    right_1 = self.get_value(constraint.Right)
                # Calculate the coefficient
                if right_1 != right_0:
                    is_right_constant = False
                coeff = (left_1 - left_0) - (right_1 - right_0)
                coeff_line.append(coeff)
                # Restore original value
                self.set_value(cell, 0)
            # The last coefficient is the upper limit to the constraint
            if is_right_constant:
                if isinstance(constraint.Right, CellAddress):
                    coeff_line.append(self.get_value(constraint.Right))
                elif isinstance(constraint.Right, float):
                    coeff_line.append(constraint.Right)
                else:
                    self.StatusDescription = "Error: unknown constraint type"
                    self.Success = False
                    self.ResultValue = 0
                    return
            else:
                # If right side is a variable value, then both sides must be equal and the rhs is zero
                coeff_line.append(0)
            # Add the coefficients of the constraint
            constr_coefficients.append(coeff_line)

        # Create the solver object (possible values are GLOP, CLP, CBC, GLPK, SCIP)
        t_end = time.time()
        print(f"Done ({t_end - t_ini} seconds)")
        t_ini = time.time()
        print("Setting up solver object... ", end='')
        # Create the solver using the selected engine (as defined in the settings dialog)
        solver = pywraplp.Solver.CreateSolver(self.ortools_engine)

        # Check if the solver engine could be instantiated
        if solver is None:
            self.StatusDescription = "Error: unable to instantiate the solver engine"
            self.Success = False
            self.ResultValue = 0
            return

        # Lower and upper bounds for variables
        min_value = -solver.infinity()
        if self.NonNegative:
            min_value = 0
        max_value = solver.infinity()

        # Create the solver variables
        model_vars = dict()
        for cell in self.Variables:
            cell_tuple = self.get_tuple(cell)
            var_type = dic_var_types[cell_tuple]
            if var_type == "float":
                model_vars[cell_tuple] = solver.NumVar(min_value, max_value, str(cell_tuple))
            elif var_type == "integer":
                model_vars[cell_tuple] = solver.IntVar(min_value, max_value, str(cell_tuple))
            elif var_type == "binary":
                model_vars[cell_tuple] = solver.IntVar(0, 1, str(cell_tuple))
            else:
                self.StatusDescription = "Error: undefined variable type"
                self.Success = False
                self.ResultValue = 0
                return

        # Set coefficients of the objective function
        obj_function = solver.Objective()
        for i in range(len(self.Variables)):
            cell_tuple = list_var_tuples[i]
            obj_function.SetCoefficient(model_vars[cell_tuple], obj_coefficients[i])

        # Set objective type
        if self.Maximize:
            obj_function.SetMaximization()
        else:
            obj_function.SetMinimization()

        # Set coefficients of all constraints
        constr_idx = 0
        n_vars = len(self.Variables)
        for constraint in self.Constraints:
            c_type = constraint.Operator
            # Only add constraint if it has left and right sides
            if c_type != CONSTR_BINARY and c_type != CONSTR_INTEGER:
                # The RHS (right hand side) of the constraint is the last coefficient
                new_constr = None
                if c_type == CONSTR_EQUAL:
                    new_constr = solver.RowConstraint(constr_coefficients[constr_idx][n_vars], constr_coefficients[constr_idx][n_vars], "")
                elif c_type == CONSTR_LESS_EQUAL:
                    new_constr = solver.RowConstraint(0, constr_coefficients[constr_idx][n_vars], "")
                elif c_type == CONSTR_GREATER_EQUAL:
                    new_constr = solver.RowConstraint(constr_coefficients[constr_idx][n_vars], solver.infinity(), "")
                else:
                    self.StatusDescription = "Error: unknown constraint type"
                    self.Success = False
                    self.ResultValue = 0
                    return
                # Now set the coefficients of the new constraint
                for j in range(n_vars):
                    cell_tuple = list_var_tuples[j]
                    new_constr.SetCoefficient(model_vars[cell_tuple], constr_coefficients[constr_idx][j])
                constr_idx += 1

        # Set solver parameters
        solverParams = pywraplp.MPSolverParameters()
        solverParams.SetDoubleParam(solverParams.RELATIVE_MIP_GAP, self.RelativeGap)
        solver.SetTimeLimit(self.Timeout * 1000)

        # Finished setting up solver object
        t_end = time.time()
        print(f"Done ({t_end - t_ini} seconds)")

        # Call the solver
        t_ini = time.time()
        print("Running the solver... ", end='')
        status = solver.Solve()
        t_end = time.time()
        print(f"Done ({t_end - t_ini} seconds)")
        print("----------------------------\n")

        # Records the success status
        if status == pywraplp.Solver.OPTIMAL:
            self.Success = True
            self.ResultValue = solver.Objective().Value()
            self.StatusDescription = "Optimal solution found"
        elif status == pywraplp.Solver.FEASIBLE:
            self.Success = True
            self.ResultValue = solver.Objective().Value()
            self.StatusDescription = "Sub-optimal feasible solution found"
        else:
            self.Success = False
            self.ResultValue = 0
            self.StatusDescription = "No solution found"

        # Records the solution
        solution = list()
        for cell_tuple in list_var_tuples:
            var_value = model_vars[cell_tuple].solution_value()
            solution.append(var_value)
        self.Solution = solution

        # Resume updating the UI
        self.Document.unlockControllers()
        self.Document.removeActionLock()


    # XSolverDescription
    def getComponentDescription(self):
        return self.ComponentDescription

    def getPropertyDescription(self, aPropertyName):
        if aPropertyName in ortools_properties:
            return ortools_properties[aPropertyName]
        else:
            return aPropertyName + " is not a supported property"

    def getStatusDescription(self):
        return self.StatusDescription

    # XPropertySet
    def getPropertySetInfo(self):
        return PropertySetInfo(self.ortools_prop_info)

    def setPropertyValue(self, aPropName, aPropValue):
        # Only change the property value if the type is correct; otherwise leave unchanged
        if aPropName == "NonNegative":
            if isinstance(aPropValue, bool):
                self.NonNegative = aPropValue
        elif aPropName == "Integer":
            if isinstance(aPropValue, bool):
                self.Integer = aPropValue
        elif aPropName == "Timeout":
            if isinstance(aPropValue, int):
                self.Timeout = aPropValue
        elif aPropName == "RelativeGap":
            if isinstance(aPropValue, float):
                if aPropValue > 0 and aPropValue < 1:
                    self.RelativeGap = aPropValue
        else:
            raise UnknownPropertyException("Unknown property: " + aPropName, self)

    def getPropertyValue(self, aPropName):
        if aPropName == "NonNegative":
            return self.NonNegative
        elif aPropName == "Integer":
            return self.Integer
        elif aPropName == "Timeout":
            return self.Timeout
        elif aPropName == "RelativeGap":
            return self.RelativeGap
        raise UnknownPropertyException("Unknown property: " + aPropName, self)

    # Leave these listeners blank for now; need to check how to implement them
    def addPropertyChangeListener(self, aPropName, aListener):
        return

    def removePropertyChangeListener(self, aPropName, aListener):
        return

    def addVetoableChangeListener(self, aPropName, aListener):
        return

    def removeVetoableChangeListener(self, aPropName, aListener):
        return

    # XServiceInfo
    def getImplementationName(self):
        return implementation_service

    def supportsService(self, svcName):
        return g_ImplementationHelper.supportsService(implementation_name, svcName)

    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(implementation_service)

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(ORToolsSolver,
                                         implementation_name,
                                         ("com.sun.star.sheet.Solver",
                                          "com.sun.star.beans.PropertySet"))
