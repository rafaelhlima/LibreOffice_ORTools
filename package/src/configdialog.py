###############################################
# This file implements the configuration dialog
###############################################

import uno
import unohelper
import sys
import traceback

from scriptforge import CreateScriptService
from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
from com.sun.star.awt import XActionListener, XTextListener, XItemListener
from com.sun.star.awt import FontWeight
from com.sun.star.task import XJobExecutor
from com.sun.star.beans import PropertyValue

FOLDER_ICON = "private:graphicrepository/cmd/lc_open.png"
MAIN_NODE = "ortools.Settings/EngineOptions"
LIST_ENGINES = ["CBC", "CP-SAT", "GLOP", "SCIP"]

# Listener for all buttons in the dialog
class ActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog, config_access):
        self.dialog = dialog
        self.config_access = config_access
        self.bas = CreateScriptService("Basic")
        self.fso = CreateScriptService("FileSystem")
        self.fso.FileNaming = "SYS"

    def disposing(self, source):
        pass

    # Returns True if OR-Tools import is possible
    def testORToolsImport(self):
        edit_control = self.dialog.getControl("Edit_Path")
        import_path = edit_control.getText()
        if import_path != "":
            sys.path.append(import_path)
        success = False
        try:
            from ortools.linear_solver import pywraplp
            success = True
        except:
            pass
        finally:
            # Remove the path tested
            if import_path != "":
                del sys.path[-1]
        return success

    # Returns True if a given solver engine can be created
    def testSolverEngine(self, engine_name):
        edit_control = self.dialog.getControl("Edit_Path")
        import_path = edit_control.getText()
        if import_path != "":
            sys.path.append(import_path)
        success = False
        try:
            if engine_name == "CPSAT":
                from ortools.sat.python import cp_model
                model = cp_model.CpModel()
                success = True
            else:
                from ortools.linear_solver import pywraplp
                model = pywraplp.Solver.CreateSolver(engine_name)
                success = True
        except:
            pass
        finally:
            # Remove the path tested
            if import_path != "":
                del sys.path[-1]
        return success

    def actionPerformed(self, ev):
        if ev.ActionCommand == "cancel":
            self.dialog.endExecute()
        elif ev.ActionCommand == "open":
            folder_path = self.fso.PickFolder(freetext = "Select the folder where OR-Tools is located")
            edit_control = self.dialog.getControl("Edit_Path")
            edit_control.setText(folder_path)
        elif ev.ActionCommand == "test":
            b_path = self.testORToolsImport()
            b_scip = self.testSolverEngine("SCIP")
            b_glop = self.testSolverEngine("GLOP")
            b_cbc = self.testSolverEngine("CBC")
            b_cpsat = self.testSolverEngine("CPSAT")
            self.setLabelStatus("Label_Import_Status", b_path)
            self.setLabelStatus("Label_SCIP_Status", b_scip)
            self.setLabelStatus("Label_GLOP_Status", b_glop)
            self.setLabelStatus("Label_CBC_Status", b_cbc)
            self.setLabelStatus("Label_CPSAT_Status", b_cpsat)
        elif ev.ActionCommand == "ok":
            new_path = self.dialog.getControl("Edit_Path").getText()
            new_engine = self.dialog.getControl("List_Engines").getSelectedItem()
            if self.config_access.CurrentEngine != new_engine or self.config_access.Path != new_path:
                self.config_access.CurrentEngine = new_engine
                self.config_access.Path = new_path
                self.config_access.commitChanges()
            self.dialog.endExecute()

    # Sets the label to OK (green) or Fail (red)
    def setLabelStatus(self, label_name, b_value):
        label_control = self.dialog.getControl(label_name)
        label_control.Model.FontWeight = FontWeight.BOLD
        if b_value:
            label_control.setText("OK")
            label_control.Model.TextColor = self.bas.RGB(0, 255, 0)
        else:
            label_control.setText("Fail")
            label_control.Model.TextColor = self.bas.RGB(255, 0, 0)


# Configuration dialog
class ConfigDialog(unohelper.Base):
    def __init__(self):
        self.bas = CreateScriptService("Basic")
        # Get current configuration from the registry
        ctx = uno.getComponentContext()
        smgr = ctx.getServiceManager()
        cp = smgr.createInstance("com.sun.star.configuration.ConfigurationProvider")
        node = PropertyValue("nodepath", 0, MAIN_NODE, 0)
        self.config_access = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
        self.current_engine = self.config_access.CurrentEngine
        self.current_path = self.config_access.Path

    def create_ui(self):
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.getServiceManager()
        # Create the dialog window
        self.dialog = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        self.dialog_model = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        self.dialog.setModel(self.dialog_model)
        self.dialog.setTitle("OR-Tools settings")
        self.dialog.setPosSize(0, 0, 526, 362, SIZE)
        # Default action listener used for all buttons
        self.btn_action = ActionListener(self.dialog, self.config_access)
        # Create buttons
        self.create_button("Btn_Ok", "OK", (294, 308, 106, 35), "ok")
        self.create_button("Btn_Cancel", "Cancel", (404, 308, 106, 35), "cancel")
        self.create_button("Btn_Test", "Test", (448, 40, 64, 35), "test")
        btn_open = self.create_button("Btn_Open", "", (400, 40, 44, 35), "open")
        btn_open.Model.ImageURL = FOLDER_ICON
        # Create labels
        self.create_label("Label_Path", "Path to OR-Tools", (16, 16, 200, 25))
        self.create_label("Label_Test", "Test", (16, 100, 200, 25), True)
        # label_test.Model.FontWeight = FontWeight.BOLD
        self.create_label("Label_Status", "Status", (166, 100, 200, 25), True)
        # label_status.Model.FontWeight = FontWeight.BOLD
        self.create_label("Label_Import", "OR-Tools library", (16, 130, 200, 25))
        self.create_label("Label_Import_Status", "Unknown", (166, 130, 200, 25))
        self.create_label("Label_SCIP", "SCIP Engine", (16, 160, 200, 25))
        self.create_label("Label_SCIP_Status", "Unknown", (166, 160, 200, 25))
        self.create_label("Label_GLOP", "GLOP Engine", (16, 190, 200, 25))
        self.create_label("Label_GLOP_Status", "Unknown", (166, 190, 200, 25))
        self.create_label("Label_CBC", "CBC Engine", (16, 220, 200, 25))
        self.create_label("Label_CBC_Status", "Unknown", (166, 220, 200, 25))
        self.create_label("Label_CPSAT", "CP-SAT Engine", (16, 250, 200, 25))
        self.create_label("Label_CPSAT_Status", "Unknown", (166, 250, 200, 25))
        self.create_label("Label_Engines", "Choose the default engine", (300, 100, 200, 25), True)
        # Edit
        edit_path = self.create_edit("Edit_Path", (16, 40, 380, 35))
        edit_path.Model.HelpText = "Leave blank if OR-Tools is accessible from LibreOffice's PYTHONPATH"
        edit_path.Model.Text = self.current_path
        # List Box
        list_engines = self.create_listbox("List_Engines", (300, 130, 212, 120))
        list_engines.addItems(LIST_ENGINES, 0)
        idx_engine = LIST_ENGINES.index(self.current_engine)
        list_engines.Model.SelectedItems = (idx_engine, )

    def create_button(self, btn_name, label, possize, action_cmd):
        btn_model = self.dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        self.dialog_model.insertByName(btn_name, btn_model)
        btn_control = self.dialog.getControl(btn_name)
        btn_control.setPosSize(possize[0], possize[1], possize[2], possize[3], POSSIZE)
        btn_control.setLabel(label)
        btn_control.setActionCommand(action_cmd)
        btn_control.addActionListener(self.btn_action)
        return btn_control

    def create_label(self, label_name, label, possize, bold=False):
        label_model = self.dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        self.dialog_model.insertByName(label_name, label_model)
        label_control = self.dialog.getControl(label_name)
        label_control.setPosSize(possize[0], possize[1], possize[2], possize[3], POSSIZE)
        label_control.setText(label)
        if bold:
            label_control.Model.FontWeight = FontWeight.BOLD
        return label_control

    def create_edit(self, edit_name, possize):
        edit_model = self.dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        self.dialog_model.insertByName(edit_name, edit_model)
        edit_control = self.dialog.getControl(edit_name)
        edit_control.setPosSize(possize[0], possize[1], possize[2], possize[3], POSSIZE)
        return edit_control

    def create_listbox(self, list_name, possize):
        list_model = self.dialog_model.createInstance("com.sun.star.awt.UnoControlListBoxModel")
        self.dialog_model.insertByName(list_name, list_model)
        list_control = self.dialog.getControl(list_name)
        list_control.setPosSize(possize[0], possize[1], possize[2], possize[3], POSSIZE)
        return list_control

    def run(self):
        self.create_ui()
        self.dialog.setEnable(True)
        self.dialog.setVisible(True)
        self.dialog.execute()


# Export component as a XJobExecutor instance
class ORToolsConfigDialog(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        # Initialize ScriptForge services
        self.dialog = ConfigDialog()
        self.dialog.run()


# Export implementation
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    ORToolsConfigDialog, "org.libreoffice.comp.ORToolsLinear_Impl.ConfigDialog",
    ("com.sun.star.Job", ), )
