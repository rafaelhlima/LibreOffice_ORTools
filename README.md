# LibreOffice OR-Tools Integration

This extension integrates the [OR-Tools](https://developers.google.com/optimization) optimization suite into LibreOffice.

## Features

The current implementation supports solving Mixed Integer Linear Programming (MILP) problems. After installing the extension, a new solver engine will be available in the Solver Options dialog.

The supported solver engines are: CBC, SCIP and GLOP.

## Configuration

This extensions uses the `ortools` Python package, so it must be accessible from your Python path.

If OR-Tools is not accessible to LibreOffice's Python interpreter, use the **OR-Tools Settings** dialog to set the path to an external environment that contains OR-Tools.
