# Revenue Operations

This repository contains Python scripts for processing Excel files related to revenue operations, specifically handling financial data for projects across different months.

## Tasks

### Task 1: Excel Manipulation - Load and Save

- **File**: `Task_1_Excel_Manipulation/Load_Save_Excel.py`
- **Purpose**: Processes monthly data from `Delta3_Apr.xlsx` and populates a template `Delta3_Output.xlsx` with revenue and cost calculations.
- **Functionality**:
  - Reads sheets named after months (e.g., "Apr", "May").
  - Extracts revenue, salaries, and other costs.
  - Calculates total costs and percentages.
  - Updates the template with the computed values.

### Task 2: Excel Manipulation - Load, Save, and Values Update

- **Files**:
  - `Task_2_Excel_Manipulation/Load_Save_Excel.py`
  - `Task_2_Excel_Manipulation/Values_Excel.py`
- **Purpose**: Similar to Task 1, but saves to `Delta3_Final_Output.xlsx`, and then updates with revenue values from MIS files for April, May, June.
- **Functionality**:
  - `Load_Save_Excel.py`: Processes `Delta3_Apr.xlsx` and saves to `Delta3_Final_Output.xlsx`.
  - `Values_Excel.py`: Extracts revenue data from `MIS_Final_April.xlsx`, `MIS_Final_May.xlsx`, `MIS_Final_June.xlsx` and updates the final output file with project-specific revenue values.
