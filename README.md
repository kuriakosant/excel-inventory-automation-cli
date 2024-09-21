# Excel Inventory Modifier (V1)

This project is a **command line interface (CLI) tool** that allows users to modify the data in an Excel file.
It was requested by a local business to automate the task of managing and modifying inventories using excel spreadsheets.

The tool takes an Excel file as input and applies modifications based on a user-defined range of values in **Column J**. Based on the user's input, it can reduce the quantity, per-unit value, or total value of products within the specified range, in all cases it recalculates the sum of Column J to account for different values in Columns F and or I. The modified Excel file is then saved as a new file with the same formatting as the original file.

## Features

- Select an Excel file placed in the same directory as the script.
- Modify rows where **Column J** (Αξία) falls within a user-defined range.
- Reduce the quantity, per-unit value, or total value of products.
- Modify values by percentage and apply rounding options.
- Automatically update calculated values in **Columns F, I, and J**.
- Save the modified data as a new Excel file with the original formatting (font, size, etc.) retained.

## Requirements

- Excel ( Microsoft or a proprietary version)
- Python 3.x
- Required Python packages:
  - `pandas`
  - `openpyxl`

## Installation

### 1. Clone the repository

To clone this repository, run:

`git clone https://github.com/kuriakosant/excel-inventory-automation-cli`

### 2. Install Dependencies

1. Navigate to the project `excel-inventory-automation-cli` directory:

`cd excel-inventory-automation-cli`

2. Create a virtual environment to isolate project dependencies( you only need to do this once ):

`python3 -m venv venv`

3. Activate the virtual environment :

- On **Linux/MacOS**:

`source venv/bin/activate`

- On **Windows**:
  `venv\Scripts\activate`

4. Install the required Python dependencies:

`pip install -r requirements.txt`

6. Deactivate the virtual environment whenever youre done:

`deactivate`

## Usage

1.  **Place the Excel File**: Ensure the Excel file you want to modify is in the same directory as the script.

2.  **Run the script**: Open a terminal and navigate to the projects directory , then run it with the following command.

    `python main.py`

3.  **Follow the command line interface (CLI) prompts** to modify the Excel file.

---

## How it Works

### Step 1: File Selection

When you run the script, it will display:

EXCEL FILE MODIFIER (V1)
Please select your next action:

1.  Choose file
2.  Exit
    Selection(1/2):

- Select `1` to choose the file.
- Enter the file name, for example: `filename.xlsx
- The program will check if the file exists and proceed.

### Step 2: Modify Values

Once the file is selected, the program will ask you to input a minimum and maximum value for **Column J (Αξία)**:

`Please input your min value (press enter for 0): 
Please input your max value:`

This selects all rows where **Column J** contains a value within this range.

### Step 3: Modifying Specific Data

The program will then prompt you to select how you want to modify the data:

Please select your next action:

1.  Reduce quantity of product
2.  Reduce per-unit value of product
3.  Reduce total value of product
    Selection(1-3):

- **Case 1 (Reduce quantity of product)**:

  - The program will reduce the quantity in **Column F** by the percentage you specify.
  - It will recalculate **Column J** as `F * I = J`.
  - It will ask how you want to round the result (nearest integer or decimal point).
  - Finally, it updates the sum of **Column F** at the bottom of the file.

- **Case 2 (Reduce per-unit value of product)**:

  - The program will reduce the value in **Column I** (per-unit value).
  - It will recalculate **Column J** as `F * I = J`.
  - The sum of **Column I** is not updated in this case.

- **Case 3 (Reduce total value of product)**:

  - The program will reduce both the quantity in **Column F** and the per-unit value in **Column I** by the specified percentage.
  - It recalculates **Column J** for each affected row.
  - The sum of **Columns F, I, and J** will be recalculated and updated at the bottom of the file.

### Rounding Behavior

When modifying values, the program offers several rounding options:

1. Round down to the nearest integer
2. Round up to the nearest integer
3. Round to 1 decimal place
4. Round to 2 decimal places

The rounding behavior has been carefully implemented to handle small values:

- For values less than 1:

  - When rounding down to the nearest integer, the value remains unchanged.
  - When rounding up to the nearest integer, the value becomes 1.
  - For decimal rounding (1 or 2 decimal places), the value is rounded as requested but never becomes zero.

- For values 1 and above:

  - Integer rounding (options 1 and 2) works as expected.
  - Decimal rounding (options 3 and 4) rounds to the specified number of decimal places.

  The special case of rounding up to 1 only applies when the value is less than 1. This is to prevent very small quantities from disappearing entirely.

In all cases, the program ensures that no value becomes zero due to rounding. This preserves small quantities and values in the inventory.

### Step 4: Save Changes

After making the modifications, the program will ask you to name the new file:

`Please select a name for your new file (just the file name, not including file type):`

It will save the new file with the specified name and retain the original formatting, including fonts and cell styles.

---

### Example Workflow

1.  **Run the program**:

    `python main.py`

2.  **Select a file**:

    - Place `file.xlsx` in the root directory and input its name.

3.  **Define a range** for **Column J**:

    - Input the min and max values, for example, `min: 100`, `max: 500`.

4.  **Modify values**:
    - Select an action, such as reducing the quantity, by 30%.
    - Choose how you want to round the results.
5.  **Save the new file**:

    - Name your new file, for example, `modified_file.xlsx`.

The modified Excel file will now contain the new values and be saved with the same formatting as the original.

---

## License

This project is licensed under the MIT license. See the [LICENSE](./LICENSE) file for details.

---
