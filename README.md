# Utility Legend Generator

**Utility Legend Generator** is a Python-based tool designed to streamline the creation of utility key legends for highway utility drawings in AutoCAD. Built using the `win32com` module, this tool offers a user-friendly interface that simplifies and accelerates the legend generation process.

## Key Features

### 1. Data Source Selection
- **Purpose:** Allows users to select a source file that contains all the necessary layers and keys for legend creation.
- **Requirement:** The data source file must be an `.xlsx` file with column names "Keys" and "LayerNames."
- **Usage:** Selecting this checkbox enables the browse button to choose your source file. This step is mandatory before proceeding.

### 2. Target Drawing Selection
- **Purpose:** Enables the generation of legends in a specific drawing file.
- **Usage:** If this checkbox is selected, you can browse and select the target drawing file where legends will be created. If not selected, legends will be generated in the currently opened drawing.

### 3. LISP File Creation
- **Purpose:** Generates a `.lsp` file containing all commands and instructions used during the legend creation process.
- **Usage:** This is an optional feature. If selected, a LISP file with the same name as the drawing file will be created, allowing you to recreate the legends by loading the file into AutoCAD at any time.

### 4. Launch Button
- **Purpose:** Starts the legend creation process.
- **Usage:** After configuring your options, click this button to launch the tool and begin generating legends.

## Installation and Requirements

- **Python:** Ensure you have Python installed on your system.
- **win32com Module:** This tool relies on the `win32com` module. Install it using pip: ```pip install pywin32```
- **ADEPT Tool:** Ensure the ADEPT tool is already installed on your system, as it is required to use this tool.

## Usage

1. **Data Source File:** Prepare an `.xlsx` file with the necessary layers and keys. Make sure the columns are named "Keys" and "LayerNames."
2. **Launch the Tool:** Open the Utility Legend Generator tool.
3. **Configure Options:** Select your data source, target drawing, and choose whether to generate a LISP file.
4. **Generate Legends:** Click the "Launch" button to create the legends at the origin of the coordinate plane in AutoCAD’s model space.

## Additions

1. Added Dynamo scipt if user want to use it through civil 3D directly, it works much faster if we compare it to fully python version.
2. Added a demo source.xlsx file for users convenient as to understand the format of the required source file also user can add their layers and keys in the respective columns in this file only and use as the source whether they are using it through python version or Dynamo version.

## Contribution

We welcome contributions to the Utility Legend Generator tool. If you have any suggestions, improvements, or encounter any issues, feel free to reach out.

**Contact:** For help, suggestions, or contributions, please send an email to [BHUTUU](mailto:raj259942@gmail.com).

