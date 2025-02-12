import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os
import sys
import re
import json
import xml.etree.ElementTree as ET


# Mapping Tableau mark classes to Power BI visual types
MARK_CLASS_TO_VISUAL_TYPE = {
    "Line": "lineChart",
    "Bar": "barChart",
    "Area": "areaChart",
    "Scatter": "scatterChart",
    "Pie": "pieChart",
    "Multipolygon": "filledMap",
    "Text": "tableEx",
    "Automatic": "columnChart",
    "columnChart": "columnChart"
}

# Function to translate Tableau formulas to DAX

def translate_to_dax(tableau_formula):
    """
    Converts Tableau calculations to Power BI DAX.

    Args:
        tableau_formula (str): The Tableau calculation string.

    Returns:
        str: Converted DAX formula or a note if no equivalent exists.
    """

    def handle_lod_expressions(formula):
        """
        Recursively resolve all LOD expressions (FIXED, INCLUDE, EXCLUDE) in the formula.
        """
        def build_calculate(expression, dimension, context_type="ALLEXCEPT"):
            if context_type == "ALLEXCEPT":
                return f"CALCULATE({expression}, ALLEXCEPT('YourTable', 'YourTable'[{dimension}]))"
            elif context_type == "ALL":
                return f"CALCULATE({expression}, ALL('YourTable'[{dimension}]))"
            elif context_type == "REMOVEFILTERS":
                return f"CALCULATE({expression}, REMOVEFILTERS('YourTable'[{dimension}]))"
            return expression

        while True:
            # Handle FIXED LOD
            fixed_match = re.search(r"\{\s*FIXED\s*\[(.+?)\]:\s*([^\}]+)\}", formula)
            if fixed_match:
                dimension, expression = fixed_match.groups()
                dax_replacement = build_calculate(expression, dimension, "ALLEXCEPT")
                formula = formula.replace(fixed_match.group(0), dax_replacement)

            # Handle INCLUDE LOD
            include_match = re.search(r"\{\s*INCLUDE\s*\[(.+?)\]:\s*([^\}]+)\}", formula)
            if include_match:
                dimension, expression = include_match.groups()
                dax_replacement = build_calculate(expression, dimension, "ALL")
                formula = formula.replace(include_match.group(0), dax_replacement)

            # Handle EXCLUDE LOD
            exclude_match = re.search(r"\{\s*EXCLUDE\s*\[(.+?)\]:\s*([^\}]+)\}", formula)
            if exclude_match:
                dimension, expression = exclude_match.groups()
                dax_replacement = build_calculate(expression, dimension, "REMOVEFILTERS")
                formula = formula.replace(exclude_match.group(0), dax_replacement)

            if not (fixed_match or include_match or exclude_match):
                break

        return formula

    # Tableau to DAX function map
    function_map = {
    # Aggregations
    r"SUM\((.+?)\)": r"SUM('YourTable'[\1])",
    r"AVG\((.+?)\)": r"AVERAGE('YourTable'[\1])",
    r"COUNT\((.+?)\)": r"COUNT('YourTable'[\1])",
    r"COUNTD\(IF\s+\[([^\]]+)\]\s+AND\s+\[([^\]]+)\]\s+THEN\s+\[([^\]]+)\]\s+END\)": (
    lambda match: f"CALCULATE(DISTINCTCOUNT('YourTable'[{match.group(3)}]), 'YourTable'[{match.group(1)}] = TRUE, 'YourTable'[{match.group(2)}] = TRUE)"),
    r"COUNTD\(IF\s+\[([^\]]+)\]\s+THEN\s+\[([^\]]+)\]\s+END\)": (
    lambda match: f"CALCULATE(DISTINCTCOUNT('YourTable'[{match.group(2)}]), 'YourTable'[{match.group(1)}])"),
    r"COUNTD\((.+?)\)": r"DISTINCTCOUNT('YourTable'[\1])",
    r"MIN\((.+?)\)": r"MIN('YourTable'[\1])",
    r"MAX\((.+?)\)": r"MAX('YourTable'[\1])",
    r"(SUM|COUNT|MAX|MIN|AVERAGE)\((.+?)\s*WITH\s*FILTER\s*\((.+?)\)\)": r"CALCULATE(\1(\2), \3)",

    # Division Handling (Dynamic DIVIDE function)
    r"(.+?)\s*/\s*(.+)": (
    lambda match: f"DIVIDE({match.group(1).strip()}, {match.group(2).strip()}, 0)"
        .replace("[[", "[")  # Fix double brackets
        .replace("]]", "]")),
    r"(.+?)\s*-\s*(.+)": (
    lambda match: f"({match.group(1).strip()} - {match.group(2).strip()})"),

    # CASE Logic
   r"CASE\s+\[([^\]]+)\](.*?)\s+END": (
    lambda match: "SWITCH(TRUE(), " +
    match.group(2)
    .replace("WHEN", ",")
    .replace("THEN", ",")
    .replace("ELSE", ",")
    + " BLANK())"),

    r"ZN\((.+?)\)": (
    lambda match: f"IF(ISBLANK({match.group(1).strip()}), 0, {match.group(1).strip()})"), # ZN functions
    # RANK Logic
    r"RANK\((.+?),\s*'(.+?)'\)": (
        lambda match: f"RANKX(ALL('YourTable'), {match.group(1).strip()}, {match.group(2).upper()})"
    ),

    # Logical Operators
    r"\bOR\b": r"||",
    r"\bAND\b": r"&&",

    # String Functions
    r"LEFT\((.+?),\s*(.+?)\)": r"LEFT('YourTable'[\1], \2)",
    r"RIGHT\((.+?),\s*(.+?)\)": r"RIGHT('YourTable'[\1], \2)",
    r"MID\((.+?),\s*(.+?),\s*(.+?)\)": r"MID('YourTable'[\1], \2, \3)",
    r"STR\((.+?)\)": r"FORMAT(\1, \"General Number\")",
    r"LEN\((.+?)\)": r"LEN('YourTable'[\1])",

    # Date Functions
    r"DATEADD\((.+?),\s*(.+?),\s*(.+?)\)": r"DATEADD('YourTable'[\1], \2, \3)",
    r"DATEDIFF\((.+?),\s*(.+?),\s*(.+?)\)": r"DATEDIFF('YourTable'[\1], 'YourTable'[\2], \3)",
    r"YEAR\((.+?)\)": r"YEAR('YourTable'[\1])",
    r"MONTH\((.+?)\)": r"MONTH('YourTable'[\1])",
    r"DAY\((.+?)\)": r"DAY('YourTable'[\1])",
    r"DATETRUNC\((.+?),\s*(.+?)\)": r"TRUNC('YourTable'[\2], \1)",

    # NULL/ISNULL Checks
    r"ISNULL\((.+?)\)": r"ISBLANK('YourTable'[\1])",
    

    # Conditional Statements
    r"IIF\s*\(\s*(.+?)\s*,\s*'(.+?)'\s*,\s*'(.+?)'\s*\)": r"IF(\1, \2, \3)",
    r"IF\s+(.+?)\s+THEN\s+(.+?)\s+ELSE\s+(.+?)\s+END": (
        lambda match: (
            f"IF({match.group(1).strip()}, {match.group(2).strip()}, {match.group(3).strip()})"
            if "IF" not in match.group(2) and "IF" not in match.group(3)
            else f"SWITCH(TRUE(), {match.group(1).strip()}, {match.group(2).strip()}, {match.group(3).strip()}, BLANK())"
        )
    ),

   

    # Arithmetic
    r"\[([^\]]+)\]\s*\+\s*\[([^\]]+)\]": r"'YourTable'[\1] + 'YourTable'[\2]",
    r"\[([^\]]+)\]\s*\-\s*\[([^\]]+)\]": r"'YourTable'[\1] - 'YourTable'[\2]",
    r"\[([^\]]+)\]\s*\*\s*\[([^\]]+)\]": r"'YourTable'[\1] * 'YourTable'[\2]",
    r"\[([^\]]+)\]\s*/\s*\[([^\]]+)\]": r"DIVIDE('YourTable'[\1], 'YourTable'[\2])",
}


    # Handle LOD Expressions
    dax_formula = handle_lod_expressions(tableau_formula)

    # Replace Tableau functions with DAX equivalents
    for tableau_pattern, dax_replacement in function_map.items():
        dax_formula = re.sub(tableau_pattern, dax_replacement, dax_formula, flags=re.IGNORECASE)
        
    # Determine whether it's a measure or calculated column
    if any(agg in tableau_formula.upper() for agg in ["SUM", "AVG", "COUNT", "COUNTD", "MIN", "MAX", "FIXED"]):
        classification = "measure"
    else:
        classification = "calculated_column"

    return {
        "dax_formula": dax_formula.strip(),
        "type": classification
    }

   
# Function to extract calculated fields
def extract_calculated_fields(root):
    calculated_fields = []
    for column in root.findall(".//column"):
        calc = column.find("calculation")
        if calc is not None:
            formula = calc.get("formula")
            if formula:
                cleaned_tableau_formula = formula.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()
                dax_result = translate_to_dax(cleaned_tableau_formula)
                cleaned_dax_formula = dax_result["dax_formula"].replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()
                calculated_fields.append({
                    "name": column.get("caption", "Unnamed Calculation"),
                    "tableau_formula": cleaned_tableau_formula,
                    "dax_formula": cleaned_dax_formula,
                    "type": dax_result["type"]
                })
    return calculated_fields


# Function to extract dashboards and visuals

def extract_dashboards_and_visuals_from_xml(xml_data):
    dashboards_data = {}
    root = ET.fromstring(xml_data)

    extract_dashboards(dashboards_data, root)
    extract_worksheets(dashboards_data, root)

    return dashboards_data, root

def extract_dashboards(dashboards_data, root):
    for dashboard in root.findall(".//dashboard"):
        dashboard_name = dashboard.get("name", "Unnamed Dashboard")
        dashboards_data[dashboard_name] = {}

        for zone in dashboard.findall(".//zone"):
            worksheet_name = zone.get("name")
            if worksheet_name:
                dashboards_data[dashboard_name][worksheet_name] = []

def extract_worksheets(dashboards_data, root):
    for worksheet in root.findall(".//worksheet"):
        worksheet_name = worksheet.get("name", "Unnamed Worksheet")
        visuals = extract_visuals(worksheet)

        add_visuals_to_dashboards(dashboards_data, worksheet_name, visuals)

def extract_visuals(worksheet):
    visuals = []
    for pane in worksheet.findall(".//pane"):
        for mark in pane.findall(".//mark"):
            mark_class = mark.get("class")
            if mark_class and mark_class in MARK_CLASS_TO_VISUAL_TYPE:
                visuals.append(MARK_CLASS_TO_VISUAL_TYPE[mark_class])
    return visuals

def add_visuals_to_dashboards(dashboards_data, worksheet_name, visuals):
    for dashboard_name, worksheets in dashboards_data.items():
        if worksheet_name in worksheets:
            dashboards_data[dashboard_name][worksheet_name].extend(visuals)


# Function to create Power BI JSON
def create_power_bi_json(dashboards_data):
    power_bi_report = initialize_power_bi_report()
    visual_width, visual_height = 300, 200
    page_width, margin = 1280, 20

    for ordinal, (dashboard_name, worksheets) in enumerate(dashboards_data.items(), start=1):
        section = create_section(dashboard_name, ordinal, page_width)
        x_offset, y_offset, z_offset = margin, margin, 0
        processed_visuals = set()

        for worksheet_name, visuals in worksheets.items():
            x_offset, y_offset, z_offset = process_visuals(visuals, worksheet_name, section, processed_visuals, x_offset, y_offset, z_offset, visual_width, visual_height, page_width, margin)

        power_bi_report["sections"].append(section)

    return power_bi_report

def initialize_power_bi_report():
    return {
        "config": json.dumps({
            "version": "5.59",
            "themeCollection": {
                "baseTheme": {"name": "CY24SU10", "version": "5.59", "type": 2}
            },
            "activeSectionIndex": 0,
            "settings": {
                "useNewFilterPaneExperience": True,
                "allowChangeFilterTypes": True
            }
        }),
        "layoutOptimization": 0,
        "resourcePackages": [
            {
                "resourcePackage": {
                    "disabled": False,
                    "items": [
                        {
                            "name": "CY24SU10",
                            "path": "BaseThemes/CY24SU10.json",
                            "type": 202
                        }
                    ],
                    "name": "SharedResources",
                    "type": 2
                }
            }
        ],
        "sections": []
    }

def create_section(dashboard_name, ordinal, page_width):
    return {
        "config": "{}",
        "displayName": dashboard_name,
        "displayOption": 1,
        "filters": "[]",
        "height": 720.0,
        "name": dashboard_name.replace(" ", "_"),
        "ordinal": ordinal,
        "visualContainers": [],
        "width": page_width
    }

def process_visuals(visuals, worksheet_name, section, processed_visuals, x_offset, y_offset, z_offset, visual_width, visual_height, page_width, margin):
    for visual_type in visuals:
        if (worksheet_name, visual_type) not in processed_visuals:
            add_visual_container(section, worksheet_name, visual_type, x_offset, y_offset, z_offset, visual_width, visual_height)
            processed_visuals.add((worksheet_name, visual_type))
            x_offset, y_offset, z_offset = update_offsets(x_offset, y_offset, z_offset, visual_width, visual_height, page_width, margin, section)
    return x_offset, y_offset, z_offset

def add_visual_container(section, worksheet_name, visual_type, x_offset, y_offset, z_offset, visual_width, visual_height):
    section["visualContainers"].append({
        "config": json.dumps({
            "name": "",
            "layouts": [{
                "id": 0,
                "position": {
                    "x": x_offset,
                    "y": y_offset,
                    "z": z_offset,
                    "width": visual_width,
                    "height": visual_height
                }
            }],
            "singleVisual": {
                "visualType": visual_type,
                "drillFilterOtherVisuals": True
            }
        }),
        "filters": "[]",
        "height": visual_height,
        "width": visual_width,
        "x": x_offset,
        "y": y_offset,
        "z": z_offset
    })

def update_offsets(x_offset, y_offset, z_offset, visual_width, visual_height, page_width, margin, section):
    x_offset += visual_width + margin
    z_offset += 1000

    if x_offset + visual_width > page_width:
        x_offset = margin
        y_offset += visual_height + margin

    if y_offset + visual_height > section["height"]:
        section["height"] += visual_height + margin

    return x_offset, y_offset, z_offset

# Main function for processing
def main(twb_file_path, output_json_path, dax_output_path):
    """
    Main function to process Tableau XML, generate Power BI-compatible JSON, and save DAX formulas separately.
    """
    try:
        # Extract directory and file name from the TWB path
        twb_directory = os.path.dirname(twb_file_path)
        twb_name = os.path.splitext(os.path.basename(twb_file_path))[0]

        # Paths for Power BI-related files
      
        report_json_path = os.path.join(twb_directory, f"{twb_name}.Report/report.json")
        report_folder = os.path.dirname(report_json_path)
        definition_pbir_path = os.path.join(report_folder, "definition.pbir")

        # Read Tableau XML
        with open(twb_file_path, 'r') as xml_file:
            xml_data = xml_file.read()

        # Generate Power BI JSON and DAX Calculations JSON
        dashboards_data, root = extract_dashboards_and_visuals_from_xml(xml_data)
        power_bi_json = create_power_bi_json(dashboards_data)
        calculated_fields = extract_calculated_fields(root)

        # Save the generated JSON files
        with open(output_json_path, "w") as file:
            json.dump(power_bi_json, file, indent=4)

        with open(dax_output_path, "w") as file:
            json.dump(calculated_fields, file, indent=4)

        print(f"Power BI JSON saved to {output_json_path}")
        print(f"DAX Calculations JSON saved to {dax_output_path}")

        # Update report.json with the Power BI-compatible JSON data
        if os.path.exists(report_json_path):
            print(f"Found report.json at {report_json_path}, updating it...")
            with open(output_json_path, "r") as generated_json_file:
                generated_data = json.load(generated_json_file)

            with open(report_json_path, "w") as report_file:
                json.dump(generated_data, report_file, indent=4)

            print(f"Updated report.json successfully: {report_json_path}")

            if os.path.exists(definition_pbir_path):
                print(f"Found definition.pbir at {definition_pbir_path}, opening it...")
                os.startfile(definition_pbir_path)
            else:
                print(f"'definition.pbir' not found at {definition_pbir_path}")
        else:
            print(f"report.json not found at {report_json_path}")

    except Exception as e:
        print(f"An error occurred: {e}")


# Function to run the script through GUI
def run_script():
    """
    Executes the main logic through GUI input and closes the GUI after completion.
    """
    file_path = file_entry.get()
    if not file_path or not os.path.isfile(file_path):
        messagebox.showerror("Error", "Please select a valid .twb file!")
        return

    output_json_path = os.path.splitext(file_path)[0] + "_power_bi_report.json"
    dax_output_path = os.path.splitext(file_path)[0] + "_dax_calculations.json"

    try:
        main(file_path, output_json_path, dax_output_path)  # Call main function here
        messagebox.showinfo("Success", f"Conversion completed!\n\nOutputs:\n- {output_json_path}\n- {dax_output_path}")
        root.destroy()  # Close the GUI after success
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        root.destroy()  # Ensure GUI closes even on error


# GUI Setup
root = tk.Tk()
root.title("VM BIConvert")
root.geometry("650x400")  # Increased size for better layout
root.resizable(False, False)
root.configure(bg="#f2f2f2")  # Light gray background

def resource_path(relative_path):
    """Get absolute path to resource, works for PyInstaller"""
    try:
        # PyInstaller stores temp files in _MEIPASS when bundled
        base_path = sys._MEIPASS
    except AttributeError:
        # Normal development environment
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Logo
logo_path = resource_path("VM logo2.png")  # Use resource_path to get the correct path
try:
    logo_image = Image.open(logo_path)
    logo = ImageTk.PhotoImage(logo_image)
    logo_label = tk.Label(root, image=logo, bg="#f2f2f2")
    logo_label.place(x=550, y=10)  # Top-right placement
except Exception as e:
    print(f"Error loading logo: {e}")
# Title
segoe_ui = "Segoe UI"
title_label = tk.Label(
    root,
    text="VM BIConvert",
    font=(segoe_ui, 16, "bold"),
    bg="#f2f2f2",
    fg="#333333"
)
title_label.place(x=325, y=20, anchor="center")  # Centered title


# Dropdown for tool selection
dropdown_label = tk.Label(
    root,
    text="Tool:",
    font=(segoe_ui, 12),
    bg="#f2f2f2",
    fg="#333333"
)
dropdown_label.place(x=40, y=100)

tool_var = tk.StringVar(value="Select Tool")  # Default value

tool_dropdown = tk.OptionMenu(root, tool_var, "Tableau", "Cognos")
tool_dropdown.config(
    font=(segoe_ui, 12),
    bg="white",
    fg="#333333",
    width=25  # Sets the width of the main dropdown button
)

# Adjust dropdown menu options to match the button width
menu = tool_dropdown["menu"]
menu.config(font=(segoe_ui, 12), bg="white", fg="#333333")

tool_dropdown.place(x=150, y=95, width=120, height=30)

# File Path Entry Section
path_label = tk.Label(
    root,
    text="Provide the file Path:",
    font=(segoe_ui, 12),
    bg="#f2f2f2",
    fg="#333333"
)
path_label.place(x=40, y=160)

file_entry = tk.Entry(root, width=35, font=(segoe_ui, 12), borderwidth=2, relief="groove")
file_entry.place(x=150, y=160)

browse_button = tk.Button(
    root,
    text="Browse",
    command=lambda: file_entry.insert(0, filedialog.askopenfilename(filetypes=[("Tableau Workbook Files", "*.twb")])),
    font=(segoe_ui, 10, "bold"),
    bg="#007BFF",
    fg="white",
    relief="raised",
    width=10
)
browse_button.place(x=500, y=157)

# Run Button
run_button = tk.Button(
    root,
    text="Run",
    font=(segoe_ui, 10, "bold"),
    bg="#007BFF",
    fg="white",
    relief="raised",
    width=10,
    command=run_script
)
run_button.place(x=325, y=250, anchor="center")  # Centered run button

# Footer
footer_label = tk.Label(
    root,
    text="Developed by ValueMomentum",
    font=(segoe_ui, 10, "italic"),
    bg="#f2f2f2",
    fg="#666666"
)
footer_label.place(x=325, y=370, anchor="center")  # Centered footer

# Run GUI
root.mainloop()
