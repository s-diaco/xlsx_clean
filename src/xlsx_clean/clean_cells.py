from datetime import datetime
import pathlib
import shutil

import pandas
from beaupy import prompt, select

# from openpyxl import load_workbook
from rich.console import Console
import win32com.client as win32

NEW_VALUE = ""
# import questionary
# load questions from text file
with open("strings.txt") as f:
    content = f.readlines()
# remove whitespace characters like `\n` at the end of each line
content = [x.strip() for x in content]
q1 = content[0]
q2 = content[1]
q3 = content[2]

path_df = pandas.read_csv("file_data.csv")
path_df["set_dir"] = [pathlib.Path(path).parent.parent for path in path_df["dir"]]
path_list = [pathlib.Path(path) for path in path_df["dir"]]
parent_path = list(
    path_df.drop_duplicates(subset=["set_name"], keep="first")["set_name"]
)

# get parent path index
console = Console()
console.print(q1)
selected_set = select(parent_path, cursor_style="cyan")

console.print(q2)
selected_dir = select(
    [
        pathlib.Path(path).stem
        for path in list(path_df[path_df["set_name"] == selected_set]["dir"])
    ]
)
console.print(f"Selected: {selected_dir}")
path_ = pathlib.Path(
    path_df[
        (path_df["set_name"] == selected_set)
        & (path_df["dir"].str.endswith(selected_dir))
    ]["dir"].iloc[0]
)
pattern = path_df[
    (path_df["set_name"] == selected_set) & (path_df["dir"].str.endswith(selected_dir))
]["pattern"].iloc[0]
search_pattern = pattern.replace("[SERIAL]", "*")
names = path_.glob(search_pattern)
files = [str(x) for x in names if x.is_file()]

batch_serial = 0
# Prompt for serial
batch_serial = prompt(q3, target_type=str)


def find_last_workbook(files):
    # Sort the list of files by their names
    sorted_files = sorted(files)

    # The last file will be the one at the end of the sorted list
    last_file = sorted_files[-1]

    return last_file


def get_workbook_names(files, batch_serial):
    last_workbook = find_last_workbook(files)
    path_of_the_year = pathlib.Path(str(path_).replace("2024", str(datetime.now().year)))
    return last_workbook, path_of_the_year / pattern.replace(
        "[SERIAL]", batch_serial.split("/")[0]
    )


ref_workbook_name, new_workbook_name = get_workbook_names(files, batch_serial)
# Selected Excel file
console.print(f"Selected: {pathlib.Path(ref_workbook_name).stem}")
# workbook = load_workbook(ref_workbook_name)
temp_workbook = str(pathlib.Path.cwd() / "temp_workbook.xlsx")
shutil.copyfile(ref_workbook_name, temp_workbook)

# in case of this error
# AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'
# see https://stackoverflow.com/questions/52889704
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = True

# Need to load the addins before opening the workbook
workbook = excel.Workbooks.Open(temp_workbook)
addin_path_1 = content[3]
excel.Workbooks.Open(addin_path_1)
addin_path_2 = content[4]
excel.Workbooks.Open(addin_path_2)
cells_to_clear = path_df[
    (path_df["set_name"] == selected_set) & (path_df["dir"].str.endswith(selected_dir))
]["cells_to_clear"].iloc[0]
for workbook_data in cells_to_clear.split(","):
    worksheet_data = workbook_data.split("!")
    # worksheet = workbook[worksheet_data[0].replace("'", "")]
    worksheet = workbook.Worksheets(worksheet_data[0].replace("'", ""))
    is_range = len(worksheet_data[1].split(":")) > 1
    cell_range = worksheet.Range(worksheet_data[1]).Value = NEW_VALUE

serial_cell = path_df[
    (path_df["set_name"] == selected_set) & (path_df["dir"].str.endswith(selected_dir))
]["serial_cell"].iloc[0]
for workbook_data in serial_cell.split(","):
    worksheet_data = workbook_data.split("!")
    worksheet = workbook.Worksheets(worksheet_data[0].replace("'", ""))
    cell_range = worksheet.Range(worksheet_data[1])
    cell_range.Value = batch_serial
if not new_workbook_name.is_file():
    pathlib.Path(new_workbook_name.parents[0]).mkdir(parents=True, exist_ok=True)
    workbook.SaveAs(str(new_workbook_name))
    excel.WindowState = win32.constants.xlMaximized
    # excel.Application.Quit()
    # if os.name == "nt":
    # os.system(f"start excel.exe \"{new_workbook_name}\"")
