import os
import sys
import glob
import subprocess
import csv
import time

try:
    import pandas as pd
except ImportError:
    subprocess.run(["pip", "install", "pandas"])
    import pandas as pd

try:
    import openpyxl
except ImportError:
    subprocess.run(["pip", "install", "openpyxl"])
    import openpyxl

# Clear the screen
os.system('cls' if os.name == 'nt' else 'clear')

def human_friendly_time(seconds):
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    time_str = ""
    if hours > 0:
        time_str += f"{int(hours)} hour{'s' if hours > 1 else ''}, "
    if minutes > 0:
        time_str += f"{int(minutes)} minute{'s' if minutes > 1 else ''}, "
    time_str += f"{seconds:.2f} second{'s' if seconds != 1 else ''}"
    return time_str

def validate_headers(file_list):
    header = None
    for file in file_list:
        excel_data = pd.read_excel(file, sheet_name=0, nrows=0)
        if header is None:
            header = excel_data.columns.tolist()
        elif header != excel_data.columns.tolist():
            return False
    return True

def sanitize_value(value):
    if isinstance(value, str):
        return value.replace("'", "''").replace('"', '""').replace("/", "//").replace("\\", "\\\\")
    return value

def get_xlsx_files(input_dir):
    file_list = []
    for file in os.listdir(input_dir):
        if file.lower().endswith('.xlsx'):
            file_list.append(os.path.join(input_dir, file))
    return file_list

def combine_xlsx_to_csv(input_dir, output_dir, output_filename):
    file_list = get_xlsx_files(input_dir)

    if not validate_headers(file_list):
        print("\033[91mThe excel files do not have an identical structure.\033[0m")  # Bright red text
        return

    os.makedirs(output_dir, exist_ok=True)
    if not output_filename.lower().endswith('.csv'):
        output_filename += '.csv'
    output_file = os.path.join(output_dir, output_filename)

    total_files = len(file_list)
    num_digits = len(str(total_files))

    start_time = time.time()

    max_filename_length = max([len(os.path.basename(file)) for file in file_list])

    with open(output_file, 'w', encoding='utf-8', newline='') as csv_file:
        for idx, file in enumerate(file_list, start=1):
            file_start_time = time.time()
            progress = (idx / total_files) * 100
            current_filename = os.path.basename(file)
            padding = ' ' * (max_filename_length - len(current_filename))
            percentage_padding = ' ' * (3 - len(f"{progress:.0f}"))
            print(f"\033[1;97m{str(idx).zfill(num_digits)}/{total_files} ({percentage_padding}{progress:.0f}%):\033[0m {current_filename}{padding}", end=" ")  # Bright bold white text

            excel_data = pd.read_excel(file, sheet_name=0, engine='openpyxl')
            excel_data = excel_data.applymap(sanitize_value)
            excel_data.to_csv(csv_file, index=False, header=(idx == 1), quoting=csv.QUOTE_ALL)

            file_end_time = time.time()
            elapsed_time = file_end_time - file_start_time
            print(f"\033[92mDONE: {human_friendly_time(elapsed_time)}\033[0m")  # Bright green text

    end_time = time.time()
    total_elapsed_time = end_time - start_time
    print(f"\033[92mConverted {total_files} files in {human_friendly_time(total_elapsed_time)}\033[0m")  # Bright green text

if __name__ == "__main__":
    input_dir = None
    output_dir = None
    output_filename = None

    if len(sys.argv) >= 2:
        input_dir = sys.argv[1]
    else:
        input_dir = input("Please enter the input directory: ")

    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        output_dir = input("Please enter the output directory: ")

    if len(sys.argv) >= 4:
        output_filename = sys.argv[3]
    else:
        output_filename = input("Please enter the output filename: ")

    combine_xlsx_to_csv(input_dir, output_dir, output_filename)
