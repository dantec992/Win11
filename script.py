import pandas as pd
import re

# File paths
csv_file_path = "Software - T20250127.0044.csv"
excel_file_path = "Devices - T20250127.0044.xlsx"

# Windows 11 minimum requirements
MIN_RAM_GB = 4
MIN_BIOS_YEAR = 2018  # TPM 2.0 requirement

WINDOWS_11_CPU_REGEX = re.compile(
    r"(i[3579]-[89]\d{2})|(i[3579]-10\d{2})|(i[3579]-11\d{2})|(i[3579]-12\d{2})|"
    r"(Ryzen [3579] \d{3,4})|(Ryzen PRO)", re.IGNORECASE
)

# Known software issues with Windows 11
INCOMPATIBLE_SOFTWARE = {
    "Adobe Flash Player": "Discontinued, security risk",
    "Internet Explorer": "No longer supported on Windows 11",
    "McAfee Endpoint Security": "Conflicts with Windows 11 security",
    "Symantec Endpoint Protection": "Legacy security software, not compatible",
    "Trend Micro OfficeScan": "Known performance issues",
    "VMware Workstation": "Versions older than 15 may not work",
    "AutoCAD": "Versions older than 2021 may have compatibility issues",
    "Lexmark Printer Software": "Legacy drivers may not function",
    "HP Printer Drivers": "Pre-2018 drivers may not be compatible"
}

try:
    software_df = pd.read_csv(csv_file_path)
    excel_data = pd.ExcelFile(excel_file_path)
    pc_df = excel_data.parse(excel_data.sheet_names[0])  # Load first sheet
except Exception as e:
    print(f"Error loading data: {e}")
    exit()

# Function to check Windows 11 compatibility
def check_windows_11_compatibility(cpu, ram, bios_date):
    if not isinstance(cpu, str) or cpu.lower() == "nan":
        return "Not Compatible (Missing CPU Data)"

    cpu_ok = bool(WINDOWS_11_CPU_REGEX.search(cpu))

    try:
        ram_gb = float(re.search(r"\d+(\.\d+)?", str(ram)).group()) if ram else 0
    except AttributeError:
        ram_gb = 0
    ram_ok = ram_gb >= MIN_RAM_GB

    try:
        bios_year = int(str(bios_date).split("-")[0]) if bios_date else 0
    except ValueError:
        bios_year = 0
    bios_ok = bios_year >= MIN_BIOS_YEAR

    debug_info = f"CPU_OK: {cpu_ok}, RAM_OK: {ram_ok}, BIOS_OK: {bios_ok}"

    if cpu_ok and ram_ok and bios_ok:
        return "Compatible"
    else:
        return f"Not Compatible ({debug_info})"

pc_df["Windows 11 Compatibility"] = pc_df.apply(
    lambda row: check_windows_11_compatibility(
        row.get("Device CPU", ""), row.get("Memory (Usable)", ""), row.get("BIOS Released", "")
    ),
    axis=1
)

incompatible_devices = pc_df[pc_df["Windows 11 Compatibility"].str.contains("Not Compatible")]["Hostname"].tolist()

output_file = "windows_11_final_fixed_report.txt"

with open(output_file, "w") as file:
    file.write("Windows 11 Compatibility Report\n")
    file.write("=" * 60 + "\n\n")

    file.write("PC Compatibility:\n")
    file.write("-" * 60 + "\n")
    for _, row in pc_df.iterrows():
        file.write(f"Hostname: {row.get('Hostname', 'Unknown')}\n")
        file.write(f"CPU: {row.get('Device CPU', 'Unknown')}\n")
        file.write(f"RAM: {row.get('Memory (Usable)', 'Unknown')}\n")
        file.write(f"BIOS Release Date: {row.get('BIOS Released', 'Unknown')}\n")
        file.write(f"Compatibility: {row.get('Windows 11 Compatibility', 'Unknown')}\n")
        file.write("-" * 60 + "\n")

    file.write("\nSoftware Compatibility Issues:\n")
    file.write("-" * 60 + "\n")
    if incompatible_devices:
        file.write("Incompatible Devices:\n")
        for hostname in incompatible_devices:
            file.write(f" - {hostname}\n")
    else:
        file.write("No incompatible devices detected.\n")

incompatible_text_file = "incompatible.txt"
with open(incompatible_text_file, "w") as file:
    if incompatible_devices:
        file.write("Incompatible Devices:\n")
        for hostname in incompatible_devices:
            file.write(f"{hostname}\n")
    else:
        file.write("No incompatible devices detected.")

print(f"Windows 11 compatibility report saved to {output_file}")
print(f"Incompatible device list saved to {incompatible_text_file}")