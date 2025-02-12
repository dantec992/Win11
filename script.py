import pandas as pd
import re

# File paths
csv_file_path = "Software - T20250127.0044.csv"
excel_file_path = "Devices - T20250127.0044.xlsx"

# Windows 11 minimum requirements
MIN_RAM_GB = 4
MIN_BIOS_YEAR = 2018  # TPM 2.0 requirement

# ✅ Expanded CPU detection using a more flexible approach
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

# Load device and software data
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

    # ✅ More flexible CPU detection
    cpu_ok = bool(WINDOWS_11_CPU_REGEX.search(cpu))

    # ✅ Improved RAM detection (extracts numeric part)
    try:
        ram_gb = float(re.search(r"\d+(\.\d+)?", str(ram)).group()) if ram else 0
    except AttributeError:
        ram_gb = 0
    ram_ok = ram_gb >= MIN_RAM_GB

    # ✅ Extract BIOS year properly
    try:
        bios_year = int(str(bios_date).split("-")[0]) if bios_date else 0
    except ValueError:
        bios_year = 0
    bios_ok = bios_year >= MIN_BIOS_YEAR

    # Debugging output (shows reason for failure)
    debug_info = f"CPU_OK: {cpu_ok}, RAM_OK: {ram_ok}, BIOS_OK: {bios_ok}"

    if cpu_ok and ram_ok and bios_ok:
        return "Compatible"
    else:
        return f"Not Compatible ({debug_info})"

# Apply Windows 11 compatibility check
pc_df["Windows 11 Compatibility"] = pc_df.apply(
    lambda row: check_windows_11_compatibility(
        row.get("Device CPU", ""), row.get("Memory (Usable)", ""), row.get("BIOS Released", "")
    ),
    axis=1
)

# Function to check software compatibility
def check_software_compatibility(software_name, version):
    for known_software, issue in INCOMPATIBLE_SOFTWARE.items():
        if known_software.lower() in software_name.lower():
            return f"Incompatible - {issue} (Installed: {version})"
    return "Compatible"

# ✅ Track incompatible software per PC
incompatible_software_per_pc = {}

for _, row in software_df.iterrows():
    hostname = row["Device Hostname"]
    software_name = row["Software"]
    software_version = row["Version"]

    compatibility_status = check_software_compatibility(software_name, software_version)

    if compatibility_status.startswith("Incompatible"):
        if hostname not in incompatible_software_per_pc:
            incompatible_software_per_pc[hostname] = []
        incompatible_software_per_pc[hostname].append(f"{software_name} (v{software_version}) - {compatibility_status}")

# Generate output report
output_file = "windows_11_final_fixed_report.txt"

with open(output_file, "w") as file:
    file.write("Windows 11 Compatibility Report\n")
    file.write("=" * 60 + "\n\n")

    # ✅ Write PC compatibility report
    file.write("PC Compatibility:\n")
    file.write("-" * 60 + "\n")
    for _, row in pc_df.iterrows():
        file.write(f"Hostname: {row.get('Hostname', 'Unknown')}\n")
        file.write(f"CPU: {row.get('Device CPU', 'Unknown')}\n")
        file.write(f"RAM: {row.get('Memory (Usable)', 'Unknown')}\n")
        file.write(f"BIOS Release Date: {row.get('BIOS Released', 'Unknown')}\n")
        file.write(f"Compatibility: {row.get('Windows 11 Compatibility', 'Unknown')}\n")
        file.write("-" * 60 + "\n")

    # ✅ Write software compatibility report
    file.write("\nSoftware Compatibility Issues:\n")
    file.write("-" * 60 + "\n")
    if incompatible_software_per_pc:
        for hostname, issues in incompatible_software_per_pc.items():
            file.write(f"\nDevice: {hostname}\n")
            for issue in issues:
                file.write(f" - {issue}\n")
    else:
        file.write("No incompatible software detected.\n")

print(f"Windows 11 compatibility report saved to {output_file}")

# ------------------------------------------------------------------
# Lists all incompatible computers
# ------------------------------------------------------------------
incompatible_file = "windows_11_incompatible_computers.txt"
incompatible_dict = {}

# First, gather hardware-incompatible machines
for _, row in pc_df.iterrows():
    hostname = row.get("Hostname", "Unknown")
    comp_status = row.get("Windows 11 Compatibility", "")
    if "Not Compatible" in comp_status:
        # Add a note about hardware
        incompatible_dict.setdefault(hostname, []).append(f"Hardware Requirements - {comp_status}")

# Next, gather software-incompatible machines
for hostname, issues in incompatible_software_per_pc.items():
    # If no entry for this hostname yet, create one
    if hostname not in incompatible_dict:
        incompatible_dict[hostname] = []
    # Add a note about software (list each issue in one line or multiple lines)
    for issue in issues:
        incompatible_dict[hostname].append(f"Software Requirements - {issue}")

# Write out the new file
with open(incompatible_file, "w") as f:
    f.write("List of Incompatible Computers\n")
    f.write("=" * 60 + "\n\n")

    if not incompatible_dict:
        f.write("No incompatible computers found.\n")
    else:
        for host, reasons in incompatible_dict.items():
            f.write(f"Hostname: {host}\n")
            for r in reasons:
                f.write(f" - {r}\n")
            f.write("\n")

print(f"Incompatible computers list saved to {incompatible_file}")