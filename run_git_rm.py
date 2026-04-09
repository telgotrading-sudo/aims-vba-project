#!/usr/bin/env python3
import subprocess
import os

os.chdir(r"C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project")

# Command 1
cmd1 = [
    "git", "rm", "-f",
    "vba/companies/modules/Module1.bas",
    "vba/companies/modules/Module2.bas",
    "vba/companies/modules/Module3.bas",
    "vba/companies/modules/Module4.bas",
    "vba/companies/modules/Module5.bas",
    "vba/companies/modules/Module6.bas"
]

print("=" * 80)
print("COMMAND 1:")
print(" ".join(cmd1))
print("=" * 80)

try:
    result1 = subprocess.run(cmd1, capture_output=True, text=True)
    print("STDOUT:")
    print(result1.stdout if result1.stdout else "(no output)")
    if result1.stderr:
        print("STDERR:")
        print(result1.stderr)
    print(f"Return code: {result1.returncode}")
except Exception as e:
    print(f"ERROR: {e}")

# Command 2
print("\n" + "=" * 80)
print("COMMAND 2:")
cmd2 = [
    "git", "rm", "-f",
    "vba/aimswrap/modules/Module1.bas",
    "vba/aimswrap/modules/Module2.bas",
    "vba/aimswrap/modules/Module3.bas",
    "vba/aimswrap/modules/Module4.bas",
    "vba/aimswrap/modules/Module5.bas",
    "vba/aimswrap/modules/Module6.bas",
    "vba/aimswrap/modules/Module7.bas",
    "vba/aimswrap/modules/Module8.bas",
    "vba/aimswrap/modules/Module9.bas",
    "vba/aimswrap/modules/Module10.bas"
]

print(" ".join(cmd2))
print("=" * 80)

try:
    result2 = subprocess.run(cmd2, capture_output=True, text=True)
    print("STDOUT:")
    print(result2.stdout if result2.stdout else "(no output)")
    if result2.stderr:
        print("STDERR:")
        print(result2.stderr)
    print(f"Return code: {result2.returncode}")
except Exception as e:
    print(f"ERROR: {e}")
