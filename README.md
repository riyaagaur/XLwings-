## Xlwings xlwings Excel Automation

This repository contains a collection of Python scripts that automate Excel tasks using xlwings

The scripts included in this repository focus on:
	•	Manipulating data in specific sheets.
	•	Copying and pasting data between workbooks.
	•	Implementing custom “undo” functionality by manually backing up and restoring data.

These examples are especially useful for users working on macOS, where traditional COM automation is not available

Requirements
	•	Python 3.7+
	•	Excel:
A compatible version of Microsoft Excel installed on your system.
	•	xlwings:
The Python package to interact with Excel

Notes for macOS
	•	API Differences:
Traditional COM-based API calls (e.g., sheet.api.Rows(...)) are not available on macOS. This repo uses xlwings’ built-in methods like range.delete() and range.insert() to ensure cross-platform compatibility.
	•	Excel Configuration:
Ensure that your Excel installation is properly configured for use with xlwings. Refer to the [xlwings documentation](https://docs.xlwings.org/) for macOS-specific instructions.

