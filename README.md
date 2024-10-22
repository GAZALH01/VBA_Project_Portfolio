# My VBA Projects

This repository contains various VBA projects, including data manipulation and inventory valuation scripts.

## Table of Contents
- [Project 1: Data Formatting]
- [Project 2: FIFO Valuation]

## Project 1: Data Formatting

## Overview

This Data formatting script automates the conversion of string data in **Column F** of an Excel shhet into a date format by inserting forward slashes ('/') between day, month, and year. The script handles both **7-digit** and **8-digit** strings:
- If the string is 7 digits long, a leading '0' is added before convering it to a date.
- If the string is 8 digits long, it is converted directly to a date.

### Purpose

Many datasets contain dates stored as unformatted 7- or 8-digit strings (e.g. '1122023' for '01/12/2023'. This script automatically converts these strings into a valid date format ('DD/MM/YYYY'), which allows for easier manipulation and analysis in Excel.

## How the VBA Script Works

-The Script loops through each cell in **Column F**.
-If the string has 7 digits, the script adds a leading zer ('0') to the string.
-If the cell is empty, the script loops to the next cell.
-It then inserts slashes to convert the string to the date format 'DD/MM/YYYY'.
-The string is then converted into an Excel-recognized date value.

## Download

- [Download Input Data](https://raw.githubusercontent.com/GAZALH01/VBA_Project_Portfolio/3dd9a128e090595af434a444ec6cbce985d61363/assets/vba_script_demo.JPG)
  - [Download Output Data](https://raw.githubusercontent.com/GAZALH01/VBA_Project_Portfolio/3dd9a128e090595af434a444ec6cbce985d61363/assets/vba_script_demo_ii.JPG)


## Project 2: FIFO Valuation

## Overview

This project includes a VBA script for calculating FIFO (First In, First Out) inventory valuation.

## How the VBA Script Works

The script calcalates the total FIFO valuation based on the quantity and cost of items in inventory.
