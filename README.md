# SolidWorks Python Automation

This project provides a collection of Python functions to automate SolidWorks workflows using the SolidWorks COM API (pywin32). It enables parametric modification of parts, batch generation of variants, and extraction of model properties.

---

## Features

- Modify dimensions of SolidWorks parts programmatically  
- Read and update global variables  
- Generate multiple part variants from a CSV file  
- Compute mass properties of parts and assemblies  
- Perform parameter sweeps for mass analysis  
- Diagnose and inspect model global variables  
- Test SolidWorks COM connectivity  

---

## Available Functions

- assembly_mass_sweep – compute mass variations across parameters or configurations  
- change_dimension – modify specific dimensions in a part  
- change_globals – update global variables  
- diagnose_globals – inspect and debug global variables  
- generate_variants – create part variants based on CSV input  
- get_mass – retrieve mass properties  
- read_globals – read global variables  

Utility / Test Scripts:

- test_COM  
- test_pywin32  
- test_solidworks_com  

These are used to verify that the SolidWorks COM interface is working correctly.

---

## Requirements

Install dependencies with:

pip install -r requirements.txt

Main dependencies:
- pywin32  
- pandas  
- numpy  
- matplotlib (optional)  
- openpyxl (if using Excel files)  

---

## Project Structure

src/  
  assembly_mass_sweep.py  
  change_dimension.py  
  change_globals.py  
  diagnose_globals.py  
  generate_variants.py  
  get_mass.py  
  read_globals.py  

data/  
  box.sldprt  
  variants.csv  

tests/  
  test_COM.py  
  test_pywin32.py  
  test_solidworks_com.py  

output/  

requirements.txt  
.gitignore  
README.md  

---

## Usage

Run a script from the src folder. For example:

python src/generate_variants.py

Typical workflow:
1. Place the base SolidWorks file (box.sldprt) in the data/ folder  
2. Define parameters in variants.csv  
3. Run the desired script  
4. Check generated parts in the output/ folder  

Make sure:
- SolidWorks is installed and running  
- The COM interface is accessible  
- Paths inside scripts are correctly set  

---

## Inputs

The scripts require the following input files:

- data/box.sldprt  
  Base SolidWorks part used as a template model  

- data/variants.csv  
  CSV file defining dimensions or parameters for each variant  

Example CSV structure:

length,width,height  
100,50,30  
120,60,40  

Each row represents one generated variant.

---

## Outputs

Generated files are saved in:

output/

Outputs include:
- New SolidWorks part files (.sldprt)  
- One file per CSV entry  
- Variants with modified dimensions or global variables  

---

## Notes

- Works on Windows with SolidWorks installed  
- Uses the SolidWorks COM API via pywin32  
- SolidWorks must be properly configured for automation  
- Test scripts in tests/ help verify connectivity  
