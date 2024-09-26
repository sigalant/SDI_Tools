# SDI_Tools
## Pre-Requisites

Before using these tools, you must first follow these steps:

Install Python3

Clone this repository (Instructions here: https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository )

Enter the following commands into the command prompt:

```
pip install -r requirements.txt
```


## How to use

Each application has a help menu that opens an html with application-specific instructions on how to use it.

## Budget Formatting Tool

Build from terminal with: 
```
python -m PyInstaller -F -w -i ./data/Shared/SDI_Logo.ico --add-data './data/Shared/:.' --add-data './data/BudgetTool/:.' BudgetingTool.py 
```

## Specs Generation Tool

Build from terminal with: 
```
python -m PyInstaller -D -w -i ./data/Shared/SDI_Logo.ico --add-data './data/Shared/:.' --add-data './data/SpecTool/:.' SpecTool.py
```
## Utility Loads Formatting Tool

Build from terminal with: 
```
python -m PyInstaller -D -w -i ./data/Shared/SDI_Logo.ico --add-data './data/Shared/:.' --add-data './data/UtilityTool/:.' BulkLoads.py
```
