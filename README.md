# SDI_Tools

## Budget Formatting Tool

Build from terminal with: 
```
python -m PyInstaller -F -w -i ./data/SDI_Logo.PNG BudgetingTool.py 
```

## Specs Generation Tool

Build from terminal with: 
```
python -m PyInstaller -D --add-binary SpecDB.py:. -w -i ./data/SDI_Logo.ico SpecTool.py
```
## Utility Loads Formatting Tool

Build from terminal with: 
```
python -m PyInstaller -D --add-binary LogErrors.py:. -w -i ./data/SDI_Logo.ico BulkLoads.py
```
