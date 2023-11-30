# Pricebook Comparison Tool
Using Python to compare two weeks of produce prices in excel format, copying any price changes to a new file. Also uses matplotlib and seaborn to create a pivot table and png file graph showing the most signifigant price increases. Project used in real world application.

# Features
* Read TWO data files (Excel)
* Clean data and perform a pandas merge with two data sets, then calculate some new values based on the new data set
* Make at least 1 Pandas pivot table and 1 matplotlib/seaborn plot
* Utilize a virtual environment and include instructions in README on how the user should set one up
* Annotate your .py files with well-written comments and a clear README.md

# Prerequisites
* Python

# Usage
1. Navigate to the project directory using the cd command:
   cd (path to your project)
   
2. Create a virtual environment named venv. Run the following command:
   python -m venv venv
   
3. Activate the virtual environment:
  source venv/Scripts/activate  # For Windows
  source venv/bin/activate  # For Unix/Linux
  Once activated, your terminal prompt should change to indicate that the virtual environment is active.

4. Install project dependencies listed in the requirements.txt file:
   pip install -r requirements.txt

5. With the virtual environment active, you can now run the project:
  python.py
  This will create the changes.xls file and plot in the data folder, showing week to week changes and most signifigant price changes in pivot table and plot.

6. When you're done working on the project, deactivate the virtual environment:
   deactivate


  
