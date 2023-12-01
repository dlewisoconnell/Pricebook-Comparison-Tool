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
1. Clone repo

2. Navigate to the project directory using the cd command:<br>
   cd (path to your project)<br><br>
   
3. Create a virtual environment named venv. Run the following command:<br>
   python -m venv venv<br><br>
   
4. Activate the virtual environment:<br>
  .venv\Scripts\activate  # For Windows<br>
  source venv/bin/activate  # For Unix/Linux<br><br>
  Once activated, your terminal prompt should change to indicate that the virtual environment is active.<br><br>

5. Install project dependencies listed in the requirements.txt file:<br>
   pip install -r requirements.txt<br><br>

6. With the virtual environment active, you can now run the project:<br>
  python python.py<br><br>
  This will create the changes.xls file and plot in the data folder, showing week to week changes and most signifigant price changes in pivot table and plot.

7. When you're done working on the project, deactivate the virtual environment:<br>
   deactivate


  
