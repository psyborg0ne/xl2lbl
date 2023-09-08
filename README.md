# Excel To Label
A script to convert structured Excel files into english and italian label files

## How to use:

### Initial setup
Install python 3.10 or later. Clone the repo ```git clone https://github.com/psyborg0ne/xl2lbl```

### Setting a virtual environment
Open a cmd/powershell and navigate inside the folder you cloned above. Type ```python -m venv .venv``` and leave this window open.
This will create an isolated virtual environment for you to install the necessary libraries and keep the global python package manager clean.

### Activate the new environment
In the same window, type
  * Cmd
    * ```.\.venv\Scripts\activate```
  * Powershell
    * ```.\.venv\Scripts\activate.ps1```

This will make the shell use the isolated interpreter and package manager.
Shell prompt should now begin with **( .venv )**
That means, every package you install with **pip** from now on will be installed inside the virtual environment

### Install dependencies
2 library packages are needed as of the time of writting.

 * **pandas**
     * Open, read and manipulate data in Excel files
 * **openpyxl**
     * One of the engines used by pandas to work with Excel files.
     * You can read more about pandas excel handling [here](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html) under **engine**

There are 2 ways to install the libraries
  * **Manual**
    * Use ```pip install pandas openpyxl``` while inside a shell with an active virtual environment. The caveat here is that this command will get all the latest versions for each package which may cause incompatibilities and strange behaviour
  * **Automatic** *(Recommended)*
    * Use ```pip install -r path/to/requirements.txt``` while inside a shell with an active virtual environment. This will use the ```requirements.txt``` file provided in the repository, with the correct versions of the packages needed.

### Run the script
Type ```python xl2lbl.py <path/to/excel/file> <path/to/output_english.txt> <path/to/output_english.txt>```
The script accepts 3 command line arguments. All of the arguments are **optional**.

| Argument | Default |
| -------- | ------- |
| ```path/to/input_excel.xlsx``` | "Labels.xlsx" |
| ```path/to/label_output_en.txt``` | "en-US.txt" |
| ```path/to/label_output_it.txt``` | "it-IT.txt" |
