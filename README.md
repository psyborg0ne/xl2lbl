#Excel To Label - A script to convert structured xml into english and italian label files

How to use:

Clone the repo  
git clone https://github.com/psyborg0ne/xl2lbl

Install python 3.10 or later

Navigate inside the repo folder open a cmd and type: 
python -m venv .venv 
Do not close this window!
This will create an isolated environment to install the necessary libraries and keep the global python package manager clean

Activate the new environment by typing:
.\.venv\Scripts\activate (or activate.ps1 if using powershell in the step above)
This will make the cmd prompt use the isolated interpreter and package manager

Run the script by typing: 
python xl2lbl.py <Excel filename to read data> <English output filename> <Italian output filename>
All of the arguments are optional. 
Defaults are: [Excel File = labels.xlsx] [English Output = en-US.txt] [Italian Output = it-IT.txt]