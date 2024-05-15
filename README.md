# Simple Tkinter Input form

Description:
    - This application will take in property information, contact information and others and format a word document and excel document with a fixed format. 


## Setup
To run this application you will need Python (v3.11+ recomended) along with the requirements listed inside the requirements.txt file. It is recomended to utilize pip to download all requirements.

```pip install requirements.txt```


## Running the Application
This application can be run inside any IDE that supports python (VS Code being used) or by using the following CLI command (assuming ```PYTHONPATH``` is set correctly):

```
cd <file-directory>
python3 display.py
```

## How this project works
This project takes user inputs (similar to a form in HTML), formats a word doc, excel doc and label and returns them in a given directory.

*display.py* is used to create the tkinter pages and take the user input

*src/document.py* formats the word document with the inputted data

*src/excel_output.py* formats the excel document with the inputted data

*src/labels.py* formats the printable label as a png utilizing the inputted information.




