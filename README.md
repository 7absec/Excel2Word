# Generate Vulnerability Report in word from excel

[![GitHub Repo](https://img.shields.io/badge/GitHub-Repository-blue)](https://github.com/7absec/Excel2Word)
![Banner](https://github.com/7absec/Excel2Word/raw/main/Banner.png)

## Description
The Excel to Vulnerability Report Template Converter is a Python program designed to generate vulnerability report templates in Word format from single or multi-sheet Excel files. This tool is particularly useful when you have an Excel workbook with specific columns representing vulnerability data, and you want to create standardized vulnerability reports based on the data in those columns.


### Example Excel File Columns
The input Excel file should have the following columns, just like the example Excel file provided:

|   column1   |   column2   |   column3   |   column4   |   column5   |   column6   |   column7   |   column8   |
|-------------|-------------|-------------|-------------|-------------|-------------|-------------|-------------|
|   Target    | Vulnerability Name |   Severity  |    CVSS     |  Parameter  | Description |   Impact    | Remediation |

### Directory Structure
To organize your project, you can follow this directory structure:
- Image folder and fiels should be in the Ascending Order as shown below
- Folder and file names must be in the following format
  - 1_folderName   (the numeric order must be separate by and underscore (_))
  - 1_imageName

For single sheet 
```
Root Dir (e.g., eample.com)
│
├── 1_subdirectory (e.g., RCE)
│   ├── 1_anything.png
│   │── 2_anything.png
│   │── ...
|
├── 2_subdirectory (e.g., XSS)
│   ├── 1_anything.png
│   │── 2_anything.png
│   ├── ...
│
└── ...
```
For multiple sheet
```
Root Dir (e.g., Eample)
│
├── 1_subdirectory (e.g., eample.com)
│   ├── 1_secondsubdir (e.g., RCE)
│   │   ├── 1_anything.png
│   │   ├── 2_anything.png
│   │   └── ...
│   ├── 2_secondsubdir (e.g., XSS)
│   │   ├── 1_anything.png
│   │   ├── 2_anything.png
│   │   └── ...
│   └── ...
│
├── 2_subdirectory (e.g., google.com)
│   ├── 1_secondsubdir (e.g., RCE)
│   │   ├── 1_anything.png
│   │   ├── 2_anything.png
│   │   └── ...
│   ├── 2_secondsubdir (e.g., XSS)
│   │   ├── 1_anything.png
└── ...
```


<h3 align="left">Languages and Tools:</h3>
<a href="https://www.python.org/" target="_blank"> 
  <img src="https://github.com/devicons/devicon/blob/master/icons/python/python-original.svg" alt="css3" width="40" height="40"/> 
</a>

## Installation

```sh
git clone https://github.com/7absec/Excel2Word.git
cd Excel2Word
pip install -r requirements.txt --upgrade  or pip3 install -r requirements.txt --upgrade
python excel2word.py  or python3 excel2word.py
```


## Usage
Script takes 4 input (1 file input, 2 folder input, and 1 string input)
Below is the input sequence as per the script

- Excel File --- This is your excel file
- Image Folder --- This the root folder for the images (Root Dir as per above mentioned structure)
- Output Folder --- This is the output folder where the word file will be saved 
- Client name --- This is the name of your client for the given report. 


## Bonus
Test image directories and excel files are added to the repository to test the code. :D

## Contact
Feel free to reach out to me on [Twitter](https://twitter.com/7absec) or [LinkedIn](https://linkedin.com/in/7absec) for any questions or feedback or if you want to modify the excel(input) or word(output) files as per your requirement.

If you find this project helpful, consider to support further development.

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/7absec)
