# BLS - CPI Extractor

This script is responsible for download data from the BLS website| most specifically from the Consumer Price Index Supplemental Files page.

## How to use this script

Please follow the following steps:

- Download the python runtime| if you do not have available. This script was tested with Python 3.9.X or above
- Download or clone this repository

```bash
# To clone a repo
$ git clone https://github.com/WalterPaixaoCortes/bls-cpi-extract.git
```

- On the folder where the script was downloaded or cloned| create a virtual environment

```bash
$ python -m venv /path/to/new/virtual/environment
```

- Install the required components (see the command below)

```
$ pip install -r requirements.txt
```

If you have executed all the steps above| you should be able to execute the script.

```bash
# To run the script
$ python main.py
```

## Results

If the script have executed successfully| the folders input and output should have been filled with files.

The input folder should have zip files and Excel files with the content required.

The output folder should have the processed_data.csv file| which contains the data that is required. See a sample below:

| Year | Month | Expenditure category           | Relative importance | Unadjusted percent change | Unadjusted effect on All Items |
| ---- | ----- | ------------------------------ | ------------------- | ------------------------- | ------------------------------ |
| 2014 | 12    | Food                           | 14.131              | 3.4                       | 0.473                          |
| 2014 | 12    | Energy                         | 8.443               | -10.6                     | -0.955                         |
| 2014 | 12    | All items less food and energy | 77.426              | 1.6                       | 1.238                          |
| 2014 | 12    | All items less food            | 85.869              | 0.3                       | 0.283                          |
| 2014 | 12    | All items less shelter         | 67.518              | -0.3                      | -0.173                         |
