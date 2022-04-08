import datetime

DOWNLOAD_XLSX_URL = "https://www.bls.gov/cpi/tables/supplemental-files/news-release-table7-%s%02d.xlsx"
DOWNLOAD_ZIP_URL = "https://www.bls.gov/cpi/tables/supplemental-files/archive-%s.zip"

XSLX_FILE_NAME = 'input/table7-%s%02d.xlsx'
ZIP_FILE_NAME = 'input/archive-%s.zip'
EXTRACT_FILE_NAME = {2020: 'archive-%s/news-release-table7-%s%02d.xlsx',
2019: 'archive-%s/news-release-table7-%s%02d.xlsx',
2018: '%s/news-release-table7-%s%02d.xlsx',
2017: '%s/news-release-table7-%s%02d.xlsx',
2016: 'supplemental-files-archive-%s/CpiPress7_FINAL_%s%02d.xlsx',
2015: 'supplemental-files-archive-%s/CpiPress7_FINAL_%s%02d.xlsx',
2014: 'supplemental-files-archive-%s/CpiPress7_FINAL_%s%02d.xlsx'}
INPUT_FOLDER = 'input/'

START_YEAR = datetime.date.today().year # should be replaced by Datetime.Year function
END_YEAR = 2014

PROCESSED_DATA = 'output/processed_data.csv'
