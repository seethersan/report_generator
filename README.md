# SITE STATS REPORT GENERATOR

This project reads site stats from a XLSX file and return a XLSX file with the summary report per site and date

This project use Python 3.10+ and openpyxl 3.1.1

This report generator can handle any number of sites and stats and will ignore a value if the stat or the date is empty

The dates will be filtered using A1, A2 values

## Run the project

To run this project you need to create a virtualenv with the following command

```bash
virtualenv .env
source .env/bin/activate
```

Then you need to install requirements

```bash
pip install -r requirements.txt
```

Finally run the following command

```bash
python report_generator.py
```

To run the test execute

```bash
python test_report_generator.py
```
