# EPN-SEPN Financial Model
___

### Description
___
This is a simple financial model built on Microsoft Excel for automating data entry, cleaning, analysis and visualisation of a Power company's financial data.
This project extracts annual sales records for a Power Company, analysing the data across their networks with respect to customer bands. Analyses is expected to reveal profit/loss trends, consumption patterns, customer churn/retention, etc., facilitating decision-making processes.  

### Table of Contents
___
- [Description](#description)
- [Technologies](#technologies)
- [Data Sources](#data-sources)
- [Usage Instructions](#usage-instructions)
- [Data Cleaning and Preprocessing](#data-cleaning-and-preprocessing)
- [Data Analysis](#data-analysis)
- [Results and Discussions](#results-and-discussions)
- [Recommendations](#recommendations)
- [References](#references)
- [License](#license)


### Technologies
___
- Microsoft Excel


### Data Sources
___
Datasets for this project was obtained from a data company and files named according to the regions within their power network


### Usage Instructions
___
- Data analysis was performed on Microsoft Excel. Download and install MS Excel on your machine. If you have not already, this software can be downloaded [here](https://www.microsoft.com/en-gb/microsoft-365/excel?ef_id=_k_7f6ebb9ae2b216bcec3edc83309dd670_k_&OCID=AIDcmmp20rgnjr_SEM__k_7f6ebb9ae2b216bcec3edc83309dd670_k_&msclkid=7f6ebb9ae2b216bcec3edc83309dd670).
- Ensure all data files are located in the same directory and maintain a consistent naming pattern for all files.
> [!WARNING]
> Saving files on an online platform (e.g. OneDrive) could cause any link in the file to load slower, throwing up warning alerts. So, preferrably, save files on desktop or "C Drive"


### Data Cleaning and Preprocessing
___

On the financial model...
- Data was cleaned using 'trim' function to remove trailing whitespaces; 'ifferror' to assign values (0 for numerical columns or "-" for text columns ) to empty cells that may return '#N/A' errors.
- To get the names of customers on the "Name" column for each sheet, link/extract this from their respective data sheets.
> [!Warning]
> On your excel sheet, ensure the number format for the "Name" column is set to "General"
- Using VLOOKUP function, extract the Residual Charging Bands of each customer from their respective data sheets.
- Populate the Import Fixed Charges (IFC) for each customer from 2020 - 2024 using VLOOKUP function.
- Calculate the Annual Fixed Charge (AFC) (in £) for each customer by dividing IFC (this is in pence) by 100 and multiplying by 365 or 366 (for a leap year)
- Calculate the Year-on-Year percentage change on AFC using the formular "(current year AFC/ previous year AFC)-1"
- Summarise data using pivot table and chart

### Data Analysis
___
```SQL
SELECT *
FROM table;
```

### Results and Discussions
___

### Recommendations
___

### References
___

### License
___
