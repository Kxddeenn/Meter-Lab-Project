# Certification and Testing Components for Meter Certification Project

This folder contains components and scripts for certification generation, XML reading, CSV reading and UI components related to electrical meters.

## Components Overview

### 1. Excel Files

- **SealingLogTest.xlsx**: Excel workbook used for logging sealing information. (Also is the historic log for the sealed meters)
- **emptyCertificate.xlsx**: Excel workbook for all the blank certificate data. 
- **WECO4150_CertErrors.xlsx**: Excel workbook containing error data specific to the WECO4150 bench.
- **WECO2350_CertErrors.xlsx**: Excel workbook containing error data specific to the WECO2350 bench.

### 2. Scripts

- **certNumb.py**: Python script to generate a new certificate number based on the latest entry in `SealingLogTest.xlsx`.
- **exportSealing.py**: Python script to export sealing information from `modifiedCert.xlsx` to `SealingLogTest.xlsx`.
- **exportErrors.py**: Python script for extracting data in `WECO2350_CertErrors.xlsx` or `WECO4150_CertErrors.xlsx` depending on meter type and bench.
- **loadRawData.py**: Python script for reading CSV files with `pandas` and saving it into arrays. 
- **gatewayConfig.py**: Python script for reading JSON files using in-built `Json` libraries.
- **readXML.py**: Python script for reading XML files and converting it into variables for later use.
- **settingFunctions.py**: Python script for the UI's page settings. Handles opening excel files and manipulating `emptyCertificate.xlsx`.
- **updateCompleted.py**: Python script for updating the table picture for completed meters in the `completed tab`.
