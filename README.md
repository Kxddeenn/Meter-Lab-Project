# Meter Lab Certificate Generation 💻

The purpose of this project is to find a solution to the VBA macros in the Excel files. 
VBA is arguably a dated language and there are many issues with its compatibility especially with growing technology around the world. 

Currently, the team is using a variety of macro scripts to automate the generation of the certificates.
This was effective maybe 10 years ago, but today the world uses more complex, smoother code to automate processes. 

A brief overview of the project was to have a User Interface, and better functionality to lessen the risk of potential damage and make the process smoother. 

## Table of Contents 📜

- [Installation](#installation)
- [Usage](#usage)
- [Features](#features)

## Installation 🛠️

Follow the code below in the terminal in order to set up in your local IDE. 
This project is entirely written in Python, so the User must have the required dependencies. 

```bash
# Do this in Github
$ git clone "https://github.com/Kxddeenn/Meter-Lab-Project.git"
# Do this in Terminal
$ cd "use the folder's location (copy address)"
$ pip install requirements.txt
```
## Usage 🚀
How to launch the application. (This app will not work because of Excel Files Permissions)
```bash
$ cd app.py
# Launch the application
$ py app.py 

# OR

# Click app.py in folder view (Make sure to have dependencies installed)
```

If there's any updates 
```bash
$ cd "use folder's location"
$ git pull
```

## UI Features ✨
- **Tabbed Interface**: The application offers a clean, tabbed interface with the following sections:
  - **Main**: The primary data entry and file submission tab.
  - **Weekly Jig**: For managing weekly calibration jigs.
  - **Settings**: Customize the application's settings and preferences.
  - **Completed**: View and manage completed certificate generation tasks.


### Main Screen

![Main Screen](https://github.com/Kxddeenn/Meter-Lab-Project/blob/main/UI/images/filledout.png)

- **Required Data Section**: The main form includes fields for entering data:
  - **Product Type**: Dropdown to select the type of meter (e.g., 6312).
  - **Voltage**: Input for voltage specification (e.g., 347V).
  - **Customer/Owner**: Dropdown to select or input the customer or owner.
  - **Verification**: Dropdown to select the verification type (e.g., Re-verified).
  - **Customer Name**: Input for the customer name.
  - **Firmware**: Dropdown for firmware version (e.g., 2.08).
  - **Address**: Input for the customer's address.
  - **Badge Number**: Input for the meter's badge number (e.g., MSI0004).
  - **Regulation #**: Input for the company's regulation number.

- **File Submission**: Users can submit XML and CSV files for certificate generation.
  - Buttons for selecting the respective files are highlighted in green once a file is chosen.

- **Generate Certificate**: A button that initiates the certificate generation process based on the input data and selected files.
  - Allows the user to view the certificate they generated
  - Exports the certificate to the completed folder. 


### Weekly Jig Screen 

![Weekly Jig](https://github.com/Kxddeenn/Meter-Lab-Project/blob/main/UI/images/testjig.png)

- **Required Data Entry**: 
  - **Room Temp in C**: Input field to record the current room temperature in Celsius.
  - **Bench**: Dropdown to select the bench used for testing.
  - **Test Jig**: Dropdown for selecting the specific test jig.
  - **Voltage**: Input or dropdown to specify the voltage used during testing.
  - **Meter Points**: Field to input the number of meter points being tested.

- **Submit Files**:
  - **Data File (.xlsx)**: Button to upload the data file in .xlsx format, containing the test results.
  - **Generate Certificate**: Button that processes the input data and generates the certificate.
  - **Open Empty Certificate & Errors**: Option to open a blank certificate template and view any errors. This also allows you to update the certificate or change any errors.


### Settings Screen

![Setting Screen](https://github.com/Kxddeenn/Meter-Lab-Project/blob/main/UI/images/settings.png)

- **Test Jig**:
  - **Last Comparison Test**: Input field where users can log the date of the last comparison test performed.
  - **Bench**: Dropdown menu to select the bench that is currently being used for testing. The selected bench is shown as "WECO4150" in this example.
  - **Save Button**: After making any changes, users can click the "Save" button to store the new settings.

- **Update Files**:
  - **Certificate Template**: Button to open and modify the template used for generating certificates. This ensures that the certificates match to the latest standards and formats.
  - **Meter Errors File**: Button to open and update the file that has all meter errors used in equations within the certification generation. 

- **Support**:
  - **Feedback Forum**: Button to open a link where users can provide feedback on the application.
  - **Documentation**: Button to open a link to a .loop page where users can see all processes, installation guides or error troubleshooting. 


### Completed Screen

![Completed Screen](https://github.com/Kxddeenn/Meter-Lab-Project/blob/main/UI/images/completed.png)

- **Meter Seals Completed**:
  - **Table**: A table to represent the amount of meter seals per month.
  - **Refresh Button**: This will refresh the table above, showing the total amount of meter seals within the lab.
  - **Open Sealing Log Button**: Button used to open the sealing log which stores all the certificates numbers and data for each creation.


