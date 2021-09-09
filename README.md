# PerformanceTest
PerformanceTest is a PowerShell script that measures adding, getting, and opening time of SAP Business One and CompuTec ProcessForce objects through SAP Business One DI/UI API and CompuTec ProcessForce API.

> :warning: You must run it only on a test database and not any productive database as the script will create dummy objects in the database.

## Installation

> :information_source: SAP Business One client and PowerShell script have to run as the same user and platform bitness (x86/x64).

Download all files from the repo to a folder on a machine where SAP Business One DI API, SAP Business One client, and CompuTec ProcessForce API are installed.

## Settings
In the file `conf/Connection.xml`, define the company database connection information so that the script can add through DI API objects used in the SAP Business One UI tests.

Out of the box, we provide three test scopes: 1 - Short, 2 - Medium, 3 - Long, where 1 (Short) is the default. You can choose the scope in the `Test` element in the `conf/Connection.xml` file.

The test scope can be customized if needed in the `conf/TestConfig*.xml` files.

## How to run
Start SAP Business One & ProcessForce on the same company database as defined in the `conf/Connection.xml` file.

Run PowerShell script PerformanceTest.ps1 from PowerShell terminal, Visual Studio Code, or PowerShell ISE.

## What the script measures

The scripts measures adding these objects to the company database through DI API:
* Item Master Data
* Item Details
* BOM
* Resources
* Operations
* Routings
* Production Process

and from SAP Business One client UI:
* Opening above objects' form in regular and maximized mode
* Switching between them

## Results
Measurements are written to folder RESULTS_[CurrentDate]_[CurrentTime] with two files in it: `Result_Enviroment.csv` and `Results_Details.csv`.

`Result_Enviroment.csv` contains basic information about the host on which the script is run, like memory, processor, and network responses between the host, database server, and the SLD server.

`Results_Details.csv` contains detailed information about the time spent on each operation defined in the test configuration file.