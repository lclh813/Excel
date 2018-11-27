# Excel
## Obejective

## Demo 1. Line Chart
<img src="https://j.gifs.com/MQMXzG.gif" width="350" height="300" />

### 1.1. Select Data for the Chart
- Formula: ```INDEX``` ```MATCH```
### 1.2. Add an Interactive Vertical Line
- Formula: ```IF```
### 1.3. Dynamically Update Data according to the Location of the Vertical Line
- VBA:   
  - ```Class Module``` Specify which line and which point to be referred to.  
  - ```Module``` Specify to which chart the Class Module is to be applied.    
- Formula: ```INDEX``` ```OFFSET```  
 
## Demo 2. Bubble Chart
### 2.1. Select Data for the Chart
- Formula: ```IF``` ```INDEX``` ```MATCH```
### 2.2. Add an Interactive Legend
- VBA:   
  - ```Microsoft Excel Object``` Determine which legend items has been selected and the corresponding datasets.  

## Demo 3. Cumulative Table
### 3.1. Create a Drop-down List for Start Time
> *Point: Create a drop-down list without blank by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEN``` ```ROW``` ```SMALL```  
### 3.2. Create a Drop-down List for End Time
> *Point 1: End time should always be greater than start time.*  
> *Point 2: Create a drop-down list without blank by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```COUNTIF``` ```DATE``` ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEFT``` ```LEN``` ```MONTH``` ```RIGHT``` ```ROW``` ```OFFSET``` ```SMALL```
### 3.3. Calculate Cumulative Sum
- Formula: ```COLUMN``` ```INDEX``` ```LEFT``` ```LEN``` ```MATCH``` ```RIGHT``` ```SUMPRODUCT```
