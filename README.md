# Excel
## Obejective
Create charts and tables that are easy to track data across time.   
## Demo 1. Line Chart
### 1.1. Introduction  
> **_Action:_** *Move cursor to a specific point in the line chart.*  
> **_Result:_** *Cells that display corresponding time and data change accrodingly.*     
<img src="https://j.gifs.com/nrBVMW.gif" width="400" height="300" />

### 1.2. Steps  
**_Step 1: Select Data for the Chart_**  
 - Formula: ```INDEX``` ```MATCH```    
**_Step 2: Add an Interactive Vertical Line_**    
- Formula: ```IF```    
**_Step 3: Dynamically Update Data according to the Location of the Vertical Line_**    
- VBA:   
  - ```Class Module``` Specify which line and which point to be referred to.    
  - ```Module``` Specify to which chart the Class Module is to be applied.      
- Formula: ```INDEX``` ```OFFSET```    
 
## Demo 2. Bubble Chart
<img src="https://j.gifs.com/pQEqWp.gif" width="500" height="300" />    

### 2.1. Select Data for the Chart
- Formula: ```IF``` ```INDEX``` ```MATCH```
### 2.2. Add an Interactive Legend
- VBA:   
  - ```Microsoft Excel Object``` Determine which legend items has been selected and the corresponding datasets.  

## Demo 3. Cumulative Table
<img src="https://j.gifs.com/jqxkWl.gif" width="600" height="240" />
 
### 3.1. Create a Drop-down List for Start Time
> *Point: Create a drop-down list without blank by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEN``` ```ROW``` ```SMALL```  
### 3.2. Create a Drop-down List for End Time
> *Point 1: End time should always be greater than start time.*  
> *Point 2: Create a drop-down list without blank by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```COUNTIF``` ```DATE``` ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEFT``` ```LEN``` ```MONTH``` ```RIGHT``` ```ROW``` ```OFFSET``` ```SMALL```
  - ```=IFERROR(
                   OFFSET($AD$1,SMALL(IF(ISERROR(FIND("合計",AD:AD)),
                                                               ROW($AD:$AD),
                                                               ""),
                                                           ROW(4:4))-1,0),"")```
### 3.3. Calculate Cumulative Sum
- Formula: ```COLUMN``` ```INDEX``` ```LEFT``` ```LEN``` ```MATCH``` ```RIGHT``` ```SUMPRODUCT```
