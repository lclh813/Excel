# Excel
## Obejective
Create charts and tables that are easy to track data across time.   
## Demo 1. Line Chart
### 1.1. Introduction
**_Action:_** *Move cursor to a specific point in the line chart.*  
**_Result:_** *Cells that display corresponding time and data change accrodingly.*

<img src="https://j.gifs.com/nrBVMW.gif" width="400" height="300" />

### 1.2. Steps  
#### **_Step 1. Select Data for the Chart_**    
- Formula: ```INDEX``` ```MATCH```        
#### **_Step 2. Add an Interactive Vertical Line_**      
- Formula: ```IF```      
#### **_Step 3. Dynamically Update Data according to the Location of the Vertical Line_**    
- VBA:     
  - ```Class Module``` Specify which line and which point to be referred to.      
  - ```Module``` Specify to which chart the Class Module is to be applied.        
- Formula: ```INDEX``` ```OFFSET```    
 
## Demo 2. Bubble Chart
### 2.1. Introduction  
**_Action:_** *Click on a specific legend item.*  
**_Result:_** *Only the selected one is to be shown.*

<img src="https://j.gifs.com/pQEqWp.gif" width="500" height="300" />    

### 2.2. Steps  
#### **_Step 1. Select Data for the Chart_**
- Formula: ```IF``` ```INDEX``` ```MATCH```
#### **_Step 2. Add an Interactive Legend_**
- VBA:   
  - ```Microsoft Excel Object``` Determine which legend items has been selected and show the corresponding datasets.  

## Demo 3. Cumulative Table
### 3.1. Introduction  
**_Action 1:**_ *Select a specific time from a drop-down list for **_Start Time_**.*  
**_Result 1:_** *Drop-down list for **_End Time_** starts one month later than the selected **_Start Time_**.* 

**_Action 2:_** *Select a specific time from a drop-down list for **_End Time_**.*   
**_Result 2:_** *Data for the specific period of time is to be shown.*    

<img src="https://j.gifs.com/jqxkWl.gif" width="600" height="240" />

### 3.2. Steps  
#### **_Step 1. Create a Drop-down List for Start Time_**
> **_Point:_** *Create a drop-down list without blanks by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEN``` ```ROW``` ```SMALL```  
#### **_Step 1.1. Get the List of Time from the Monthly Dataset_**
> **_Point:_** *There are 7 columns per set of monthly data.*
```
=IF(ISBLANK(INDEX(Data!$6:$6,1,(ROW())*7)),
    "",
    INDEX(Data!$6:$6,1,(ROW())*7))
```   
#### **_Step 1.2. Create a List of Time without Blanks_**
> **_Point 1:_** *Romove **_Annual Total_** from the list of time.*  
> **_Point 2:_** *Remove blanks from the list.*  
```
=IFERROR(OFFSET($AD$1,
                SMALL(IF(ISERROR(FIND("Total",AD:AD)),
                         ROW($AD:$AD),
                         ""),
                      ROW(1:1))-1,
                0),
         "")
```     
#### **_Step 1.3. Define a Name for Start Time_**
> **_Point:_** *Name cells with formula results and exclude those not showing formula results.*    
```
=IF(LEN(AE1)>0,1,0)
```
```
=OFFSET(Data!$AE$1,
        0,0,
        COUNTIF(Data!$AF:$AF,1),1)
```  

#### **_Step 2. Create a Drop-down List for End Time_**  
> **_Point 1:_** *End time should always be greater than start time.*  
> **_Point 2:_** *Create a drop-down list without blank by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```COUNTIF``` ```DATE``` ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEFT``` ```LEN``` ```MONTH``` ```RIGHT``` ```ROW``` ```OFFSET``` ```SMALL```  
#### **_Step 2.1. Make Sure _**
#### **_Step 3. Calculate Cumulative Sum_**  
- Formula: ```COLUMN``` ```INDEX``` ```LEFT``` ```LEN``` ```MATCH``` ```RIGHT``` ```SUMPRODUCT```
