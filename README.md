# Excel
## Objective
Create interactive charts and tables to track data across time.   
## Demo 1. Line Chart
### 1.1. How to Use
**_Action:_** *Move cursor to the specific point in the line chart.*  
**_Result:_** *Cells that display corresponding time and data change accrodingly.*

<img src="https://j.gifs.com/jqyy9W.gif" width="400" height="300" />

### 1.2. How to Create  
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
### 2.1. How to Use 
**_Action:_** *Click on the specific legend item.*  
**_Result:_** *Only the selected item will be shown in the bubble cahrt.*

<img src="https://j.gifs.com/E9DDLK.gif" width="500" height="300" />    

### 2.2. How to Create  
#### **_Step 1. Select Data for the Chart_**
- Formula: ```IF``` ```INDEX``` ```MATCH```
#### **_Step 2. Add an Interactive Legend_**
- VBA:   
  - ```Microsoft Excel Object``` Determine which legend item has been selected and show the corresponding data.  

## Demo 3. Cumulative Table
### 3.1. How to Use  
**_Action 1:_** *Select a specific time from a drop-down list for **_Start Time_**.*  
**_Result 1:_** *Drop-down list for **_End Time_** starts one month later than the selected **_Start Time_**.* 

**_Action 2:_** *Select a specific time from a drop-down list for **_End Time_**.*   
**_Result 2:_** *Data for the specific period of time will be displayed.*    

<img src="https://j.gifs.com/jqxkWl.gif" width="600" height="240" />

### 3.2. How to Create  
#### **_Step 1. Create a Drop-down List for Start Time_**
> **_Point:_** *Create a drop-down list without blanks by ignoring cells without showing formula results.*    
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
> **_Point 1:_** *Differentiate cells with formula results from those not showing formula results.*    
```
=IF(LEN(AE1)>0,1,0)
```
> **_Point 2:_** *Define a name for a cell range showing formula results.*      
```
=OFFSET(Data!$AE$1,
        0,0,
        COUNTIF(Data!$AF:$AF,1),1)
```  

#### **_Step 2. Create a Drop-down List for End Time_**  
> **_Point 1:_** *End time should always be greater than start time.*  
> **_Point 2:_** *Create a drop-down list without blanks by ignoring cells not showing formula results.*    
- Feature: Data Validation
- Formula: ```COUNTIF``` ```DATE``` ```FIND``` ```IF``` ```IFERROR``` ```INDEX``` ```ISBLANK``` ```ISERROR``` ```LEFT``` ```LEN``` ```MONTH``` ```RIGHT``` ```ROW``` ```OFFSET``` ```SMALL```  
#### **_Step 2.1. Make Sure the List of Time is Correct_**    
> **_Point:_** *There is a possibility that May-14 to be considered as 2018/5/14 by Excel and therefore should be formatted to 2018/5/1 before further processing.*  
```
=IFERROR(DATE("20"&RIGHT(AE1,2),
              MONTH(LEFT(AE1,3)&"-1"),
              1),
         0)
```   
#### **_Step 2.2. Create a List of End Time Greater than the Selected Start Time_** 
```
=IFERROR(OFFSET($AE$1,
                SMALL(IF(($AG:$AG1000-(DATE("20"&RIGHT(StartTime,2),
                                            MONTH(LEFT(StartTime,3)&"-1"),
                                            1)))>0,
                         ROW($AG$1:$AG$1000),
                         ""),
                      ROW(1:1))-1,
                0),
         "")
```   
#### **_Step 2.3. Define a Name for End Time_**
> **_Point 1:_** *Differentiate cells with formula results from those without showing formula results.*   
```
=IF(LEN(AH1)>0,1,0)
```  
> **_Point 2:_** *Define a name for a cell range showing formula results.*      
```
=OFFSET(Data!$AH$1,
        0,0,
        COUNTIF(Data!$AI:$AI,1),1)
```  
#### **_Step 3. Calculate Cumulative Sum_**  
- Formula: ```COLUMN``` ```INDEX``` ```LEFT``` ```LEN``` ```MATCH``` ```RIGHT``` ```SUMPRODUCT```   
#### **_Step 3.1. Cumulative Sum of Current Year_**    
```
=SUMPRODUCT(Data!13:13,
            (Data!$5:$5=H$5)*1,
            (COLUMN(Data!13:13)>=MATCH(B7,Data!$6:$6,0))*1,
            (COLUMN(Data!13:13)<=MATCH(C7,Data!6:6,0)+6)*1
            )
```  
#### **_Step 3.2. Cumulative Sum of Previous Year_**  
```
=SUMPRODUCT(Data!13:13,
            (Data!$5:$5=H$5)*1,
            (COLUMN(Data!13:13)>=MATCH(LEFT(B5,LEN(B5)-2)&(RIGHT(B5,2)-1),Data!6:6,0))*1,
            (COLUMN(Data!13:13)<=MATCH(LEFT(C5,LEN(C5)-2)&(RIGHT(C5,2)-1),Data!6:6,0)+6)*1
            )
```
