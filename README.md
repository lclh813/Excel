# Excel
## Obejective

## Demo 1. Line Chart
### 1.1. Select Data for the Chart
- Formula: INDEX, MATCH
### 1.2. Add an Interactive Vertical Line
- Formula: IF
### 1.3. Dynamically Update Data according to the location of the Vertical Line
- VBA: 
  Class Module: Specify which line and which point to be referred to.
  Module: Specify to which chart the Class Module is to be applied.
- Formula: INDEX, OFFSET  
 
## Demo 2. Bubble Chart
### 2.1. Select Data for the Chart
- Formula: IF, INDEX, MATCH
### 2.2. Add an Interactive Legend
- VBA: 
  Microsoft Excel Object: Determine which legend items has been selected and data that should be displayed on the chart.

## Demo 3. Cumulative Table
### 3.1. Create a Drop-down List for Start Time
- Point: Create a Drop-down List without Blank
- Function: Data Validation
- Formula: FIND, IF, IFERROR, INDEX, ISBLANK, ISERROR, LEN, ROW, SMALL
### 3.2. Create a Drop-down List for End Time
- Point 1: End Time should always be greater than Start Time
- Point 2: Ignore Cells without Formula Results since the Corresponding Data has not yet been Imported
- Function: Data Validation
- Formula: COUNTIF, DATE, FIND, IF, IFERROR, INDEX, ISBLANK, ISERROR, LEFT, LEN, MONTH, RIGHT, ROW, OFFSET, SMALL
### 3.3. Calculate Cumulative Sum
- Formula: COLUMN, INDEX, LEFT, LEN, MATCH, RIGHT, SUMPRODUCT

[![Watch the video](https://img.youtube.com/vi/Youtubeid/hqdefault.jpg)](https://youtu.be/Youtubeid)

[![Demo CountPages alpha](https://j.gifs.com/Youtubeid)]
