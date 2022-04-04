#weeklyScheduler

weeklyScheduler is a JavaScript program that generates an Excel file in the form of a graphical weekly schedule.

It accepts as input a tabular format Excel file with department course data as the rows.

It extracts data from the input file to generate the output file.

Customized weekly schedules and reports for abnormal course data are generated.

##Usage
* Install node.js and import node modules into project directory

* To color code sections by Instructor, run

```
node.js Fall_2022_CS_Reflist_3.1.22.xlsx CS "ASTR PHYS" I
```

* To color code sections by Course (Subject + Catalog), run

```
node.js Fall_2022_CS_Reflist_3.1.22.xlsx "ASTR PHYS" C
```

## Files/Directories
1. Excel_Extra_Input_Files: contains older input files to work with
   
    Note: Input Excel we want to run with must be located in main directory
   

2. Excel_Output: output files are saved here


3. Config files must be located in main directory
##Notes

###Report Worksheets
1. Courses related by auto enroll
    1. Sections in an auto group that are missing the units field, can infer units from 
       a related component.
    2. Main component sections that cannot be matched to their listed auto enroll components
       can still be drawn to the schedule but will be flagged on the report.
    3. Section units are adjusted (if possible) according to the reported number of auto
       enroll sections by the main component, for that group.
       

2. Dual Reports
    
    Courses can be drawn to the schedule but also flagged
   
###Regular Schedule Worksheets
1. Sections with missing fields
   1. Required fields to be drawn 
      
      ```Term, Subject, Catalog, Pat, START TIME, END TIME```
        
    2. Non-required missing fields will be omitted from schedule block text.
       
        ```Component, Section, Last, Facil ID```
    

2. Courses related by auto enroll
   1. Auto enroll components with issues with their main components can still be drawn to the schedule,
      as long as all of their required scheduling info is present.
   2. Main component courses with issues with their auto enroll components can still be drawn (under the same conditions as above).


3. Courses without colors 
   
   Courses that cannot be matched to an Instructor or Course color will be drawn in gray.

###Flexibility
* Matches department/subject command line arguments to correct tab in input Excel file
  

* Order of columns in input  Excel file does not matter
  

* Order of rows/auto enroll group components does not matter
  

* Time is accepted in various formats, including:
    * hh:mm:ss A -> 13:00:01 pm, 1:10:45 aM 
    * hh:mm:ssA  -> 13:00:01pm, 1:10:45aM
    * hh:mm:ss   -> 13:00:01, 1:10:45
    * hh:m A     -> 13:00 pm, 1:10 aM
    * hh:mA      -> 13:00pm, 1:10aM
    * hh:m       -> 13:00, 1:10
    * h A        -> 13 pm, 1 aM
    * hA         -> 13pm, 1aM
    * h          -> 13, 1
    

* Config file names are inferred from CL args

  Example: "ASTR PHYS" ==> astr_phys_config.json
  

* Worksheet title inferred from CL args and "Term" found in input Excel

  Example: "ASTR PHYS" & Term: 2217 = "ASTR PHYS Fall 2021 Schedule"


* Attribute row index is located, not assumed


* Output Excel filename dynamically generated

