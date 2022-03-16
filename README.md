# WeeklyScheduler
Node project extracts and processes course data from an input Excel file, and produces a new Excel file containing a visually pleasing weekly schedule and report

# Installation
Install node in project directory
https://nodejs.org/en/download/

# Current Work
I am adding in more flexibility into the weeklySchedule.js file so it is more flexible and does not break so easily. Previous versions of the project expected the input file to come in a specific format. The flexibility I am adding anticipates various configurations of the input file.

# Changes from previous versions
1. Searching for attribute row instead of assuming its location

2. Seaching for specific required attributes needed for processing, instead of assuming their presence.

3. Using attribute indices to access specific columns instead of assuming hard coded locations.

4. Assuming sections may be unordered.

5. Section START TIME and END TIME values may be formatted as a Date object or String that can be parsed as a valid time.

6. Anticipation of invalid/missing cells. If cells necessary to schedule a course are missing, that course will be omitted from a regular schedule tab in the output file, and identified in a separate report tab in the output file.

7. Anticipation of units/hours discrepancies. Calculations are executed to determine if course units/hour are reasonable. If no,t the section is identified in a report tab while still being drawn on the regular schedule.

8. Project accepts more arguments from the command line instead of being hard coded. Ultimately, all data needed for the program to run will come in on the command line. No files will be saved in the project directory (config.json, input file, etc.).

9. Project can handle input file containing multiple tabs, as well as multiple subjects/departments coming in on the command line. It can locate the command line subject(s) corresponding tab(s).

10. It will still draw sections with missing info, as long as that info is not necessary to put on the schedule.
