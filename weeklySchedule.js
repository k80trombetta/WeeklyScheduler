let Excel = require('exceljs');
let Moment = require('moment'); //Using this to format time
const fs = require('fs');
const config = require('./config');



// Converts time to string representation of time with am or pm
function timeToString(t) {  // v contains start/end time as a string or a date
    let time = undefined;
    let amPm = null;
    if(typeof t === "string") {
        t = t.toUpperCase();
        let hours = "00";
        let minutes = "00";
        let amPmIdx = t.indexOf("A") != -1 ? t.indexOf("A") : (t.indexOf("P") != -1 ? t.indexOf("P") : -1);
        let colonIdx = t.indexOf(":");

        if (amPmIdx != -1){
            amPm = t.substring(amPmIdx) + (!t.substring(amPmIdx).includes("M") ? "M" : "");
            t = t.substring(0, amPmIdx);
        }
        hours = t;
        if (colonIdx != -1){
            hours = t.substring(0, colonIdx);
            if (colonIdx < t.length - 1){
                minutes = t.substring(colonIdx+1);
                colonIdx = minutes.indexOf(":");
                minutes = colonIdx != -1 ? minutes.substring(0, colonIdx) : minutes.trim();
            }
        }
        amPm = amPm === null ? parseInt(hours) >= 8 && parseInt(hours) < 12 ? "AM" : "PM" : amPm;
        hours = parseInt(hours) > 12 ? (parseInt(hours) - 12).toString() : hours;
        hours = hours.length === 1 ? "0" + hours : hours;
        minutes = minutes.length === 1 ? "0" + minutes : minutes;
        time = hours + ":" + minutes + " " + amPm;
    }
    else if (Date.parse(t)){
        let hours = t.getHours();
        hours = hours > 12 ? hours - 12 : hours;
        hours = Number(t.getUTCHours()).toString().padStart(2, "0");
        let amPm = hours >= 8 && hours < 12 ? "AM" : "PM";
        let minutes = Number( t.getUTCMinutes() ).toString().padStart(2, "0");
        time = hours + ":" + minutes + " " + amPm;
    }
    else
        console.log(t, " -- start/end time is not a string or a date object." );
    return time;
}






// Gets the row data for this section to be drawn in an extra tabs row
function getRowDataForReport(rowData, section, columnHeader) {
    if (columnHeader === "COURSE") rowData.push(`${section["Subject"]}${section["Catalog"]}`);
    else if (columnHeader === "SECTION") rowData.push(section["Section"]);
    else if (columnHeader === "PATTERN") rowData.push(section["Pat"]);
    else if (columnHeader === "UNITS") rowData.push(section["Min Units"]);
    else if (columnHeader === "INSTRUCTOR") {
        if (section["Last"] || section["First Name"])
            rowData.push(section["Last"] && section["First Name"] ? `${section["Last"]}, ${section["First Name"]}` :
                (section["Last"] ? section["Last"] : section["First Name"] ? section["First Name"] : ""));
    }
    else if (columnHeader === "REPORT REASON") rowData.push(section["Report Reason"]);
    else rowData.push(section[columnHeader] === "Invalid date" ? null : section[columnHeader]);
}






// adds the times and M-F headers to the worksheet
function formatColumnHeaders(worksheet, column_values) {
    let row = worksheet.getRow(2);
    row.values = column_values;
    row.font = {name: 'Times New Roman', bold: true};
    row.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    row.border = {bottom: {style: 'medium'}};
}






// Adds the first row value/main header to the worksheet
function formatWorksheetHeader(worksheet, title){
    let row_title = worksheet.getRow(1);
    row_title.values = [title];
    row_title.height = 20;
    row_title.font = {size: 18, name: 'Times New Roman', bold: true};
    row_title.alignment = { vertical: 'middle' };
    row_title.border ={ bottom: {style:'medium'} };
}






function formatScheduleCell(cell, cell_text, color) {
    cell.border = {top: {style: 'medium'}, left: {style: 'medium'}, bottom: {style: 'medium'}, right: {style: 'medium'}};
    cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: color}};
    cell.value = cell_text;
    cell.font = {name: 'Times New Roman', bold: true};
    cell.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
}






// gets text to draw in block
function getCellText(sections){
    let cellText = sections[0]["Subject"] + sections[0]["Catalog"] + '\n';
    sections.some(section => { if (section["Component"]){
        cellText += section["Component"] + ' ';
        return true;
    }});
    let cellSections = [];
    sections.forEach(section => { if (section["Section"]) cellSections.push(section["Section"])});
    cellSections.sort();
    cellText += cellSections.length <= 3 ? cellSections.join(",") + '\n' : cellSections[0] + "-" +  cellSections[cellSections.length - 1] + '\n';
    sections.some(section => { if (section["Last"]){
        cellText += section["Last"] + '\n';
        return true;
    }});
    if (!sections.some(section => { if (section["Facil ID"]){
        cellText += section["Facil ID"] + '\n';
        return true;
    }})) cellText += "TBD\n";
    return cellText;
}






function getColumnLetters(size){
    let integer_letter = 65; // A
    let integer_letter2 = 64;
    let letters = [];
    let doubleLetter = false;
    for(let i = 0; i < size; i++){
        letters.push(doubleLetter ? String.fromCharCode(integer_letter2)+String.fromCharCode(integer_letter) : String.fromCharCode(integer_letter));
        if(++integer_letter >= 91){
            integer_letter2++;
            integer_letter = 65;
            doubleLetter = true;
        }
    }
    return letters;
}






//Just sets up a basic array of timestamps for the excel schedule sheet
function militaryTime(){
    let startTime = new Moment('08:00:00',"h:mm:ss");
    let timeArray = [];
    while(startTime.hours() !== 22){
        timeArray.push(startTime.clone());
        startTime.add(5,"m")
    }
    return timeArray;
}






/*ARGS- 1. schedule
        2. courseTimeGroup
        3. courseTimeGroupKey
DESCRIPTION- */
function addSectionToSchedule(schedule, section){
    let meetingPattern = parseMeetingPattern(section["Pat"]);
    let newStartTime = section["START TIME"];
    let newEndTime = section["END TIME"];
    let sectionGroupKey = `${section["Subject"]}${section["Catalog"]}_${section["Pat"]}_${newStartTime}_${newEndTime}`;

    meetingPattern.forEach((day) => {
        let openColumnId = -1;
        let foundKey = false;// column.push({[section["MetaData"]["CourseTimeId"]]: section});
        let openColumn = !schedule[day].every((column, columnId) => {
            let timeOverlap = Object.entries(column).some(([scheduleBlockKey, scheduledBlock]) => {
                let scheduleStartTime = Object.values(scheduledBlock)[0]["START TIME"];
                let scheduleEndTime = Object.values(scheduledBlock)[0]["END TIME"];
                foundKey = scheduleBlockKey === sectionGroupKey;
                return (foundKey || ((newStartTime >= scheduleStartTime && newStartTime < scheduleEndTime) ||
                    (newEndTime > scheduleStartTime && newEndTime <= scheduleEndTime)));
            });
            if (foundKey) timeOverlap = false;
            openColumnId = !timeOverlap ? columnId : openColumnId;
            return timeOverlap;
        });
        if (openColumn){
            if (foundKey && !schedule[day][openColumnId][sectionGroupKey].filter(sectionb => sectionb["Section"] === section["Section"]).length)
                schedule[day][openColumnId][sectionGroupKey].push(section);
            else
                schedule[day][openColumnId][sectionGroupKey] = [section];
        }
        else{
            schedule[day].push([]);
            schedule[day][schedule[day].length - 1][sectionGroupKey] = [section];
        }
    });
}






// Written by previous programmer
function getRowTimes(rowsForTimes) {
    let times = militaryTime();
    let rowTimes = [];
    let rowIdx = 3;
    times.forEach(time => {
        let formattedTime = time.format("HH:mm:ss");
        rowsForTimes[formattedTime] = rowIdx;
        if (time.minutes() % 15 === 0)
            rowTimes[rowIdx++] = time.format("hh:mm a");
    });
    return rowTimes;
}






// Writes the worksheets for the extra tabs in this workbook. extra tabs are either courses missing details to draw them on the regular schedule
// or courses with questionable units/hour.
function generateReportWorksheets(workbook, reportGroups, colorGrouping) {
    reportGroups.forEach(reportGroup => {
        let columnHeaders = ["COURSE", "SECTION", "PATTERN", "START TIME", "END TIME", "UNITS", "INSTRUCTOR", "REPORT REASON"];
        let worksheet = workbook.addWorksheet(reportGroup[0]["Report Code"] === "Flagged" ? "Flagged Courses Report" : "Unscheduled Courses Report");
        formatWorksheetHeader(worksheet, config.title);
        formatColumnHeaders(worksheet, columnHeaders);

        reportGroup.forEach((section, sectionIdx) => {
            let rowData = [];
            let color = section[(colorGrouping === "I" ? "Instructor Color" : "Course Color")].substring(1);
            columnHeaders.forEach(columnHeader => getRowDataForReport(rowData, section, columnHeader));
            worksheet.getRow(sectionIdx + 3).values = rowData;
            worksheet.getRow(sectionIdx + 3).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: color}};
            worksheet.getRow(sectionIdx + 3).alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
        });
        worksheet.getColumn(columnHeaders.indexOf("SECTION") + 1).width = "SECTION".length + 4;
        worksheet.getColumn(columnHeaders.indexOf("PATTERN") + 1).width = "PATTERN".length + 4;
        worksheet.getColumn(columnHeaders.indexOf("START TIME") + 1).width = "START TIME".length + 4;
        worksheet.getColumn(columnHeaders.indexOf("END TIME") + 1).width = "END TIME".length + 4;
        worksheet.getColumn(columnHeaders.indexOf("INSTRUCTOR") + 1).width = "INSTRUCTOR".length * 3;
        worksheet.getColumn(columnHeaders.indexOf("REPORT REASON") + 1).width = "REPORT REASON".length * 3;
        worksheet.commit();
    });
}






// Writes the a worksheet for which a weekly schedule should be generated, based on a tab in config.js
// Returns that worksheet
function createWorksheetForTab(workbook, tabName, tabSections, colorGrouping){
    let worksheet = workbook.addWorksheet(tabName);
    let columnHeader = ["TIMES"];
    let dayHeaders = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"];
    let schedule = {"M": [[]], "T": [[]], "W": [[]], "TH": [[]], "F": [[]]};
    let columnIdx = 1;
    let letterIdx = 2;
    let rowsForTimes = {};
    let rowTimes = getRowTimes(rowsForTimes);
    let columnLetters = getColumnLetters(78);
    worksheet.getColumn(columnIdx).values = rowTimes;

    Object.values(tabSections).forEach(section => addSectionToSchedule(schedule, section));

    Object.entries(schedule).forEach(([dayKey, daySchedule], dayIdx) => {
        columnIdx += daySchedule.length + 3;
        worksheet.getColumn(columnIdx).values = rowTimes;
        let startingLetter = columnLetters[letterIdx];
        worksheet.getColumn(letterIdx).width = 1;
        columnHeader[letterIdx] = dayHeaders[dayIdx];

        daySchedule.forEach(columnInDay => {
            let letter = columnLetters[letterIdx++];
            worksheet.getColumn(letter).width = 14;
            Object.entries(columnInDay).forEach(([groupKey, sectionGroup]) => {
                let startRow = rowsForTimes[sectionGroup[0]["START TIME"]];
                let endRow = rowsForTimes[sectionGroup[0]["END TIME"]];
                worksheet.mergeCells(`${letter}${startRow}`, `${letter}${endRow - 1}`);
                let cell = worksheet.getCell(`${letter}${startRow}`);
                let cellText = getCellText(sectionGroup);
                let color = sectionGroup[0][(colorGrouping === "I" ? "Instructor Color" : "Course Color")].substring(1);
                formatScheduleCell(cell, cellText, color);
            });
        });

        worksheet.getColumn(letterIdx + 1).width = 1;
        columnHeader[letterIdx + 1] = "TIME";
        worksheet.mergeCells(`${startingLetter}2`, `${columnLetters[letterIdx - 1]}2`);
        letterIdx += 3
    });
    formatWorksheetHeader(worksheet, config.title);
    formatColumnHeaders(worksheet, columnHeader);
    return worksheet;
}






/*ARGS- meetingPattern string - Examples: "M" or "MW"
DESCRIPTION- Parses the given meeting pattern into individual days
RETURN- Returns an array of the individual days of the meeting pattern */
function parseMeetingPattern(meetingPattern){
    return meetingPattern.length === 1 || meetingPattern === "TH" || meetingPattern === "ARR" ? [meetingPattern] :
        meetingPattern === "MW" ? ["M", "W"] : ["T", "TH"];
}






function flagReportSections(autoGroup){
    let missingUnits = false;

    autoGroup.forEach((section, sectionIdx) => {
        if (!section["Subject"] || !section["Catalog"] || !section["Pat"] || section["START TIME"] === "Invalid date" || section["END TIME"] === "Invalid date"){
            section["Report Code"] = "Unscheduled";
            section["Report Reason"] = section["Pat"] === "ARR" ? "Weekly hours by arrangement" : "Missing critical scheduling data";
        }
        if (sectionIdx === 0 && !section["Min Units"])
            missingUnits = true;
    });

    if (autoGroup[0]["Report Code"] === undefined && !missingUnits){
        let base = autoGroup[0];
        let units = base["Min Units"];
        var minutes = Moment(base["END TIME"], "HH:mm").diff(Moment(base["START TIME"], "HH:mm"), 'minutes');
        let numDays = parseMeetingPattern(base["Pat"]).length;
        let numAutoEnrolls = autoGroup[0]["Auto Enrol"] && autoGroup[0]["Auto Enr 2"] ? 2 : (autoGroup[0]["Auto Enrol"] ? 1 : 0);
        if ((Math.abs(((units - numAutoEnrolls) * 50) - (numDays * minutes)) > 30)
            && !autoGroup[0]["Catalog"].includes("W") && !autoGroup[0]["Component"].includes("LAB")){
            autoGroup.forEach(section => {
                section["Report Code"] = "Flagged";
                section["Report Reason"] = "Units/hours discrepancy";
            })} // ((actual units - numAutoEnrolls)*(50 min/1 unit)) - (numDaysMeeting * minutes per meeting day) > 30 minutes is flagged
    }

    if (autoGroup[0]["Report Code"] === "Unscheduled")
        autoGroup.slice(1).filter(auto => auto["Report Code"] === undefined).forEach(autob => {
            autob["Report Code"] = "Flagged";
            autob["Report Reason"] = "Main component missing data";
        });
}






/*ARGS- 1. sections: Object that contains all course sections, grouped by their CourseTimeIds
        2. tabs:
DESCRIPTION- Adds the section objects' grouping ids that help us associate related/combined/autoEnroll sections with each other */
function addTabIdsToSections(sections, tabs){
    let tabSections = null;

    tabs.forEach((tab, tabId) => {
        let filter = tab["entries"][0];
        if (filter === "all")
            tabSections = Object.values(sections);
        else if (filter === "all except")
            tabSections = Object.values(sections).filter(section => !tab["entries"].slice(1).includes(section["Subject"] + section["Catalog"]));
        else if (filter !== "rest")
            tabSections = Object.values(sections).filter(section => tab["entries"].includes(section["Subject"] + section["Catalog"]));
        else if (tabSections !== null) {
            tabSections = Object.values(sections).filter(section => Object.values(tabSections).every(tabCourse => {
                return `${tabCourse["Subject"]}${tabCourse["Catalog"]}` !== `${section["Subject"]}${section["Catalog"]}`
            }));
        }

        tabSections.forEach(section => {
            if (section["Report Code"] === undefined || section["Report Code"] === "Flagged"){
                if (section["Tabs"] === undefined)
                    section["Tabs"] = [tabId];
                else
                    section["Tabs"].push(tabId);
            }
        });
    });
}






function groupSections(subjectSections){
    let groupIds = {};
    let groupId = 0;
    let autoId = 0;

    Object.values(subjectSections).forEach(section => {
        if (section["Group ID"] === undefined){
            let sectionGroupId = `${section["Subject"]}${section["Catalog"]}_${section["Pat"]}_${section["START TIME"]}_${section["END TIME"]}`;
            if (groupIds[sectionGroupId] === undefined)
                groupIds[sectionGroupId] = groupId++;
            section["Group ID"] = groupIds[sectionGroupId];
            section["Auto ID"] = autoId++;

            let autos = Object.values(subjectSections).filter(auto => (section["Auto Enrol"] && (`${auto["Subject"]}${auto["Catalog"]}_${auto["Section"]}` ===
                `${section["Subject"]}${section["Catalog"]}_${section["Auto Enrol"]}`)) || (section["Auto Enr 2"] && (`${auto["Subject"]}${auto["Catalog"]}_${auto["Section"]}`
                === `${section["Subject"]}${section["Catalog"]}_${section["Auto Enr 2"]}`)));

            Object.values(autos).forEach(auto  => {
                auto["Group ID"] = section["Group ID"];
                auto["Auto ID"] = section["Auto ID"];
            });
            flagReportSections([section].concat(autos));
        }
    });
}






/*ARGS- 1. row: Excel Row Object: Contains details for a class section.
        2. attributeCellIds: Object of attribute/cell id pairs
DESCRIPTION- Creates a section object to hold data extracted from the excel row and metadata for that data
RETURN- Array containing section details from the given row and columns specified by the Keys in attributeCellIds */
function getSectionDataFromRow(row, attributeCellIds){
    let section = {};

    for (let attribute in attributeCellIds){
        let cellValue = row.getCell(attributeCellIds[attribute]).value;
        if (attribute === "START TIME" || attribute === "END TIME")
            cellValue = new Moment(timeToString(cellValue) , "hh:mm A").format("HH:mm:ss");
        section[attribute] = cellValue;
    }

    let instructor = `${section["First Name"]} ${section["Last"]}`;
    let course = `${section["Subject"]}${section["Catalog"]}`;
    section["Instructor Color"] = instructor in config.instructors ? config.instructors[instructor].color : "#DDDDDD";
    section["Course Color"] = course in config.course_colors ? config.course_colors[course] : "#DDDDDD";
    return section;
}






/*ARGS- tab: Excel Worksheet (tab)
DESCRIPTION- Creates an object with pairs of course attributes and their cell ids found in the given tab
RETURN- Object of pairs whose keys are course attributes and values are the cell ids (columns) at which those attributes occur */
function getAttributeCellIds(worksheet, attributeRow){
    let attributeCellIds = {"Subject":null, "Catalog":null, "Pat":null, "START TIME":null, "END TIME":null,
        "Section":null, "Auto Enrol":null, "Auto Enr 2":null, "Component":null, "Last":null, "First Name":null, "Facil ID":null, "Min Units":null};

    for (let attributeCellIdsKey in attributeCellIds) {
        worksheet.getRow(attributeRow).values.some((cellValue, cellIdx) => {
            if (cellValue === attributeCellIdsKey){
                attributeCellIds[attributeCellIdsKey] = cellIdx;
                return true;
            }
        });
    }
    return attributeCellIds;
}






/*ARGS- 1. workbook: Excel Workbook Object
        2. subject: Array of string representation of a subject passed in from the CL. -Examples: ["CS"] or ["ASTR", "PHYS"]
DESCRIPTION- Looks through the sheets in the given workbook for a sheet that contains the given subject
RETURN- Number: The index of the first Worksheet (tab) that has the given subject. */
function getSubjectWorksheetIds(workbook, subject) {
    let subjectWorksheetIds = [];

    if (workbook.worksheets.length === 1){ // Search for the given subject in the rows of this tab
        let worksheet = workbook.getWorksheet(1);
        subject.some(subSubject => {
            return worksheet.getRows(1, worksheet.rowCount).some( (row, rowIdx) => {
                if (subSubject === row.getCell(2).value) {
                    subjectWorksheetIds.push(1);
                    return true;
                }
            });
        });
    }
    else { // Search for the given subject in the tab names of this workbook or vice versa
        subject.forEach(subSubject => {
            workbook.worksheets.some((worksheet, worksheetIdx) => {
                if ((subSubject.includes(worksheet.name) || worksheet.name.includes(subSubject)) && !subjectWorksheetIds.includes(worksheetIdx+1)){
                    subjectWorksheetIds.push(worksheetIdx+1);
                    return true;
                }
            });
        });
    }
    return subjectWorksheetIds;
}






/*For each tab in config.js, generates a worksheet with a regular schedule, as well as any extra tabs needed based on
course data. Extra tabs would include the reports on classes that either 1. were missing info to put them on a weekly
schedule, or 2. had questionable units/hours */
function generateWorkbook(subject, sections, colorGrouping){
    let workbookOptions = {
        filename: `weeklySchedule_${subject}_${Moment().format("MM_DD_YYYY_hh_mm")}.xlsx`,
        useStyles: true, useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(workbookOptions);

    Object.entries(config.tabs).forEach( ([tabKey, tab], tabIdx) => {
        let tabSections = Object.values(sections).filter(section => section["Tabs"] !== undefined && section["Tabs"].includes(tabIdx));
        let tabWorksheet = createWorksheetForTab(workbook, tab.name, tabSections, colorGrouping);
        tabWorksheet.commit();
    });
    let unscheduledSections = Object.values(sections).filter(section => section["Report Code"] !== undefined && section["Report Code"] === "Unscheduled");
    let flaggedSections = Object.values(sections).filter(section => section["Report Code"] !== undefined && section["Report Code"] === "Flagged");
    generateReportWorksheets(workbook, [unscheduledSections, flaggedSections].filter(reportSections => reportSections.length), colorGrouping);
    workbook.commit().then(function() {});
}






function processData(departmentData){
    Object.values(departmentData).forEach(subjectSections => {
        groupSections(subjectSections);
        addTabIdsToSections(subjectSections, config.tabs);
    });
}






/* This function takes a file name of a JSON in the current directory (no ./ is need in the name) and reads from it.
       The file must be an array of flat JSON objects that holds information needed to generate a weekly schedule.
       NOTE auto_enrol does not exist in JSON form yet. Returns another array of flat objects that has only the info needed for weekly schedule
       Currently has no way to add color by class, needs auto_enroll in the JSON */
async function reduceJSON(filename, reducedJSON) {
    let name_keys = ["subject", "catalog", "component", "section", "instructor_lName", "instructor_fName", "facility_name",
        "meeting_pattern", "start_time", "end_time"];

    let fullJSON = require( "./" + filename);
    var text = JSON.parse(fullJSON);
    fullJSON.forEach( obj => {
        if(obj.meeting_pattern !== "ARR"){
            let reduced_item = {};
            name_keys.forEach( name_key => reduced_item[name_key] = obj[name_key]);
            reducedJSON.push(reduced_item);
        }
    });
}






/*ARGS- 1. filename: The string representation of an Excel file name passed in from the CL
        2. subjects: Array of subject(s) arrays passed in from the CL. Example: [["CS"], ["ASTR", "PHYS"]]
        3. subjectsData: Array to be filled with each subject's section data object.
DESCRIPTION- Adds each subject's section data to an object and adds that object to subjectsData
RETURN- An object that contains an object for each subject's section data */
function extractData(filename, subjects, subjectData){
    let workbook = new Excel.Workbook();
    let requiredAttributes = ["Subject", "Catalog", "Pat", "START TIME", "END TIME"];

    return workbook.xlsx.readFile(filename).then( function() {
        subjects.forEach( (subject) => {
            let subjectWorksheetIds = getSubjectWorksheetIds(workbook, subject);
            subject = subject.join(" ");

            subjectWorksheetIds.forEach(worksheetId => {
                let worksheet = workbook.getWorksheet(worksheetId);
                let attributeRow = null;
                worksheet.getRows(1, worksheet.rowCount).some((row, rowIdx) => {
                    if (row.values.includes("Subject")) {attributeRow = rowIdx + 1;return true;}
                });
                let attributeCellIds = getAttributeCellIds(worksheet, attributeRow);
                if (attributeRow && requiredAttributes.every(attribute => attributeCellIds[attribute])){ //sheet has all required attributes
                    subjectData[subject] = [];
                    worksheet.eachRow((row, rowId) => { if (rowId > attributeRow) {
                        let section = getSectionDataFromRow(row, attributeCellIds);
                        subjectData[subject].push(section);
                    }});
                }
            });
        });
    });
}






/*ARGS- 1. argv: Array of string args passed on CL
                 Example: SSU_SCHED.xlsx config.json CS "ASTR PHYS" I = ["SSU_SCHED.xlsx", "config.json", "CS", "ASTR PHYS", "I"]
        2. argc: Number of args in argv
DESCRIPTION- Parses the CL args, validating the 1st is a filename of correct type (xlsx or json).
             For each subject on CL, creates a schedule object containing course data extracted from input file
             For each schedule object, returns a visually pleasing Weekly Schedule as an Excel file */
async function readAndExecute(argv, argc){
    let file = argv[0];
    let fileExt = file.substring(file.lastIndexOf('.') + 1);
    let congigFile = argv[1];
    let departments = [];
    for (let i = 2; i < argc - 1; i++)
        departments.push(argv[i].split(' ')); // array of arrays. CLI args CS and "ASTR PHYS" will be [["CS"],["ASTR", "PHYS"]]
    const colorGrouping = argv[argc - 1];

    let departmentData = [];
    if(fileExt === "xlsx")
        await extractData(file, departments, departmentData).catch(err => console.log(err));
    else if(fileExt === "json")
        await reduceJSON(file, departmentData);
    else
        console.log(file + " must end with xlsx or json");

    processData(departmentData);

    if(fileExt === "xlsx" || fileExt === "json"){
        Object.entries(departmentData).forEach(([subjectKey, subjectData]) => {
            generateWorkbook(subjectKey, subjectData, colorGrouping);
        });
    }
}






async function main(){
    const argv = process.argv.slice(2);
    await readAndExecute(argv, argv.length).catch(err => console.log(err));
}






main();
