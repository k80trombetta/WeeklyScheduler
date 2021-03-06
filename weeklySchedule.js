let Excel = require('exceljs');
let Moment = require('moment'); //Using this to format time
// const fs = require('fs');



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
        amPm = amPm === null ? (parseInt(hours) >= 8 && parseInt(hours) < 12 ? "AM" : "PM") : amPm;
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






/*ARGS- 1. rowData: array we are populating with section info
        2. section: section to extract data from for this report
        3. columnHeaders: array of column headers for report
DESCRIPTION- Gets the row data for this section to be drawn in an extra tabs row */
function getRowDataForReport(rowData, section, columnHeaders) {
    columnHeaders.forEach(columnHeader => {
        if (columnHeader === "SUBJECT") rowData.push(section["Subject"] ? section["Subject"] : "-");
        else if (columnHeader === "CATALOG") rowData.push(section["Catalog"] ? section["Catalog"] : "-");
        else if (columnHeader === "SECTION") rowData.push(section["Section"] ? section["Section"] : "-");
        else if (columnHeader === "PATTERN") rowData.push(section["Pat"] ? section["Pat"] : "-");
        else if (columnHeader === "UNITS") rowData.push(section["Min Units"] ? section["Min Units"] : "-");
        else if (columnHeader === "INSTRUCTOR") {
            if (section["Last"] || section["First Name"])
                rowData.push(section["Last"] && section["First Name"] ? `${section["Last"]}, ${section["First Name"]}` :
                    (section["Last"] ? section["Last"] : section["First Name"] ? section["First Name"] : ""));
            else
                rowData.push("-");
        }
        else if (columnHeader === "REPORT REASON"){
            section["Report Reason"].forEach((reason, reasonIdx) => {
                rowData.push( (reasonIdx < section["Report Reason"].length - 1) ? `${reason}\n` : reason);
            });
        }
        else rowData.push(!section[columnHeader] || section[columnHeader] === "Invalid date" ? "-" : Moment(section[columnHeader] , "HH:mm:ss").format("hh:mm A"));
    });
}






/*ARGS- 1. worksheet: worksheet we are generating
        2. column_values: Time and M-F column header values
DESCRIPTION- adds the Time and M-F headers to the worksheet */
function formatColumnHeaders(worksheet, column_values) {
    let row = worksheet.getRow(2);
    row.values = column_values;
    row.font = {name: 'Times New Roman', bold: true};
    row.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    row.border = {bottom: {style: 'medium'}};
}






/*ARGS- 1. worksheet: worksheet we are generating
        2. title: string title of the worksheet
DESCRIPTION- Adds the first row value/main header to the worksheet */
function formatWorksheetHeader(worksheet, title){
    let row_title = worksheet.getRow(1);
    row_title.values = [title];
    row_title.height = 20;
    row_title.font = {size: 18, name: 'Times New Roman', bold: true};
    row_title.alignment = { vertical: 'middle' };
    row_title.border ={ bottom: {style:'medium'} };
}






/*ARGS- 1. cell: the cell to populate in
        2. cell_text: the text to draw in the cell
        3. color: the color to fill the cell with
DESCRIPTION- formats the given cell */
function formatScheduleCell(cell, cell_text, color) {
    cell.border = {top: {style: 'medium'}, left: {style: 'medium'}, bottom: {style: 'medium'}, right: {style: 'medium'}};
    cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: color}};
    cell.value = cell_text;
    cell.font = {name: 'Times New Roman', bold: true};
    cell.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
}






/*ARGS- 1. sections: list of sections that will be represented in this cell
DESCRIPTION- Gets the contents that will be drawn in a cell. Typically the Subject, Catalog, Component, Section(s),
             Instructor Last Name, and the Facility
RETURN - String containing the text to be drawn in this cell */
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






// returns list of possible letter headers to be used in columns of the output excel file
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






/*ARGS- 1. schedule: weekly schedule object
        2. section: the section to be added to the schedule
DESCRIPTION- Adds the given section to the weekly schedule object. If a section with the same subject/catalog/pat/time is already on the schedule,
             this section's section number is added to that block's list of sectio numbers that will be included in the block. If a section
             with  conflicting time is already on the schedule, we will add this section to another column on the day where there is no
             conflicting time */
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
            if (!foundKey || schedule[day][openColumnId][sectionGroupKey].filter(sectionb => sectionb["Section"] === section["Section"]).length)
                schedule[day][openColumnId][sectionGroupKey] = [];
            schedule[day][openColumnId][sectionGroupKey].push(section);
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






/*ARGS- 1. workbook: the workbook we are generating
        2. reportGroups: array of sections that will be reported
        3. title: title associated with this department's schedule
        4. colorGrouping: color grouping indication (instructor or course), given by user on the CL
DESCRIPTION- Writes the worksheets for the reports in this workbook. Reported courses are either courses missing details
             to draw them on the regular schedule or courses with questionable units/hour */
function generateReportWorksheets(workbook, reportGroups, title, colorGrouping) {
    reportGroups.forEach(reportGroup => {
        let columnHeaders = ["SUBJECT", "CATALOG", "SECTION", "PATTERN", "START TIME", "END TIME", "UNITS", "INSTRUCTOR", "REPORT REASON"];
        let worksheet = workbook.addWorksheet(`${Object.keys(reportGroup)[0]} Courses Report`);
        formatWorksheetHeader(worksheet, title + " Schedule");
        formatColumnHeaders(worksheet, columnHeaders);

        Object.values(reportGroup[Object.keys(reportGroup)[0]]).forEach((section, sectionIdx) => {
            let rowData = [];
            let color = section[(colorGrouping === "I" ? "Instructor Color" : "Course Color")].substring(1);
            getRowDataForReport(rowData, section, columnHeaders);
            let row = worksheet.getRow(sectionIdx + 3);
            row.values = rowData;
            row.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: color}};
            row.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
        });

        columnHeaders.forEach((header, h) =>
            worksheet.getColumn(h+1).width = h < columnHeaders.length - 2 ? header.length + 4 : header.length * 3
        );
        worksheet.commit();
    });
}






/*ARGS- 1. workbook: the workbook we are generating
        2. tabName: name of tab from config file
        2. tabSections: array of sections to be drawn on this tab's worksheet
        3. title: title associated with this department's schedule
        4. colorGrouping: color grouping indication (instructor or course), given by user on the CL
DESCRIPTION- Writes the a worksheet for which a weekly schedule should be generated, based on tab in config file
RETURN- Worksheet made for the given tab */
function createWorksheetForTab(workbook, tabName, tabSections, title, colorGrouping){
    let worksheet = workbook.addWorksheet(tabName);
    let columnHeader = ["TIME"];
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
    formatWorksheetHeader(worksheet, title + " Schedule");
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






/*ARGS- autoGroup - array of sections associated by auto enrollment
DESCRIPTION- adds report codes and reasons to sections that can't be scheduled or are questionable */
function flagReportSections(autoGroup){
    let required = ["Subject", "Catalog", "Pat", "START TIME", "END TIME"];
    const csNumbers = require("./CSU_CourseClassification.json");
    let numAutosListed = [autoGroup[0]["Auto Enrol"], autoGroup[0]["Auto Enr 2"]].filter(actualAuto =>  actualAuto).length;

    let totalUnits = null;
    autoGroup.some(section => { if (section["Min Units"]){
        totalUnits = section["Min Units"];
        return true;
    }});

    autoGroup.forEach((section, sectionIdx) => {
        let reportCodes = [];
        let reportReasons = [];
        let cannotDraw = required.some(attribute => !section[attribute] || section[attribute] === "Invalid date");
        if (cannotDraw) {
            reportCodes.push("Unscheduled");
            let arrError = section["Subject"] && section["Catalog"] && section["Pat"] && section["Pat"] === "ARR";
            reportReasons.push(arrError ? "Weekly hours by arrangement" : "Missing critical scheduling data");
        }
        else {
            section["Min Units"] = totalUnits ? (sectionIdx === 0 ? totalUnits - numAutosListed : (totalUnits - (totalUnits - numAutosListed)) / numAutosListed) : null;
            if (!section["Min Units"] || !section["CS Number"]){
                reportCodes.push("Flagged");
                reportReasons.push("Units/hour could not be verified");
            }
            else{
                let csUnitsPerHour = csNumbers[section["CS Number"]];
                if (!isNaN(csUnitsPerHour)){
                    let numDays = parseMeetingPattern(section["Pat"]).length;
                    let time = Moment(section["END TIME"], "HH:mm").diff(Moment(section["START TIME"], "HH:mm"), 'minutes');
                    let marginOfErrorMinutes = numDays === 1 ? 5 : (numDays === 2 ? 10 : 0)
                    if (Math.abs((numDays * (time + 10)) / (section["Min Units"]) - (parseInt(csUnitsPerHour) * 60)) > marginOfErrorMinutes){
                        reportCodes.push("Flagged");
                        reportReasons.push("Units/hour discrepancy");
                    }
                }
            }
            if (sectionIdx === 0){
                let numAutosMatched = autoGroup.slice(1).length;
                if (numAutosListed - numAutosMatched){
                    reportCodes.push("Flagged");
                    reportReasons.push("Missing auto enroll component");
                }
            }
        }
        if (reportCodes.length){
            section["Report Code"] = reportCodes;
            section["Report Reason"] = reportReasons
        }
    });
}






/*ARGS- 1. sections: array containing section data extracted from the input excel file
        2. tabs: tabs found in this department's corresponding config file
DESCRIPTION- Also adds tab id's to each section based on the config file's tab object */
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
            if (section["Report Code"] === undefined || !section["Report Code"].includes("Unscheduled")){
                if (section["Tabs"] === undefined)
                    section["Tabs"] = [];
                section["Tabs"].push(tabId);
            }
        });
    });
}






/*ARGS- sections: array containing, for each department passed in on CL, section data extracted from the input excel file
DESCRIPTION- Adds unique grouping id's to sections matching on subject/catalog/pattern/time and their corresponding auto enroll sections,
             and a second unique grouping id, matching main sections with their auto enroll section(s) */
function groupSections(sections){
    let groupIds = {};
    let groupId = 0;
    let autoId = 0;

    Object.values(sections).forEach(section => {
        if (section["Group ID"] === undefined){
            let sectionGroupId = `${section["Subject"]}${section["Catalog"]}_${section["Pat"]}_${section["START TIME"]}_${section["END TIME"]}`;
            if (groupIds[sectionGroupId] === undefined)
                groupIds[sectionGroupId] = groupId++;
            section["Group ID"] = groupIds[sectionGroupId];
            section["Auto ID"] = autoId++;

            let autos = Object.values(sections).filter(auto => (section["Auto Enrol"] && (`${auto["Subject"]}${auto["Catalog"]}_${auto["Section"]}` ===
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
        3. config: config file associated with the current subject
DESCRIPTION- Creates a section object to hold data extracted from the excel row and metadata for that data
RETURN- Array containing section details from the given row and columns specified by the Keys in attributeCellIds */
function getSectionDataFromRow(row, attributeCellIds, config){
    let section = {};
    let allStrikethrough = true;

    for (let attribute in attributeCellIds){
        let cell = row.getCell(attributeCellIds[attribute]);
        let cellValue = cell.value;
        let strikethrough = cell.style.font.strike;

        if (attribute === "START TIME" || attribute === "END TIME")
            cellValue = new Moment(timeToString(cellValue) , "hh:mm A").format("HH:mm:ss");

        section[attribute] = strikethrough ? null : cellValue;

        if (allStrikethrough && !strikethrough)
            allStrikethrough = false;
    }

    let instructor = `${section["First Name"]} ${section["Last"]}`;
    let course = `${section["Subject"]}${section["Catalog"]}`;
    section["Instructor Color"] = instructor in config.instructors ? config.instructors[instructor].color : "#DDDDDD";
    section["Course Color"] = course in config.course_colors ? config.course_colors[course] : "#DDDDDD";

    return allStrikethrough ? {} : section;
}






/*ARGS- 1. worksheet: Excel Worksheet
        2. attributeRow: index of the row we are extracting section data from
DESCRIPTION- Creates an object with pairs of course attributes and their cell ids found in the given worksheet
RETURN- Object of pairs whose keys are course attributes and values are the cell ids (columns) at which those attributes occur */
function getAttributeCellIds(worksheet, attributeRow){
    let attributeCellIds = {"Term": null, "Subject":null, "Catalog":null, "Pat":null, "START TIME":null, "END TIME":null,
        "Section":null, "Auto Enrol":null, "Auto Enr 2":null, "Component":null, "Last":null, "First Name":null,
        "Facil ID":null, "Min Units":null, "CS Number":null};

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
function getDeptWorksheetIds(workbook, subject) {
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






/*ARGS- 1. title: title associated with this department
        2. sections: sections of this department
        3. colorGrouping: id indicating how to color coordinate sections
        4. tabs: from this department's config file
DESCRIPTION- Generates a workbook for a department containing worksheets for each tab in the config file
             and report worksheets for unschedulable or questionable sections */
function generateWorkbook(title, sections, colorGrouping, tabs){
    let workbookOptions = {
        filename: `./Excel_Output/weeklySchedule_${title.slice(0,-2)}_${Moment().format("MM_DD_YYYY_hh_mm")}.xlsx`,
        useStyles: true, useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(workbookOptions);

    Object.entries(tabs).forEach( ([tabKey, tab], t) => {
        let tabSections = Object.values(sections).filter(section => section["Tabs"] !== undefined && section["Tabs"].includes(t));
        let tabWorksheet = createWorksheetForTab(workbook, tab.name, tabSections, title, colorGrouping);
        tabWorksheet.commit();
    });
    let unscheduledSections = {"Unscheduled": Object.values(sections).filter(section => section["Report Code"] !== undefined && section["Report Code"].includes("Unscheduled"))};
    let flaggedSections = {"Flagged": Object.values(sections).filter(section => section["Report Code"] !== undefined && section["Report Code"].includes("Flagged"))};
    generateReportWorksheets(workbook, [unscheduledSections, flaggedSections].filter(report => Object.values(report[Object.keys(report)[0]]).length), title, colorGrouping);
    workbook.commit().then(function() {});
}






/*ARGS- 1. departmentData: array containing, for each department passed in on CL, section data extracted from the input excel file
        2. configFiles: array of config files, each corresponding to a subject passed in on the CL
DESCRIPTION- For each department, groups their sections for the schedule and adds tab ids to sections */
function processData(departmentData, configFiles){
    Object.values(departmentData).forEach((departmentSections, d) => {
        groupSections(departmentSections);
        addTabIdsToSections(departmentSections, configFiles[d].tabs);
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






/*ARGS- 1. worksheet: ExcelJS Worksheet
        2. department: department we are creating worksheet for
        3. attributeRow: row we are getting term info from
        4. termCellId: column of "Term" values
DESCRIPTION- Parses the term into semester and year
RETURN- Returns string containing department semester year */
function parseTerm(worksheet, department, attributeRow, termCellId){
    let title = "";
    worksheet.getRows(attributeRow + 1, worksheet.rowCount).some( row => {
        let term = row.getCell(termCellId).value;
        if (term) {
            let sem = term.slice(3);
            sem = sem === "1" ? "Winter" : (sem === "3" ? "Spring" : (sem === "5" ? "Summer" : (sem === "7" ? "Fall" : (""))));
            title += " " + sem + " " + term[0] + "0" + term.slice(1, 3);
            return true;
        }
    });
    return department + title;
}






/*ARGS- 1. filename: The string representation of an Excel file name passed in from the CL
        2. departments: Array of department subjects passed in from the CL. Example: [["CS"], ["ASTR", "PHYS"]]
        3. departmentData: Array to be filled with each department's sections
        4. configFiles: array of config files associated with given departments
DESCRIPTION- Adds each departments's section data to an object and adds that object to subjectsData
RETURN- An object that contains an object for each department's section data */
function extractData(file, departments, departmentData, configFiles){
    let workbook = new Excel.Workbook();

    return workbook.xlsx.readFile(file).then( function() {
        let requiredAttributes = ["Term", "Subject", "Catalog", "Pat", "START TIME", "END TIME"];

        departments.forEach( (department, d) => {
            let deptWorksheetIds = getDeptWorksheetIds(workbook, department);
            department = department.join(" ");

            deptWorksheetIds.forEach(worksheetId => {
                let worksheet = workbook.getWorksheet(worksheetId);
                let attributeRow = null;
                worksheet.getRows(1, worksheet.rowCount).some((row, r) => {
                    if (row.values.includes("Term")) {attributeRow = r + 1; return true;}
                });
                let attributeCellIds = getAttributeCellIds(worksheet, attributeRow);
                if (attributeRow && requiredAttributes.every(attribute => attributeCellIds[attribute])){ //sheet has all required attributes
                    let title = parseTerm(worksheet, department, attributeRow, attributeCellIds["Term"]);
                    departmentData[title] = [];
                    worksheet.getRows(attributeRow+1, worksheet.rowCount-attributeRow-1).forEach( row => {
                        let section = getSectionDataFromRow(row, attributeCellIds, configFiles[d]);
                        if (Object.values(section).length)
                            departmentData[title].push(section);
                    });
                }
            });
        });
    });
}






/*ARGS- 1. argv: Array of string args passed on CL
        2. argc: Number of args in argv
DESCRIPTION- Parses the CL args, validating the 1st is a filename of correct type (xlsx or json).
             For each subject on CL, creates a schedule object containing course data extracted from input file
             For each schedule object, returns a visually pleasing Weekly Schedule as an Excel file */
async function readAndExecute(argv, argc){
    let file = argv[0];
    let fileExt = file.substring(file.lastIndexOf('.') + 1);
    let configFiles = [];
    // let configFiles = argv.filter(arg => arg.includes("config"));
    // configFiles.forEach((c,i) => configFiles[i] = require(`./${c}`));
    // let subjectArgsStart = configFiles.length + 1;
    let subjectArgsStart = 1;
    let departments = [];
    for (let i = subjectArgsStart; i < argc - 1; i++){
        let config = require(`./${argv[i].replace(" ","_").toLowerCase()}_config.json`);
        configFiles.push(config);
        departments.push(argv[i].split(' ')); // array of arrays. CLI args CS and "ASTR PHYS" will be [["CS"],["ASTR", "PHYS"]]
    }
    const colorGrouping = argv[argc - 1];

    let departmentData = [];
    if(fileExt === "xlsx")
        await extractData(file, departments, departmentData, configFiles).catch(err => console.log(err));
    else if(fileExt === "json")
        await reduceJSON(file, departmentData);
    else
        console.log(file + " must end with xlsx or json");

    processData(departmentData, configFiles);

    if(fileExt === "xlsx" || fileExt === "json"){
        Object.entries(departmentData).forEach(([title, sections], i) => {
            generateWorkbook(title, sections, colorGrouping, configFiles[i].tabs);
        });
    }
}






async function main(){
    const argv = process.argv.slice(2);
    await readAndExecute(argv, argv.length).catch(err => console.log(err));
}






main();
