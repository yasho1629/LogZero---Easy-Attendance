const xlsx = require('xlsx');
const path = require('path');

const exportExcel = (data, workSheetColumnNames, workSheetName, filePath) => {
    const workBook = xlsx.utils.book_new();
    const workSheetData = [
        workSheetColumnNames,
        ... data
    ];
    const workSheet = xlsx.utils.aoa_to_sheet(workSheetData);
    xlsx.utils.book_append_sheet(workBook, workSheet, workSheetName);
    xlsx.writeFile(workBook, path.resolve(filePath));
}

const exportUsersToExcel = (attendance, workSheetColumnNames, workSheetName, filePath) => {
    const data = attendance.map(attendance => {
        return [attendance.roll_number,attendance.name,attendance.dbms_res,attendance.coa_res,attendance.ads_res,attendance.am_res,attendance.ps_res];
    });
    exportExcel(data, workSheetColumnNames, workSheetName, filePath);
}

const exportMarksToExcel = (marks, workSheetColumnNames, workSheetName, filePath) => {
    const data = attendance.map(attendance => {
        return [marks.roll_number,marks.name,marks.dbms_res,marks.coa_res,marks.ads_res,marks.am_res,marks.ps_res];
    });
    exportExcel(data, workSheetColumnNames, workSheetName, filePath);
}

module.exports = exportUsersToExcel;
