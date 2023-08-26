const XLSX = require('xlsx');

let VNPT_IGNORED_OF_NAME_FINDING = ['Họ và tên', 'Giỏi', 'Khá', 'Trung bình', 'Yếu', 'Kém', 'Đạt', 
'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'Độc lập - Tự do - Hạnh phúc', ''];

function getWorkbook(filePath) {
        return XLSX.readFile(filePath);
}

function getWorksheet(workbook, indexOfSheet) {
        let sheetName = workbook.SheetNames[indexOfSheet];
        return workbook.Sheets[sheetName];
}

function vnptGetFullNameAndCellIndexOfStudents (sheet) {
        let names = [];
        let fullname = "";
        for (let cellAddress in sheet) {
                if (!sheet.hasOwnProperty(cellAddress)) {
                        continue;
                } 
                let cell = sheet[cellAddress];
                let { c, r } = XLSX.utils.decode_cell(cellAddress);
                let cellValue = cell.v;
                if (VNPT_IGNORED_OF_NAME_FINDING.includes(cellValue)) {
                        continue;
                }
                if (c === 2) {
                        fullname += cellValue;
                } 
                if (c === 3) {
                        fullname += " " + cellValue;
                        names.push({
                                name: fullname,
                                row: r,
                                column: c
                        });
                        fullname = "";
                }
        }
        return names;
}

function viettelGetStudentHKByName(viettelWorkSheetJsonData, studentFullName) {
        for (let jsonObject of viettelWorkSheetJsonData) {
                if (jsonObject.hasOwnProperty('__EMPTY') && jsonObject.__EMPTY === studentFullName) {
                        return {
                                HK1: jsonObject.__EMPTY_7,
                                HK2: jsonObject.__EMPTY_8,
                                CN: jsonObject.__EMPTY_9
                        };
                }
        };
        return null;
}

function updateCellsOfWorkSheet(vnptFileName, workbook, sheetIndex, students) {
        let worksheetName = workbook.SheetNames[sheetIndex];
        let sheet = workbook.Sheets[worksheetName];

        for (let student of students) {
                let targetRowIndex = student.row;
                let targetColumnIndex = student.column;
                let newValue = student.value;
                if (newValue === undefined) {
                        newValue = '';
                }
                let cellAddress = XLSX.utils.encode_cell({ c: targetColumnIndex, r: targetRowIndex });
                XLSX.utils.sheet_add_aoa(sheet, [[newValue]], {origin: cellAddress});
        }
        XLSX.writeFile(workbook, vnptFileName);
        console.log(`Finished Copying sheet: ${worksheetName}`);
}

function copyViettelSheetToVnptSheet(viettelWorkbook, indexOfViettelSheet, vnptWorkbook, indexOfVnptSheet, vnptFileName, hk) {
        
        let viettelWorksheet = getWorksheet(viettelWorkbook, indexOfViettelSheet);
        let viettelWorksheetJsonData = XLSX.utils.sheet_to_json(viettelWorksheet);

        let vnptWorksheet = getWorksheet(vnptWorkbook, indexOfVnptSheet);
        let studentInfos = vnptGetFullNameAndCellIndexOfStudents(vnptWorksheet);

        let data = [];
        for (let info of studentInfos) {
                let score = viettelGetStudentHKByName(viettelWorksheetJsonData, info.name);
                if (score === undefined || score === null) {
                        console.log(`Skip Copying student: ${info.name} because not found scores`);
                        continue;
                }
                let value = score.CN;
                if (hk === 1) {
                        value = score.HK1;
                }
                else if (hk === 2) {
                        value = score.HK2;
                }
                data.push({
                        row: info.row,
                        column: info.column + 3,
                        value: value
                });
        }
        updateCellsOfWorkSheet(vnptFileName, vnptWorkbook, indexOfVnptSheet, data);
}

function copyViettelToVnptFile(viettelWorkbook, vnptWorkbook, vnptFilePath) {
        let vnptSheetList = vnptWorkbook.SheetNames;
        let viettelSheetList = viettelWorkbook.SheetNames;

        if (vnptSheetList.length !== viettelSheetList.length) {
                console.log(`VIETTEL has ${viettelSheetList.length} sheets, VNPT has ${vnptSheetList.length} sheets.`);
                return;
        }

        for (let sheetName of vnptSheetList) {
                let vnptSheetIndex = vnptSheetList.indexOf(sheetName);
                let viettelSheetIndex = viettelSheetList.indexOf(sheetName);
                if (viettelSheetIndex === null) {
                        console.log(`Not found VNPT SheetName: ${sheetName} inside VIETTEL file.`);
                        return;
                }
                console.log(`Copying sheetName: ${sheetName}, viettel sheetIndex: ${viettelSheetIndex}, vnpt sheetIndex: ${vnptSheetIndex}`);
                let index = 3;
                if (vnptFilePath.includes("hk1")) {
                        index = 1;
                }
                else if (vnptFilePath.includes("hk2")) {
                        index = 2;
                }
                copyViettelSheetToVnptSheet(viettelWorkbook, viettelSheetIndex, vnptWorkbook, vnptSheetIndex, vnptFilePath, index);
        }
}

function main() {

        let xlsFiles = [
                {
                        viettel: './viettel.xls',
                        vnpt_hk1: './vnpt_hk1.xls',
                        vnpt_hk2: './vnpt_hk2.xls',
                        vnpt_cn: './vnpt_cn.xls',
                }
        ];
        
        for (let file of xlsFiles) {
                console.log(`Copying Viettel FILE ${file.viettel} to VNPT \n`);
                let viettelWorkbook = getWorkbook(file.viettel);
                let vnptWorkbookHk1 = getWorkbook(file.vnpt_hk1);
                let vnptWorkbookHk2 = getWorkbook(file.vnpt_hk2);
                let vnptWorkbookCn = getWorkbook(file.vnpt_cn);

                copyViettelToVnptFile(viettelWorkbook, vnptWorkbookHk1, file.vnpt_hk1);
                copyViettelToVnptFile(viettelWorkbook, vnptWorkbookHk2, file.vnpt_hk2);
                copyViettelToVnptFile(viettelWorkbook, vnptWorkbookCn, file.vnpt_cn);

                console.log('DONE!');
        }
}

main();