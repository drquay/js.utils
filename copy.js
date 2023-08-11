const XLSX = require('xlsx');

let VNPT_IGNORED_OF_NAME_FINDING = ['Họ và tên', 'Giỏi', 'Khá', 'Trung bình', 'Yếu', 'Kém', ''];

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

function viettelGetStudentScoreByName(viettelWorkSheetJsonData, studentFullName) {
        for (let jsonObject of viettelWorkSheetJsonData) {
                if (jsonObject.hasOwnProperty('__EMPTY') && jsonObject.__EMPTY === studentFullName) {
                        return {
                                DOB: jsonObject.__EMPTY_2,
                                TX1: jsonObject.__EMPTY_3,
                                TX2: jsonObject.__EMPTY_4,
                                TX3: jsonObject.__EMPTY_5,
                                TX4: jsonObject.__EMPTY_6,
                                FN1: jsonObject.__EMPTY_19,
                                FN2: jsonObject.__EMPTY_20,
                                FN3: jsonObject.__EMPTY_21,
                                NHX: jsonObject.__EMPTY_23,
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
                for (let info of student.infos) {
                        let newValue = info.value;
                        let targetColumnIndex = info.column;
                        if (newValue === undefined) {
                                newValue = '';
                        }
                        let cellAddress = XLSX.utils.encode_cell({ c: targetColumnIndex, r: targetRowIndex });
                        XLSX.utils.sheet_add_aoa(sheet, [[newValue]], {origin: cellAddress});
                }
        }
        XLSX.writeFile(workbook, vnptFileName);
        console.log(`Finished Copying ${students.length} Students in sheet: ${worksheetName}`);
}

function copyViettelSheetToVnptSheet(viettelWorkbook, indexOfViettelSheet, vnptWorkbook, indexOfVnptSheet, vnptFileName) {
        
        let viettelWorksheet = getWorksheet(viettelWorkbook, indexOfViettelSheet);
        let viettelWorksheetJsonData = XLSX.utils.sheet_to_json(viettelWorksheet);

        let vnptWorksheet = getWorksheet(vnptWorkbook, indexOfVnptSheet);
        let studentInfos = vnptGetFullNameAndCellIndexOfStudents(vnptWorksheet);

        let data = [];
        for (let info of studentInfos) {
                let score = viettelGetStudentScoreByName(viettelWorksheetJsonData, info.name);
                if (score === undefined || score === null) {
                        console.log(`Skip Copying student: ${info.name} scores`);
                        continue;
                }
                data.push({
                        row: info.row,
                        infos: 
                        [
                                {
                                        value: score.TX1,
                                        column: info.column + 1
                                },
                                {
                                        value: score.TX2,
                                        column: info.column + 2
                                },
                                {
                                        value: score.TX3,
                                        column: info.column + 3
                                },
                                {
                                        value: score.TX4,
                                        column: info.column + 4
                                },
                                {
                                        value: score.FN1,
                                        column: info.column + 5
                                },
                                {
                                        value: score.FN2,
                                        column: info.column + 6
                                },
                                {
                                        value: score.FN3,
                                        column: info.column + 7
                                },
                                {
                                        value: score.NHX,
                                        column: info.column + 8
                                }
                        ]
                });
        }
        updateCellsOfWorkSheet(vnptFileName, vnptWorkbook, indexOfVnptSheet, data);
}

function simpleVnptSheetName(vnptSheetName) {
        let parts = vnptSheetName.split('_');
        let vnptSubjectName = "";
        for (let i = 0; i < parts.length - 1; i++) {
                vnptSubjectName += parts[i];
        }
        return vnptSubjectName;
}

function findVnptSheetNameInViettelWorkBook(viettelSheetList, vnptSheetName) {
        let simpleVnptName = simpleVnptSheetName(vnptSheetName).toLocaleLowerCase();
        for (let viettelSheetName of viettelSheetList) {
                let simpleViettelName = viettelSheetName.split('_')[1].toLocaleLowerCase();
                if (simpleVnptName.includes(simpleViettelName) || simpleViettelName.includes(simpleVnptName)) {
                        return viettelSheetList.indexOf(viettelSheetName);
                }
        }
        return null;
}

function copyViettelToVnptFile(viettelWorkbook, vnptWorkbook, vnptFilePath) {
        let vnptSheetList = vnptWorkbook.SheetNames;
        let viettelSheetList = viettelWorkbook.SheetNames;

        if (vnptSheetList.length !== viettelSheetList.length) {
                console.log("Number Of Sheet of 2 Files ARE DIFFERENT. CAN NOT COPY ??????????????");
                return;
        }

        for (let sheetName of vnptSheetList) {
                let vnptSheetIndex = vnptSheetList.indexOf(sheetName);
                let viettelSheetIndex = findVnptSheetNameInViettelWorkBook(viettelSheetList, sheetName);
                if (viettelSheetIndex === null) {
                        console.log(`Can not found VNPT SheetName: ${sheetName} inside VIETTEL file. CAN NOT COPY ??????????????`);
                        return;
                }
                console.log(`Copying sheetName: ${sheetName}, viettel sheetIndex: ${viettelSheetIndex}, vnpt sheetIndex: ${vnptSheetIndex}`);
                copyViettelSheetToVnptSheet(viettelWorkbook, viettelSheetIndex, vnptWorkbook, vnptSheetIndex, vnptFilePath);
        }
}

function main() {

        let xlsFiles = [
                {
                        viettel: './vietel.xls',
                        vnpt: './vnpt.xls',
                }
        ];
        
        for (let file of xlsFiles) {
                let startTime = new Date();
                console.log(`==================================== VNPT: ${file.vnpt} - VIETTEL: ${file.viettel} - STARTED AT: ${startTime.getHours()}:${startTime.getMinutes()}:${startTime.getSeconds()} ====================================`);
                
                let viettelWorkbook = getWorkbook(file.viettel);
                let vnptWorkbook = getWorkbook(file.vnpt);
                copyViettelToVnptFile(viettelWorkbook, vnptWorkbook, file.vnpt);

                let endTime = new Date();
                let timeDifferenceInMinutes = (endTime - startTime) / (1000 * 60);
                console.log(`==================================== VNPT: ${file.vnpt} - VIETTEL: ${file.viettel} - FINISHED AT: ${endTime.getHours()}:${endTime.getMinutes()}:${endTime.getSeconds()} ====================================`);
                console.log(`==================================== VNPT: ${file.vnpt} - VIETTEL: ${file.viettel} - TOOK: ${timeDifferenceInMinutes} Minutes ====================================`);
        }
}

main();