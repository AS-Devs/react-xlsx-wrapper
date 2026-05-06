"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.excelSheetFromDataSet = exports.excelSheetFromAoA = exports.dateToNumber = void 0;
const xlsx_js_style_1 = require("xlsx-js-style");
const dateToNumber = (v, date1904) => {
    const epoch = Date.parse(v);
    const serial = (epoch - Number(new Date(Date.UTC(1899, 11, 30)))) / (24 * 60 * 60 * 1000);
    // 1904 date system shifts the epoch by 1462 days (applied to the serial, not the string)
    return date1904 ? serial + 1462 : serial;
};
exports.dateToNumber = dateToNumber;
/**
 * This returns the worksheet object for the given dataSet, also accept bigHeading.
 *
 * @param dataSet - The ExcelSheetData array is required
 * @param bigHeading - ExcelSheetCol (Optional)
 * @returns WorkSheet Object
 *
 * @author Susanta Chakraborty
 * @date 2023-06-14
 */
const excelSheetFromDataSet = (dataSet, bigHeading, autoFilterForAllColumn) => {
    /*
    Assuming the structure of dataset
    {
        xSteps?: number; //How many cells to skips from left
        ySteps?: number; //How many rows to skips from last data
        columns: [array | string]
        data: [array_of_array | string|boolean|number | CellObject]
        fill, font, numFmt, alignment, and border
    }
     */
    if (dataSet === undefined || dataSet.length === 0) {
        return {};
    }
    let ws = {};
    let range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    let rowCount = 0;
    dataSet.forEach(dataSetItem => {
        let columns = dataSetItem.columns;
        let xSteps = typeof (dataSetItem.xSteps) === 'number' ? dataSetItem.xSteps : 0;
        let ySteps = typeof (dataSetItem.ySteps) === 'number' ? dataSetItem.ySteps : 0;
        let data = dataSetItem.data;
        if (!dataSetItem.columns?.length && !dataSetItem.data?.length) {
            return;
        }
        rowCount += ySteps;
        if (bigHeading?.title && columns.length >= 0) {
            columns.forEach((_, index) => {
                const cellRef = xlsx_js_style_1.utils.encode_cell({ c: xSteps + index, r: rowCount });
                fixRange(range, 0, index, rowCount, xSteps, ySteps);
                getHeaderCell(bigHeading, cellRef, ws, true, index);
            });
            const mergedRange = { s: { c: xSteps, r: rowCount }, e: { c: xSteps + dataSetItem.columns.length - 1, r: rowCount } };
            ws['!merges'] = [mergedRange];
            rowCount += 1;
        }
        let columnsInfo = [];
        // if xStep has value then we need to skip some columns
        if (xSteps > 0) {
            for (let i = 0; i < xSteps; i++) {
                columnsInfo.push({ wpx: 100 });
            }
        }
        if (columns.length >= 0) {
            columns.forEach((col, index) => {
                let cellRef = xlsx_js_style_1.utils.encode_cell({ c: xSteps + index, r: rowCount });
                fixRange(range, 0, index, rowCount, xSteps, ySteps);
                let colTitle = col;
                if (typeof col === 'object') {
                    //colTitle = col.title; //moved to getHeaderCell
                    columnsInfo.push(col.width || { wpx: 100, hidden: false }); /* wch (chars), wpx (pixels) - e.g. [{wch:6},{wpx:50}] */
                }
                getHeaderCell(colTitle, cellRef, ws);
            });
            if (autoFilterForAllColumn) {
                const filterRange = { s: { c: xSteps, r: rowCount }, e: { c: xSteps + dataSetItem.columns.length - 1, r: rowCount } };
                const filterRef = xlsx_js_style_1.utils.encode_range(filterRange);
                ws['!autofilter'] = { ref: filterRef };
            }
            rowCount += 1;
        }
        if (columnsInfo.length > 0) {
            ws['!cols'] = columnsInfo;
        }
        for (let R = 0; R !== data.length; ++R, rowCount++) {
            for (let C = 0; C !== data[R].length; ++C) {
                let cellRef = xlsx_js_style_1.utils.encode_cell({ c: C + xSteps, r: rowCount });
                fixRange(range, R, C, rowCount, xSteps, ySteps);
                getCell(data[R][C], cellRef, ws);
            }
        }
    });
    if (range.s.c < 10000000) {
        ws['!ref'] = xlsx_js_style_1.utils.encode_range(range);
    }
    return ws;
};
exports.excelSheetFromDataSet = excelSheetFromDataSet;
function getHeaderCell(v, cellRef, ws, isHeader, index) {
    const bigHeadingDefualtStyle = {
        font: { bold: true, name: "Archive", sz: 24, color: { rgb: "333" } },
        fill: { patternType: "solid", fgColor: { rgb: "FFFFFF" } },
        alignment: { vertical: "center", horizontal: "center" },
    };
    const cell = {
        t: 's',
    };
    let headerCellStyle = v.style ? v.style : { font: { bold: true } }; //if style is then use it
    if (isHeader) {
        cell.v = index === 0 ? v.title : '';
        headerCellStyle = v.style ? v.style : bigHeadingDefualtStyle;
    }
    else {
        cell.v = v.title;
    }
    cell.t = 's';
    cell.s = headerCellStyle;
    ws[cellRef] = cell;
}
function getCell(v, cellRef, ws) {
    if (v === null) {
        return;
    }
    const isDate = v instanceof Date;
    const cell = { t: 's' };
    let cellValue;
    if (typeof v !== 'object') {
        cellValue = v;
        cell.v = v;
    }
    else if (isDate) {
        cellValue = v;
    }
    else {
        cellValue = v.value;
        cell.v = v.value;
        cell.s = v.style;
    }
    if (isDate) {
        cell.t = 'n';
        cell.z = xlsx_js_style_1.SSF._table[14];
        cell.v = dateToNumber(cellValue.toString(), false);
    }
    else if (typeof cellValue === 'number') {
        if (isNaN(cellValue)) {
            cell.t = 'e';
            cell.v = 0x24; // #NUM!
        }
        else if (!isFinite(cellValue)) {
            cell.t = 'e';
            cell.v = 0x07; // #DIV/0!
        }
        else {
            cell.t = 'n';
        }
    }
    else if (typeof cellValue === 'boolean') {
        cell.t = 'b';
    }
    else {
        cell.t = 's';
    }
    ws[cellRef] = cell;
}
function fixRange(range, R, C, rowCount, xSteps, ySteps) {
    if (range.s.r > R + rowCount) {
        range.s.r = R + rowCount;
    }
    if (range.s.c > C + xSteps) {
        range.s.c = C + xSteps;
    }
    if (range.e.r < R + rowCount) {
        range.e.r = R + rowCount;
    }
    if (range.e.c < C + xSteps) {
        range.e.c = C + xSteps;
    }
}
const excelSheetFromAoA = (data) => {
    let ws = {};
    let range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    for (let R = 0; R !== data.length; ++R) {
        for (let C = 0; C !== data[R].length; ++C) {
            if (range.s.r > R) {
                range.s.r = R;
            }
            if (range.s.c > C) {
                range.s.c = C;
            }
            if (range.e.r < R) {
                range.e.r = R;
            }
            if (range.e.c < C) {
                range.e.c = C;
            }
            let cell = { v: data[R][C], t: 's' };
            if (cell.v === null) {
                continue;
            }
            let cellRef = xlsx_js_style_1.utils.encode_cell({ c: C, r: R });
            if (typeof cell.v === 'number' && isNaN(cell.v)) {
                cell.t = 'e';
                cell.v = 0x24; // #NUM!
            }
            else if (typeof cell.v === 'number' && !isFinite(cell.v)) {
                cell.t = 'e';
                cell.v = 0x07; // #DIV/0!
            }
            else if (typeof cell.v === 'number') {
                cell.t = 'n';
            }
            else if (typeof cell.v === 'boolean') {
                cell.t = 'b';
            }
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = xlsx_js_style_1.SSF._table[14];
                cell.v = dateToNumber(cell.v?.toDateString(), false);
            }
            else {
                cell.t = 's';
            }
            ws[cellRef] = cell;
        }
    }
    if (range.s.c < 10000000) {
        ws['!ref'] = xlsx_js_style_1.utils.encode_range(range);
    }
    return ws;
};
exports.excelSheetFromAoA = excelSheetFromAoA;
