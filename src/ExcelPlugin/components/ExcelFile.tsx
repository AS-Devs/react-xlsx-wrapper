'use client';

import React from "react";
import { utils, writeFile } from "xlsx-js-style";
import type { BookType, WorkSheet } from "xlsx-js-style";
import { excelSheetFromAoA, excelSheetFromDataSet } from "../utils/DataUtil";
import type {
  ExcelColumnProps,
  ExcelFileProps,
  ExcelSheetProps,
  ExcelValue,
} from "react-xlsx-wrapper";

class ExcelFile extends React.Component<ExcelFileProps> {
  state = {
    fileName: "Download",
    fileExtension: "xlsx" as BookType,
    hideElement: false,
  };

  componentDidMount() {
    if (this.props.filename) {
      this.setState({
        fileName: this.props.filename,
      });
    }
    if (this.props.fileExtension) {
      this.setState({
        fileExtension: this.props.fileExtension,
      });
    }

    if (this.props.hideElement) {
      this.setState({
        hideElement: this.props.hideElement,
      });
    }
  }

  createSheetData = (sheet: React.ReactElement<any>) => {
    const columns = sheet.props.children;
    const sheetData = [
      React.Children.map(
        columns,
        (column: React.ReactElement<ExcelColumnProps>) => column.props.label
      ),
    ] as ExcelValue[][];

    const data = sheet.props.data;
    if (!data) throw new Error("No data provided");
    data.forEach((row: any) => {
      let sheetRow: ExcelValue[] = [];

      React.Children.forEach(
        columns,
        (column: React.ReactElement<ExcelColumnProps>) => {
          const getValue = (row: any) =>
            typeof column.props.value === 'function'
              ? (column.props.value as (row: any) => ExcelValue)(row)
              : row[column.props.value as string];
          const itemValue = getValue(row);
          sheetRow.push(isNaN(Number(itemValue)) ? itemValue || "" : itemValue);
        }
      );

      sheetData.push(sheetRow);
    });

    return sheetData;
  };

  download = () => {
    if (typeof window === 'undefined') return;
    const wb = utils.book_new();
    const fileName = this.getFileName();
    const fileExtension: BookType = this.getFileExtension();

    React.Children.forEach<React.ReactElement<ExcelSheetProps>>(
      this.props.children,
      (sheet) => {
        let ws: WorkSheet = {};
        const wsName = sheet.props.name || fileName.split(".")[0] || "Sheet1";
        if (
          typeof sheet.props.dataSet === "undefined" ||
          sheet.props.dataSet.length === 0
        ) {
          ws = excelSheetFromAoA(this.createSheetData(sheet));
        } else {
          ws = excelSheetFromDataSet(sheet.props.dataSet, sheet.props.bigHeading, sheet.props.autoFilterForAllColumn);
        }
        // add worksheet to workbook
        utils.book_append_sheet(wb, ws, wsName);
      }
    );

    writeFile(wb, fileName, {
      bookType: fileExtension,
      bookSST: true,
      type: "binary",
      cellStyles: true,
    });
  };

  getFileNameWithExtension = (filename: string, extension: string) => {
    return `${filename}.${extension}`;
  };

  getFileName = () => {
    if (
      this.state.fileName === null ||
      typeof this.state.fileName !== "string"
    ) {
      throw new Error("Invalid file name provided");
    }
    return this.getFileNameWithExtension(
      this.state.fileName?.split(".")[0],
      this.getFileExtension()
    );
  };

  getFileExtension = (): BookType => {
    if (this.props.fileExtension) {
      return this.state.fileExtension;
    }
    // Fall back to extension embedded in the filename
    const slugs = this.state.fileName.split(".");
    if (slugs.length > 1) {
      return slugs[slugs.length - 1] as BookType;
    }
    return this.state.fileExtension;
  };

  handleDownload = () => {
    this.download();
  };

  render() {
    const { element } = this.props;

    if (this.state.hideElement === true) {
      return null;
    } else {
      return <span onClick={this.handleDownload}>{element}</span>;
    }
  }
}

export default ExcelFile;
