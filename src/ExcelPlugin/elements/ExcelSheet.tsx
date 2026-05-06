import React, { Component } from "react";
import ExcelColumn from "./ExcelColumn";
import type { ExcelValue, ExcelSheetData } from "react-xlsx-wrapper";

export interface ExcelSheetProps {
  name: string;
  data?: any[];
  dataSet?: ExcelSheetData[];
  value?: ExcelValue[] | (() => void);
  children?: React.ReactElement<typeof ExcelColumn>[];
}
export default class ExcelSheet extends Component<ExcelSheetProps> {
  constructor(props: ExcelSheetProps) {
    super(props);

    if (!React.Children.toArray(props.children).every(
      (child) => React.isValidElement(child) && child.type === ExcelColumn
    )) {
      throw new Error("<ExcelSheet> can only have <ExcelColumn> as children");
    }
  }

  render() {
    return null;
  }
}
