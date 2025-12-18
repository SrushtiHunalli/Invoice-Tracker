import { IColumn } from "office-ui-fabric-react";

export interface ICustomDetailsListColumn extends IColumn {
  type?: "text" | "date" | "yesNo" | "multilinetext" | "number";
}