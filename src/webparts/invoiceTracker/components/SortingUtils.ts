import { IColumn } from '@fluentui/react';

export enum SortType {
  String = "string",
  Number = "number",
  Date = "date"
}

const formatDate = (date: string): string => {
  let newDate = date.split("/");
  return `${newDate[2]}-${newDate[1]}-${newDate[0]}`;
};

const sortItems = (items: any[], sortBy: string, descending = false): any[] => {
  return items.sort((a, b) => {
    const aLowerCase = a[sortBy]?.toLowerCase() ?? '';
    const bLowerCase = b[sortBy]?.toLowerCase() ?? '';
    if (aLowerCase === bLowerCase) return 0;
    return descending ? (aLowerCase < bLowerCase ? 1 : -1) : (aLowerCase < bLowerCase ? -1 : 1);
  });
};

const sortNumber = (items: any[], sortBy: string, descending = false): any[] => {
  return items.sort((a, b) => {
    const aNo = parseInt(a[sortBy]);
    const bNo = parseInt(b[sortBy]);
    if (aNo === bNo) return 0;
    return descending ? (aNo < bNo ? 1 : -1) : (aNo > bNo ? 1 : -1);
  });
};

const sortDate = (items: any[], sortBy: string, descending = false): any[] => {
  return items.sort((a, b) => {
    const dateA = new Date(formatDate(a[sortBy]));
    const dateB = new Date(formatDate(b[sortBy]));
    if (dateA === dateB) return 0;
    return descending ? (dateA < dateB ? 1 : -1) : (dateA > dateB ? 1 : -1);
  });
};

export const columnSort = (
  ev: React.MouseEvent<HTMLElement>,
  selectedColumn: IColumn,
  columns: IColumn[],
  items: any[],
  type: SortType
): { sortedItems: any[]; updatedColumns: IColumn[] } => {
  const newColumns = columns.map(col => {
    if (col.key === selectedColumn.key) {
      col.isSortedDescending = !col.isSortedDescending;
      col.isSorted = true;
    } else {
      col.isSorted = false;
      col.isSortedDescending = true;
    }
    return col;
  });

  const currColumn = newColumns.find(col => col.key === selectedColumn.key);
  if (!currColumn) {
    console.warn(`Column key ${selectedColumn.key} not found in columns.`);
    // Return original items and columns without modification to avoid crash
    return { sortedItems: items, updatedColumns: columns };
  }

  let sortedItems: any[] = [];

  switch (type) {
    case SortType.String:
      sortedItems = sortItems([...items], currColumn.fieldName!, currColumn.isSortedDescending);
      break;
    case SortType.Number:
      sortedItems = sortNumber([...items], currColumn.fieldName!, currColumn.isSortedDescending);
      break;
    case SortType.Date:
      sortedItems = sortDate([...items], currColumn.fieldName!, currColumn.isSortedDescending);
      break;
    default:
      sortedItems = [...items];
  }

  return { sortedItems, updatedColumns: newColumns };
};

