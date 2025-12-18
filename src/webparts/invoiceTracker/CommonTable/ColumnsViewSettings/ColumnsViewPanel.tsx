import * as React from 'react';
import { Panel, PanelType, Checkbox, PrimaryButton } from 'office-ui-fabric-react';

interface IColumnsViewPanelProps {
  isOpen?: boolean;
  visibleColumns: any[]; // Array of currently visible columns (entire objects)
  allColumns: any[]; // Array of all available columns (entire objects)
  onClose: () => void;
  onUpdateColumns: (updatedColumns: any[]) => void;
  localStorageKey?: string;
}

class ColumnsViewPanel extends React.Component<IColumnsViewPanelProps, { selectedColumns: any[]; }> {
  constructor(props: IColumnsViewPanelProps) {
    super(props);

    this.state = {
      selectedColumns: this.props.visibleColumns,
    };
  }

  private handleUpdateColumns = () => {
    const { allColumns, onUpdateColumns, onClose } = this.props;
    const { selectedColumns } = this.state;

    if (!selectedColumns || selectedColumns.length === 0) {
      console.warn("No columns selected. Aborting update.");
      onClose();
      return;
    }
    // Make a copy of the existing selectedColumns array to avoid modifying it directly
    const updatedColumns = [...selectedColumns];

    // Update the copy based on user selection
    allColumns.forEach((column) => {
      const columnName = column.fieldName;
      const isChecked = selectedColumns.find((col) => col.fieldName === columnName);

      if (isChecked && !updatedColumns.find((col) => col.fieldName === columnName)) {
        updatedColumns.push(column);
      } else if (!isChecked && updatedColumns.find((col) => col.fieldName === columnName)) {
        updatedColumns.splice(
          updatedColumns.findIndex((col) => col.fieldName === columnName),
          1
        );
      }
    });

    // Update the parent component with the modified array
    onUpdateColumns(updatedColumns);
    console.log("Updated Columns");
    console.log(updatedColumns);

    onClose();
  };

  public render() {
    const { allColumns, onClose } = this.props;
    const { selectedColumns } = this.state;

    return (
      <Panel isOpen type={PanelType.smallFixedFar} onDismiss={onClose} headerText="Edit View Columns">
        <div style={{ padding: '16px' }}>
          {allColumns.map((column) => (
            <Checkbox
              key={column.key}
              label={column.name}
              checked={selectedColumns.some((col) => col.fieldName === column.fieldName)}
              onChange={() => {
                this.setState((prevState) => {
                  const prevSelectedColumns = prevState.selectedColumns;
                  return {
                    selectedColumns: prevSelectedColumns.some((col) => col.fieldName === column.fieldName)
                      ? prevSelectedColumns.filter((col) => col.fieldName !== column.fieldName)
                      : [...prevSelectedColumns, column],
                  };
                });
              }}
              styles={{ root: { marginBottom: '8px' } }}
            />
          ))}
        </div>
        <PrimaryButton
          onClick={this.handleUpdateColumns}
          disabled={this.state.selectedColumns.length === 0}>Apply Changes</PrimaryButton>
      </Panel>
    );
  }
}

export default ColumnsViewPanel;
