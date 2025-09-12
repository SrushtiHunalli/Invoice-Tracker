import * as React from 'react';
import styles from './CommonTable.module.scss';
import { ICommonTableProps } from './ICommonTableProps';
import { ICommonTableStates } from './ICommonTableStates';
import { DetailsList, Selection, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';

export default class CommonTable extends React.Component<ICommonTableProps, ICommonTableStates> {
  private selection: Selection;

  constructor(props: ICommonTableProps) {
    super(props);
    this.state = {
      items: []
    };
    this.selection = new Selection({ 
      onSelectionChanged: () => { this._getNewSelectionDetails(); } 
    });
  }

  private _getNewSelectionDetails() {
    if (this.selection.count && this.props.selectedItem)
      this.props.selectedItem(this.selection.getSelection()[0]);
    else if (this.props.selectedItem)
      this.props.selectedItem(false);
  }

  public render(): React.ReactElement<ICommonTableProps> {
    return (
      <div className={styles.commonTable + " ms-Grid"}>
        <div className={"ms-Grid-row " + styles.topSpace}>
          <div className={"ms-Grid-col ms-sm12 ms-md12 ms-lg12 "}>
            <DetailsList
              columns={this.props.mainColumns}
              items={this.props.tableContent}
              selection={this.selection}
              selectionMode={this.props.selectionMode ? this.props.selectionMode : SelectionMode.single}
              isHeaderVisible={true}
              styles={{
                headerWrapper: {
                  selectors: {
                    '.ms-DetailsHeader-cellTitle': { fontWeight: 700 }
                  }
                }
              }}
            />
          </div>
          <div className={"ms-Grid-col ms-sm12 ms-md12 ms-lg12 " + styles.emptyList}>
            {!this.props.tableContent.length && <label>No items found</label>}
          </div>
        </div>
      </div>
    );
  }
}
