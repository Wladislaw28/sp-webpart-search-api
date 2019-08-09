import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IWebPartSearchApiProps } from './IWebPartSearchApiProps';
import * as strings from 'WebPartSearchApiWebPartStrings';

import styles from './WebPartSearchApi.module.scss';

export interface ArrayItems {
    Title: string;
    Color: string;
}

export interface WebPartSearchApiState {
    arrayItems: ArrayItems[];
    columns: IColumn[];
}

export default class WebPartSearchApi extends React.Component<IWebPartSearchApiProps, WebPartSearchApiState> {

    public state = {
        arrayItems: [],
        columns: []
    };

    public componentDidMount(): void {
        this._getSeactData();
    }

    public componentWillReceiveProps(): void {
        this._getSeactData();
    }

    private _getSeactData() : void {
        fetch(`https://mihasev28wmreply.sharepoint.com/search/_api/search/query?querytext='${this.props.idTermSet}'&selectproperties='Title%2cRefinableString50%2ctitleColor'&clienttype='ContentSearchRegular'`, {
            method: 'get',
            headers: {
                'accept': "application/json;odata=nometadata",
                'content-type': "application/json;odata=nometadata",
            }
        }).then((response) => response.json()).then((respo) =>
        {
            return this._mapArrayItems(respo.PrimaryQueryResult.RelevantResults.Table.Rows);
        });
    }

    private _mapArrayItems(arrayData: Array<any>): void {
        const dataMap: ArrayItems[] = [];
        arrayData.forEach((item) => {
            dataMap.push({
                Title: item.Cells[2].Value,
                Color: item.Cells[3].Value
            });
            this.setState({
                arrayItems: dataMap,
                columns: this._columsCreate(['Title','Color'])
            });
        });
    }

    private _columsCreate(arraySelect: Array<any>): Array<IColumn> {
        const columns: IColumn[] = [];
        arraySelect.forEach((item,index) => {
            columns.push({
                key: `column${index}`,
                name: item,
                fieldName: item,
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
            });
        });
        return columns;
    }

  public render(): React.ReactElement<IWebPartSearchApiProps> {
        const {arrayItems, columns} = this.state;
    return (
      <div className={ styles.webPartSearchApi }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{strings.WelcomeTitle}</span>
                <div>
                    {arrayItems.length > 1 ?
                        <DetailsList items={arrayItems}
                                     columns={columns}
                                     setKey="set"
                                     layoutMode={DetailsListLayoutMode.justified}
                                     isHeaderVisible={true}
                                     selectionPreservedOnEmptyClick={true}
                                     enterModalSelectionOnTouch={true}
                                     ariaLabelForSelectionColumn="Toggle selection"
                                     ariaLabelForSelectAllCheckbox="Toggle selection for all items" /> : null}
                </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
