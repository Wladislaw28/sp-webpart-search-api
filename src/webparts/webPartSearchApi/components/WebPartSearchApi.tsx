import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IWebPartSearchApiProps } from './IWebPartSearchApiProps';
import * as strings from 'WebPartSearchApiWebPartStrings';
import { taxonomy, ITermData } from "@pnp/sp-taxonomy";
import styles from './WebPartSearchApi.module.scss';

export interface ArrayItems {
    Title: string;
    Color: string;
}

export interface WebPartSearchApiState {
    arrayItems: ArrayItems[];
    columns: IColumn[];
    termsData: ITermData[];
}

export default class WebPartSearchApi extends React.Component<IWebPartSearchApiProps, WebPartSearchApiState> {

    public state = {
        arrayItems: [],
        columns: [],
        termsData: []
    };

    public componentDidMount(): void {
        this._getSeactData();
        this._getTermData();
    }

    public componentWillReceiveProps(): void {
        this._getSeactData();
        this._getTermData();
    }

    // @ts-ignore
    private async _getTermData(): void {
        const termData:ITermData[] = await taxonomy.termStores.getById('6883f0ba60c844ed8e7245852ff59257').getTermGroupById('e478e78e-5c2a-45c2-bdac-e8b7cdc908ef').termSets.getById(this.props.idTermSet).terms.get();
       this.setState({
          termsData: termData
       });
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
                columns: this._columsCreate([strings.TitleColums,strings.ColorColums])
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

    private _filterItems(nameColor: string, e: any): void {
        e.preventDefault();
        fetch(`https://mihasev28wmreply.sharepoint.com/search/_api/search/query?querytext='GTSet|%23c1a40892-eea2-4fd1-b75e-179737829e20'&selectproperties='Title%2ctitleColor%2cRefinableString50%2cRefinableString5'&refiners='RefinableString50%2cRefinableString51'&refinementfilters='RefinableString50:equals("${nameColor}")'&clienttype='ContentSearchRegular'`, {
            method: 'get',
            headers: {
                'accept': "application/json;odata=nometadata",
                'content-type': "application/json;odata=nometadata",
            }
        }).then((response) => response.json()).then((respo) =>
        {
            const arrayData: Array<any> = respo.PrimaryQueryResult.RelevantResults.Table.Rows;
            if (arrayData.length >= 1)  {
                return this._mapArrayItems(arrayData);
            }else {
                this.setState({
                    arrayItems: []
                });
            }
        });
    }

  public render(): React.ReactElement<IWebPartSearchApiProps> {
        const {arrayItems, columns, termsData} = this.state;
    return (
      <div className={ styles.webPartSearchApi }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{strings.WelcomeTitle}</span>
                <div className={styles.button_filter}>
                    <nav>
                        <ul className={styles.nav_filter}>
                            {termsData.length >= 1 ?
                                termsData.map((item, index) => {
                                    return(
                                        <li key={index} className={styles.nav_filter_li}
                                            onClick={(e) => this._filterItems(item.Name,e)}>
                                            {item.Name}
                                        </li>
                                    );
                                })
                                :null}
                        </ul>
                    </nav>
                </div>
                <div>
                    {arrayItems.length >= 1 ?
                        <DetailsList items={arrayItems}
                                     columns={columns}
                                     setKey="set"
                                     layoutMode={DetailsListLayoutMode.justified}
                                     isHeaderVisible={true}
                                     selectionPreservedOnEmptyClick={true}
                                     enterModalSelectionOnTouch={true}
                                     ariaLabelForSelectionColumn="Toggle selection"
                                     ariaLabelForSelectAllCheckbox="Toggle selection for all items" /> :
                        <h1 className={styles.title}>No items</h1>}
                </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
