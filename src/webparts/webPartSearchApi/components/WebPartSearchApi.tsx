import * as React from 'react';
import styles from './WebPartSearchApi.module.scss';
import { IWebPartSearchApiProps } from './IWebPartSearchApiProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class WebPartSearchApi extends React.Component<IWebPartSearchApiProps, {}> {

    public componentDidMount(): void {
        this._getList();
    }

    public componentWillReceiveProps(): void {
        this._getList();
    }


    // @ts-ignore
    private async _getList(): void {
        if (this.props.nameList !== '') {
            const Web1 = (await import(/*webpackChunkName: '@pnp_sp' */ "@pnp/sp")).Web;
            const web = new Web1(this.props.context.pageContext.web.absoluteUrl+'/sites/Dev1');
            web.lists.getByTitle(this.props.nameList).items.get().then((response) => {
                console.log(response);
            });
        }

    }

  public render(): React.ReactElement<IWebPartSearchApiProps> {
    return (
      <div className={ styles.webPartSearchApi }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.description }>{escape(this.props.nameList)}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
