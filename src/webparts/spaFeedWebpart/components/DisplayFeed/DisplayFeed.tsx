import * as React from 'react';
import styles from '../SpaFeedWebpart.module.scss';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISpaFeedWebpartProps } from '../ISpaFeedWebpartProps';
import { IListItem } from '../IListItem';

class DisplayFeed extends React.Component<ISpaFeedWebpartProps, {}>{

    state = {
        ListItems:[]
      };

      public componentDidMount():void{
        this.loadSPListData();
      }

      private loadSPListData() :void {
        this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Feeds')/items?$select=FeedTitle,FeedDescription`,  
            SPHttpClient.configurations.v1,  
            {  
              headers: {  
                'Accept': 'application/json;odata=nometadata',  
                'odata-version': ''  
              }  
            }) 
            .then((response: SPHttpClientResponse): Promise<IListItem[]> => {  
              return response.json();   
            })
            .then((items: IListItem[]): void => {  
              this.setState({ListItems:items["value"]});
            }, (error: any): void => {  
                console.log('error occurered')
            });  
       }

    public render() : React.ReactElement<ISpaFeedWebpartProps> {
            return (
                <div className={styles.DisplayWebpart}>
                        {this.state.ListItems.map(item=>(
                            <div className={styles.Feed}>
                                        <div className={styles.FeedTitle}>{item.FeedTitle}</div>
                                        <div className={styles.FeedDescription}>{item.FeedDescription}</div>
                            </div>
                        ))}                      
                </div>
            );
    }
}

export default DisplayFeed;