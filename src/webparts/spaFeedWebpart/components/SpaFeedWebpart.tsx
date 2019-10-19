import * as React from 'react';
import styles from './SpaFeedWebpart.module.scss';
import { ISpaFeedWebpartProps } from './ISpaFeedWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import DisplayFeed from '../components/DisplayFeed/DisplayFeed';
import AddFeed from '../components/AddFeed/AddFeed';
import { SPComponentLoader } from '@microsoft/sp-loader';
require('bootstrap');

export default class SpaFeedWebpart extends React.Component<ISpaFeedWebpartProps, {}> {

  public state = {
    showDisp: true,
    showAdd: false
  }

  public GenerateDispView(props):React.ReactElement<ISpaFeedWebpartProps>{
    const view = props.view;
    if (view) {
      return <DisplayFeed {...props}/>;
    }
    else  {
      return null
    }
  }

  public GenerateAddView(props):React.ReactElement<ISpaFeedWebpartProps>{
    const view = props.view;
    if (view) {
      return <AddFeed {...props}/>;
    }
    else  {
      return null
    }
  }

  private ShowDisplay(): void{
  this.setState(() => {  
    return {  
      ...this.state,  
      showDisp: true,
      showAdd:false  
    };  
  });  
  } 

  private ShowAdd(): void{
    this.setState(() => {  
      return {  
        ...this.state,  
        showDisp: false,
        showAdd:true  
      };  
    });  
  } 

  public render():React.ReactElement<ISpaFeedWebpartProps>{
    
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);

    let dispColor: string ='';
    let addColor : string='';
    if(this.state.showDisp){
      dispColor='#a10316';
      addColor='#f09311';
    }
    else
    {
     addColor='#a10316';
     dispColor='#f09311';
    }

    return (
      <div className={ styles.spaFeedWebpart }>
          <div className={ styles.container }>
            <div className={styles.HeaderRow}>   
              <button className={styles.CustomButton} style={{backgroundColor:dispColor,borderColor:dispColor}} 
              onClick={() => this.ShowDisplay()}>Display Feeds</button>
              <button className={styles.CustomButton} style={{backgroundColor:addColor,borderColor:addColor}} 
              onClick={() => this.ShowAdd()}>Add Feed</button>
            </div>
              <this.GenerateDispView view= {this.state.showDisp} {...this.props}/>
              <this.GenerateAddView view= {this.state.showAdd} {...this.props}/>           
        </div>
      </div>
    );
  }
}
