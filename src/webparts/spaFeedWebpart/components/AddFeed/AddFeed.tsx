import * as React from 'react';
import styles from '../SpaFeedWebpart.module.scss';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISpaFeedWebpartProps } from '../ISpaFeedWebpartProps';
import {Dialog} from '@microsoft/sp-dialog';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { PeoplePicker,PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';


class AddFeed extends React.Component<ISpaFeedWebpartProps, {}>{

    state={
        Feed: {
            FeedTitle:{
                value:'',
                id:1,
                required: true,
                isvalid:false
            },
            FeedDescription:{
                value:'',
                id:2,
                required: true,
                isvalid:false
            },
            FeedType:{
                value:'General',
                id:3,
                required: true,
                isvalid:true
            },
            FeedLocation: {
                value:{
                    Label: '',
                    TermGuid: '',
                    WssId: -1
                },
                id:4,
                required: true,
                isvalid:false
            },
            FeedOwnerId: {
                value:'',
                id:5,
                required: true,
                isvalid:false
            },
            Title:{
                value:'NA',
                required: true,
                isvalid:true
            }
        },
        showPanel: false,
        isFormvalid:true
    }

    public inputChangedHandler = (event,inputIdentifier) : void => {
        const updatedFeedForm = {
            ...this.state.Feed
        };
        const updatedFeedElement = {
            ...updatedFeedForm[inputIdentifier]
        };
        updatedFeedElement.value=event.target.value;
        if(updatedFeedElement.value.trim() ==='' && updatedFeedElement.required===true) {
             updatedFeedElement.isvalid=false;
        }
        else{
            updatedFeedElement.isvalid=true;
        }
        updatedFeedForm[inputIdentifier]=updatedFeedElement;
        this.setState({Feed:updatedFeedForm});
}

public onTaxPickerChange = (terms : IPickerTerms): void => {

    const updatedFeedForm = {
        ...this.state.Feed
    };
    const updatedFeedElement = {
        ...updatedFeedForm["FeedLocation"]
    };

    if(terms.length>0){
        updatedFeedElement.value.TermGuid=terms[0].key.toString();
        updatedFeedElement.value.Label=terms[0].name.toString();
        updatedFeedForm["FeedLocation"]=updatedFeedElement;
        updatedFeedElement.isvalid=true;
    }
    else{
        updatedFeedElement.value.TermGuid='';
        updatedFeedElement.value.Label='';

        if(updatedFeedElement.value.TermGuid.trim() ==='' && updatedFeedElement.required===true) {
            updatedFeedElement.isvalid=false;
       }
       else{
           updatedFeedElement.isvalid=true;
       }
    }
    this.setState({Feed:updatedFeedForm}); 
}

public getPeoplePickerItems = (items: any[]) => {
    const updatedFeedForm = {
        ...this.state.Feed
    };
    const updatedFeedElement = {
        ...updatedFeedForm["FeedOwnerId"]
    };

    if(items.length>0){
        updatedFeedElement.value=items[0].id.toString();
    }
    else{
        updatedFeedElement.value='';
    }
    
    updatedFeedForm["FeedOwnerId"]=updatedFeedElement;
    if(updatedFeedElement.value.trim() ==='' && updatedFeedElement.required===true) {
        updatedFeedElement.isvalid=false;
   }
   else{
       updatedFeedElement.isvalid=true;
   }
    this.setState({Feed:updatedFeedForm});
}

private _onSubmit = () => {
    this.setState({ showPanel: true });
}

private _onClosePanel = () => {
    this.setState({ showPanel: false });
}

private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div style={{display:'inline'}}>
        <PrimaryButton onClick={()=>this.createItem()} style={{ marginRight: '8px' }}>
          Confirm
      </PrimaryButton>
        <DefaultButton onClick={()=>this._onClosePanel()}>Cancel</DefaultButton>
      </div>
    );
  }

  private validateForm() : void{
    let isValid=true;

    for (let formElementIdentifier in this.state.Feed){

        if(this.state.Feed[formElementIdentifier].isvalid && isValid){
            isValid=true
        }
        else{
            isValid=false
        }
    }
    this.setState({isFormvalid:isValid});

    if(isValid){
        this._onSubmit();
    }
    else{
        this.setState({ showPanel: false });
    }
  }

private createItem(): void {  
    this.setState({ showPanel: false });
    const formData={};
    for (let formElementIdentifier in this.state.Feed){
        formData[formElementIdentifier]= this.state.Feed[formElementIdentifier].value;
    }
 
    const body: string = JSON.stringify(formData); 
    console.log(body); 

    this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Feeds')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: body  
    })  
    .then((response: SPHttpClientResponse): any => {  
      return response.json();  
    })  
    .then((item: any): void => {  
       Dialog.alert("List item created successfully");
    }, (error: any): void => {  
      console.log(error);
      Dialog.alert("List item creation failed");
    });  
  } 
    public render() : React.ReactElement<ISpaFeedWebpartProps> {

        var errorMesage='';
        if(!this.state.isFormvalid){
            errorMesage= 'Please fill all mandatory fields.'         
        }
            return (
            <div className={styles.Add}>
            <div className= "col-md-12" style={{backgroundColor:'white',border:'1px solid #f09311'}}>
                  <div className="col-md-12" style={{marginTop:'10px',marginBottom:'10px'}}>
                        <label style={{color:'red',fontWeight:'bold'}}>{errorMesage}</label>
                    </div>
            <div className="col-md-12" style={{marginTop:'10px'}}>
                <div className="col-md-2">
                    <label style={{fontWeight:'bold'}}>Title <label style={{color:'red'}}>*</label></label>
                </div>
                <div className="col-md-10">
                    <input className={styles.FocusDiv} style={{width:'100%'}} type="text" onChange={(event)=>this.inputChangedHandler(event,'FeedTitle')}></input>
                </div>
            </div>
            <div className="col-md-12" style={{marginTop:'10px'}}>
            <div className="col-md-2">
                <label style={{fontWeight:'bold'}}>Category <label style={{color:'red'}}>*</label></label>
            </div>
            <div className="col-md-10">
                <select className={styles.FocusDiv} style={{width:'100%'}} onChange={(event)=>
                    this.inputChangedHandler(event,'FeedType')}>
                    <option value="General">General</option>
                    <option value="Politics">Politics</option>
                    <option value="Technology">Technology</option>
                </select>
            </div>
            </div>
            <div className="col-md-12" style={{marginTop:'10px'}}>
            <div className="col-md-2">
                <label style={{fontWeight:'bold'}}>Location <label style={{color:'red'}}>*</label></label>
            </div>
            <div className="col-md-10">
                <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="FeedLocation"
                    panelTitle="Select Term"
                    label=""
                    context={this.props.context}
                    onChange={this.onTaxPickerChange}
                    isTermSetSelectable={false}/>
            </div>
            </div>
            <div className="col-md-12" style={{marginTop:'10px'}}>
            <div className="col-md-2">
                <label style={{fontWeight:'bold'}}>Owner <label style={{color:'red'}}>*</label></label>
            </div>
            <div className="col-md-10">
                <PeoplePicker context={this.props.context}
                    titleText=""
                    personSelectionLimit={1}
                    groupName={""} 
                    showtooltip={false}
                    isRequired={false}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this.getPeoplePickerItems}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    />
            </div>
            </div>
            <div className="col-md-12" style={{marginTop:'10px'}}>
            <div className="col-md-2">
                <label style={{fontWeight:'bold'}}>Description <label style={{color:'red'}}>*</label></label>
            </div>
            <div className="col-md-10">
                <textarea className={styles.FocusDiv} style={{minHeight:'200px',width:'100%'}} onChange={(event)=>this.inputChangedHandler(event,'FeedDescription')}>
                </textarea>
            </div>
            </div>
            <div className="col-md-12" style={{marginTop:'10px',marginBottom:'10px'}}>
            <div className="col-md-2">
                <button type="button" className={styles.CustomButton} style={{backgroundColor:'orange',marginLeft:'10%'}} onClick={()=>this.validateForm()}>Submit</button>
            </div>
            <div className="col-md-10">
            </div>
            </div>
            <Panel isOpen={this.state.showPanel}
            type={PanelType.smallFixedFar}
            onDismiss={this._onClosePanel}
            isFooterAtBottom={false}
            headerText="Are you sure you want to submit this request?"
            closeButtonAriaLabel="Close"
            onRenderFooterContent={this._onRenderFooterContent}>
            <span>Please check the details filled and click on Confirm button to submit this request.</span>
            </Panel>  
                        </div>
            </div>
            );
    }
}

export default AddFeed;