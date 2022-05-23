import * as React from 'react';
import styles from './ReactTaxonomy.module.scss';
import { IReactTaxonomyProps } from './IReactTaxonomyProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { taxonomy, ITerm, ITermSet, ITermStore } from '@pnp/sp-taxonomy';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { MessageBar, MessageBarType, IStackProps, Stack, ThemeSettingName } from 'office-ui-fabric-react';  
import { ComboBox, IComboBoxOption, IComboBox, PrimaryButton } from 'office-ui-fabric-react/lib/index';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

/*
export interface IPTerm {
  parent?: string;
  id: string;
  name: string;
}
*/

export interface DivisionTerm {
  id: string;
  name: string;
}
 
export interface IReactTaxonomyState {
  //terms: IPTerm[];
  divisionTag: IPickerTerms;
  officeTag: IPickerTerms;
  division : string;
  divisionSelect: any;
}

export default class ReactTaxonomy extends React.Component<IReactTaxonomyProps, IReactTaxonomyState> {
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      //terms: [],
      divisionTag: [],
      officeTag: [],
      division: "",
      divisionSelect: ""
      //showMessageBar: false
    };
    this._gettags();
  }

  private async _gettags() {
    const item: any = await sp.web.lists.getByTitle("Test_Metadata").items.getById(1).get();
    let selectedtag: any = [];
    item.Tags.forEach(function (v: any[], i) {
      selectedtag.push({ key: v["TermGuid"], name: v["Label"] })
    });
    console.log(item);
    this.setState({
      divisionTag: selectedtag
    });
  }

/*  
  public async getTermsetWithChildren(): Promise<IPTerm[]> {
    let tms: IPTerm[] = [];
    return new Promise<any[]>((resolve, reject) => {
      const tbatch = taxonomy.createBatch();
      return taxonomy.termStores.getById("0c10deb3-8dd7-4942-8006-05e787042858").get().then((resp1: ITermStore) => {        
        return resp1.getTermGroupById("1cd3d4ea-405d-4ef6-817e-eb679084fbe0").termSets.get().then((resp2: ITermSet[]) => {
          resp2.forEach((ele: ITermSet) => {
            ele.terms.select('Name', 'Id').inBatch(tbatch).get().then((resp3: ITerm[]) => {
              resp3.forEach((t: ITerm) => {
                let ip1 = {
                  parent: ele['Name'],
                  name: t['Name'],
                  id: t['Id'].replace("/Guid(", "").replace(")/", "")
                };
                alert(ip1.parent);
                //if(ip1.parent = this.state.division){
                //  tms.push(ip1);
                //}
              });
            });
          });
          tbatch.execute().then(_r => {
            resolve(tms);
          });
        });
      });
    });
  }
*/

/*  
  public componentDidMount() {
    this.getTermsetWithChildren().then((resp: IPTerm[]) => {
      console.log(resp);
      this.setState({
        terms: resp
      });
    });
  }
*/

  public componentWillMount() {
    taxonomy.getDefaultSiteCollectionTermStore().getTermSetById('6a488843-cebc-47b7-8606-e4e165e6fceb')
      .terms.get().then(Allterms => {
          console.log(Allterms);
          //this.setState({divisions : Allterms})
        }
      )
  }  

  public render(): React.ReactElement<IReactTaxonomyProps> {
    //let termrow = this.state.terms.map((t: IPTerm) => {
    //  return <tr> <td>{t.name}</td><td>{t.id}</td><td>{t.parent}</td></tr>;
    //});
    return (      
      <div className={styles.reactTaxonomy}>        
        <div className={styles.container}>                  
            <div className={styles.row}>
              <TaxonomyPicker allowMultipleSelections={false}
                initialValues={this.state.divisionTag}
                termsetNameOrID="Divisions"
                panelTitle="Select Division"
                label="Division"
                context={this.props.context}
                onChange={this.onDivisionTaxPickerChange}
                isTermSetSelectable={false} />
            </div>    
            <br/>
            <div className={ styles.title }>division={this.state.division}</div>
            <br/>
            <br/>
            <div className={styles.row}>
              <TaxonomyPicker allowMultipleSelections={false}
                initialValues={this.state.officeTag}
                termsetNameOrID={this.state.division}
                panelTitle="Select Office"
                label="Office"
                context={this.props.context}
                onChange={this.onOfficeTaxPickerChange}
                isTermSetSelectable={false} />
            </div>  
            <br/>
            <br/>
            <div className="col"><h5>Enter Details</h5></div>
            <div className="col">
              <input type="text" aria-label="text details" className="form-control" placeholder="Enter Text" id="textdetails"/>
            </div>         
        </div>
      </div>);
  }

/*

            <ComboBox 
                id="divisionDropdown"
                className="dropdown mr-1" 
                placeholder="Please Choose"
                selectedKey={this.state.divisionSelect}
                label="Division"
                autoComplete="on"
                options={this.props.divisionTerms}
                onChange={this.onDivisionChange}
              /> 
            <br/> 

            <div className={styles.column}>
              <span className={styles.title}>Terms from TermStore</span>
            </div>
          </div>  
          <table>        
            <thead><tr><th>Name</th><th>Id</th><th>Parent</th></tr></thead>            
            <tbody>{termrow}</tbody>
          </table>

*/

  //            <button type="button" className="btn btn-primary" onClick={this.createItem}>Submit</button> 

  private async onDivisionTaxPickerChange(terms: IPickerTerms) {
    const data = {};
    data['Tags'] = {
      "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
      "Label": terms[0].name,
      'TermGuid': terms[0].key,
      'WssId': '-1'
    };
    this.setState({division:terms[0].name + " Offices"});   
    alert('data ' + terms[0].name);
    return await sp.web.lists.getByTitle("Test_Metadata").items.add({
      Title: terms[0].name
    });  //getById(1).update(data);
  }  

  private async onOfficeTaxPickerChange(terms: IPickerTerms) {
    const data = {};
    data['Tags'] = {
      "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
      "Label": terms[0].name,
      'TermGuid': terms[0].key,
      'WssId': '-1'
    };
    alert('data written ' + terms[0].name);
    this.setState({division:terms[0].name});   
    return await sp.web.lists.getByTitle("Test_Metadata").items.add({
      Title: terms[0].name
    });  //getById(1).update(data);
  }   

  public onDivisionChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({ divisionSelect: option.key });
  }   

/*  
  private async createItem() {  
    //alert('data written');
    try {  
      await sp.web.lists.getByTitle('Test_Metadata').items.add({  
        Division: this.state.tags  
      });  
      this.setState({  
        //message: "Item: " + this.state.tags + " - created successfully!",  
        //showMessageBar: true,  
        //messageType: MessageBarType.success  
      });  
    }  
    catch (error) {  
      this.setState({  
        message: "Item " + this.state.tags + " creation failed with error: " + error,  
        showMessageBar: true,  
        messageType: MessageBarType.error  
      });  
    }  
  } 
  */ 
          //<select id="dropdown" onChange={this.handleDropdownChange}>
          //{ this.state.terms.map(term => {
          //    return <option value={term.Name.toLowerCase()}>{term.Name}</option>
          // })
          //}
          //</select>        

/*
  public render(): React.ReactElement<IReactTaxonomyProps> {
    return (
      <div className={ styles.reactTaxonomy }>
        <div className={ styles.container }>
          <div className={ styles.row }>

            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
*/

}
