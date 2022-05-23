import * as React from 'react';
import styles from './ReactTaxonomy.module.scss';
import { IReactTaxonomyProps } from './IReactTaxonomyProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { taxonomy, ITerm, ITermSet, ITermStore } from '@pnp/sp-taxonomy';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { MessageBar, MessageBarType, IStackProps, Stack } from 'office-ui-fabric-react';  
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

const verticalStackProps: IStackProps = {  
  styles: { root: { overflow: 'hidden', width: '100%' } },  
  tokens: { childrenGap: 20 }  
};

let officeHTML;

export interface IPTerm {
  parent?: string;
  id: string;
  name: string;
}

export interface DivisionTerm {
  id: string;
  name: string;
}
 
export interface IReactTaxonomyState {
  terms: IPTerm[];
  tags: IPickerTerms;
  //officeHTML : string;
  //showMessageBar: boolean;    
  //messageType?: MessageBarType;    
  //message?: string;
}

export default class ReactTaxonomy extends React.Component<IReactTaxonomyProps, IReactTaxonomyState> {
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      terms: [],
      tags: []
      //officeHTML: ""
      //showMessageBar: false
    };
    this._gettags();
  }

  private async _gettags() {
    const item: any = await sp.web.lists.getByTitle("Test_Metadata").items.getById(1).get();
    let selectedtags: any = [];
    item.Tags.forEach(function (v: any[], i) {
      selectedtags.push({ key: v["TermGuid"], name: v["Label"] })
    });
    console.log(item);
    this.setState({
      tags: selectedtags
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
            <div>
              <TaxonomyPicker allowMultipleSelections={false}
                initialValues={this.state.tags}
                termsetNameOrID="Divisions"
                panelTitle="Select Division"
                label="Division"
                context={this.props.context}
                onChange={this.onDivisionTaxPickerChange}
                isTermSetSelectable={false} />
            </div>    
            <br/>
            <br/> 
            {alert("2:"+officeHTML)}
            {officeHTML}            
            <br/>
            <br/>
          </div>
        </div>
      </div>);
  }

/*
            <div dangerouslySetInnerHTML={{ __html:this.state.officeHTML}}/>  

            <div className={styles.column}>
              <span className={styles.title}>Terms from TermStore</span>
            </div>
          </div>  
          <table>        
            <thead><tr><th>Name</th><th>Id</th><th>Parent</th></tr></thead>            
            <tbody>{termrow}</tbody>
          </table>

*/

  //<button type="button" className="btn btn-primary" onClick={this.createItem}>Submit</button> 
  //public async updateMeta(term: (ITerm & ITermData), list: string, field: string, itemId: number): Promise<any> {
  
  public async onDivisionTaxPickerChange(terms: IPickerTerms): Promise<any> {
    const data = {};
    //let termLabel:string="";

    data['Tags'] = {
      "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
      "Label": terms[0].name,
      'TermGuid': terms[0].key,
      'WssId': '-1'
    };
    alert('data written ' + terms[0].key);
    //termLabel=terms[0].name;
/*    
    officeHTML=(<div><TaxonomyPicker allowMultipleSelections={false}
      initialValues={this.state.tags}
      termsetNameOrID={terms[0].name}
      panelTitle="Select Office"
      label="Office"
      context={this.props.context}
      onChange={this.onOfficeTaxPickerChange}
      isTermSetSelectable={false} /></div>);
    
    alert("1:"+officeHTML);

    this.setState({officeHTML:htmlstring});   
*/
    return await sp.web.lists.getByTitle("Test_Metadata").items.getById(4).update({Title: terms[0].name});
    //return (await sp.web.lists.getByTitle("Test_Metadata").items.add({Title: termLabel}))
    //getById(1).update(data);
  }  

  private async onOfficeTaxPickerChange(terms: IPickerTerms) {
    const data = {};
    data['Tags'] = {
      "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
      "Label": terms[0].name,
      'TermGuid': terms[0].key,
      'WssId': '-1'
    };
    alert('data written ' + data);
    //this.setState({division:terms[0].name});   
    //return await sp.web.lists.getByTitle(list).items.getById(itemId).update(data);

    return await sp.web.lists.getByTitle("Test_Metadata").items.add({
      Division: data
    });  //getById(1).update(data);
  }  

  public cleanGuid(guid: string): string {
    if (guid !== undefined) {
        return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    } else {
        return '';
    }
  }   
/*
  const data = {};
  data[field] = {
    "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
    "Label": term.Name,
    'TermGuid': this.cleanGuid(term.Id),
    'WssId': '-1'
  };

  return await sp.web.lists.getByTitle(list).items.getById(itemId).update(data);
*/

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
