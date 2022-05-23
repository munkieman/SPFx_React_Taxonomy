var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './ReactTaxonomy.module.scss';
import { taxonomy } from '@pnp/sp-taxonomy';
import { TaxonomyPicker } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
var verticalStackProps = {
    styles: { root: { overflow: 'hidden', width: '100%' } },
    tokens: { childrenGap: 20 }
};
var officeHTML;
var ReactTaxonomy = /** @class */ (function (_super) {
    __extends(ReactTaxonomy, _super);
    function ReactTaxonomy(props) {
        var _this = _super.call(this, props) || this;
        sp.setup({
            spfxContext: _this.props.context
        });
        _this.state = {
            terms: [],
            tags: []
            //officeHTML: ""
            //showMessageBar: false
        };
        _this._gettags();
        return _this;
    }
    ReactTaxonomy.prototype._gettags = function () {
        return __awaiter(this, void 0, void 0, function () {
            var item, selectedtags;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, sp.web.lists.getByTitle("Test_Metadata").items.getById(1).get()];
                    case 1:
                        item = _a.sent();
                        selectedtags = [];
                        item.Tags.forEach(function (v, i) {
                            selectedtags.push({ key: v["TermGuid"], name: v["Label"] });
                        });
                        console.log(item);
                        this.setState({
                            tags: selectedtags
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
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
    ReactTaxonomy.prototype.componentWillMount = function () {
        taxonomy.getDefaultSiteCollectionTermStore().getTermSetById('6a488843-cebc-47b7-8606-e4e165e6fceb')
            .terms.get().then(function (Allterms) {
            console.log(Allterms);
            //this.setState({divisions : Allterms})
        });
    };
    ReactTaxonomy.prototype.render = function () {
        //let termrow = this.state.terms.map((t: IPTerm) => {
        //  return <tr> <td>{t.name}</td><td>{t.id}</td><td>{t.parent}</td></tr>;
        //});
        return (React.createElement("div", { className: styles.reactTaxonomy },
            React.createElement("div", { className: styles.container },
                React.createElement("div", { className: styles.row },
                    React.createElement("div", null,
                        React.createElement(TaxonomyPicker, { allowMultipleSelections: false, initialValues: this.state.tags, termsetNameOrID: "Divisions", panelTitle: "Select Division", label: "Division", context: this.props.context, onChange: this.onDivisionTaxPickerChange, isTermSetSelectable: false })),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    alert("2:" + officeHTML),
                    officeHTML,
                    React.createElement("br", null),
                    React.createElement("br", null)))));
    };
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
    ReactTaxonomy.prototype.onDivisionTaxPickerChange = function (terms) {
        return __awaiter(this, void 0, void 0, function () {
            var data, termLabel;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        data = {};
                        termLabel = "";
                        data['Tags'] = {
                            "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
                            "Label": terms[0].name,
                            'TermGuid': terms[0].key,
                            'WssId': '-1'
                        };
                        alert('data written ' + terms[0].key);
                        termLabel = terms[0].name;
                        return [4 /*yield*/, sp.web.lists.getByTitle("Test_Metadata").items.getById(4).update({ Title: terms[0].name })];
                    case 1: 
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
                    return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    ReactTaxonomy.prototype.onOfficeTaxPickerChange = function (terms) {
        return __awaiter(this, void 0, void 0, function () {
            var data;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        data = {};
                        data['Tags'] = {
                            "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
                            "Label": terms[0].name,
                            'TermGuid': terms[0].key,
                            'WssId': '-1'
                        };
                        alert('data written ' + data);
                        return [4 /*yield*/, sp.web.lists.getByTitle("Test_Metadata").items.add({
                                Division: data
                            })];
                    case 1: 
                    //this.setState({division:terms[0].name});   
                    //return await sp.web.lists.getByTitle(list).items.getById(itemId).update(data);
                    return [2 /*return*/, _a.sent()]; //getById(1).update(data);
                }
            });
        });
    };
    ReactTaxonomy.prototype.cleanGuid = function (guid) {
        if (guid !== undefined) {
            return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
        }
        else {
            return '';
        }
    };
    return ReactTaxonomy;
}(React.Component));
export default ReactTaxonomy;
//# sourceMappingURL=ReactTaxonomy.js.map