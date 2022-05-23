import * as React from 'react';
import { IReactTaxonomyProps } from './IReactTaxonomyProps';
import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IComboBoxOption, IComboBox } from 'office-ui-fabric-react/lib/index';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
export interface DivisionTerm {
    id: string;
    name: string;
}
export interface IReactTaxonomyState {
    divisionTag: IPickerTerms;
    officeTag: IPickerTerms;
    division: string;
    divisionSelect: any;
}
export default class ReactTaxonomy extends React.Component<IReactTaxonomyProps, IReactTaxonomyState> {
    constructor(props: any);
    private _gettags;
    componentWillMount(): void;
    render(): React.ReactElement<IReactTaxonomyProps>;
    private onDivisionTaxPickerChange;
    private onOfficeTaxPickerChange;
    onDivisionChange: (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption) => void;
}
//# sourceMappingURL=ReactTaxonomy%20copy.d.ts.map