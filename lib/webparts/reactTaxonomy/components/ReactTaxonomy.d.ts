import * as React from 'react';
import { IReactTaxonomyProps } from './IReactTaxonomyProps';
import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
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
}
export default class ReactTaxonomy extends React.Component<IReactTaxonomyProps, IReactTaxonomyState> {
    constructor(props: any);
    private _gettags;
    componentWillMount(): void;
    render(): React.ReactElement<IReactTaxonomyProps>;
    onDivisionTaxPickerChange(terms: IPickerTerms): Promise<any>;
    private onOfficeTaxPickerChange;
    cleanGuid(guid: string): string;
}
//# sourceMappingURL=ReactTaxonomy.d.ts.map