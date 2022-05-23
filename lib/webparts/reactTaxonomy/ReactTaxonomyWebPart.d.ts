import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReactTaxonomyProps } from './components/IReactTaxonomyProps';
export default class ReactTaxonomyWebPart extends BaseClientSideWebPart<IReactTaxonomyProps> {
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=ReactTaxonomyWebPart.d.ts.map