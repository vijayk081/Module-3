import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IEmployeeListingWebPartProps {
    description: string;
}
export default class EmployeeListingWebPart extends BaseClientSideWebPart<IEmployeeListingWebPartProps> {
    protected onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=EmployeeListingWebPart.d.ts.map