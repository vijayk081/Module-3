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
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'EmployeeListingWebPartStrings';
import EmployeeListing from './components/EmployeeListing';
import { sp } from "@pnp/sp/presets/all";
var EmployeeListingWebPart = /** @class */ (function (_super) {
    __extends(EmployeeListingWebPart, _super);
    function EmployeeListingWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    EmployeeListingWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            sp.setup({
                spfxContext: _this.context
            });
        });
    };
    EmployeeListingWebPart.prototype.render = function () {
        var element = React.createElement(EmployeeListing, {
            description: this.properties.description,
            spfxcontext: this.context
        });
        ReactDom.render(element, this.domElement);
    };
    EmployeeListingWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(EmployeeListingWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    EmployeeListingWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return EmployeeListingWebPart;
}(BaseClientSideWebPart));
export default EmployeeListingWebPart;
//# sourceMappingURL=EmployeeListingWebPart.js.map