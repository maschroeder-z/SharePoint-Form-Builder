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
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
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
import * as ReactDom from 'react-dom';
import { SPHttpClient } from '@microsoft/sp-http';
import { Guid, Version } from '@microsoft/sp-core-library';
import { PropertyPaneCheckbox, PropertyPaneDropdown, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'DynamicFormularGeneratorWebPartStrings';
import DynamicFormularGenerator from './components/DynamicFormularGenerator';
import { FluentProvider, teamsDarkTheme, teamsLightTheme, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import { Helper } from '../../Common/Helper';
import { PropertyPaneFieldRuleEditor } from '../Controls/PropertyPaneFieldRuleEditor';
export var AppMode;
(function (AppMode) {
    AppMode[AppMode["SharePoint"] = 0] = "SharePoint";
    AppMode[AppMode["SharePointLocal"] = 1] = "SharePointLocal";
    AppMode[AppMode["Teams"] = 2] = "Teams";
    AppMode[AppMode["TeamsLocal"] = 3] = "TeamsLocal";
    AppMode[AppMode["Office"] = 4] = "Office";
    AppMode[AppMode["OfficeLocal"] = 5] = "OfficeLocal";
    AppMode[AppMode["Outlook"] = 6] = "Outlook";
    AppMode[AppMode["OutlookLocal"] = 7] = "OutlookLocal";
})(AppMode || (AppMode = {}));
var DynamicFormularGeneratorWebPart = /** @class */ (function (_super) {
    __extends(DynamicFormularGeneratorWebPart, _super);
    function DynamicFormularGeneratorWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._appMode = AppMode.SharePoint;
        _this._theme = webLightTheme;
        _this.availableLists = [];
        _this.viewsInList = [];
        _this.viewData = null;
        _this.fieldsInView = null;
        _this.loadingLists = false;
        return _this;
    }
    DynamicFormularGeneratorWebPart.prototype.render = function () {
        var element = React.createElement(DynamicFormularGenerator, {
            description: this.properties.description,
            isDarkTheme: this._isDarkTheme,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName,
            viewID: this.properties.viewID,
            listID: this.properties.sourceListName,
            httpClient: this.context.spHttpClient,
            viewXml: (this.properties.viewXML !== null ? this.properties.viewXML : ""),
            siteURL: this.GetSelectedUrl(),
            successMessage: this.properties.successMessage,
            uploads: this.properties.attachmentFields,
            allowedUploadFileTypes: this.properties.allowedUploadFileTypes,
            addionalFieldRules: this.properties.addionalFieldRules,
            emailSubject: this.properties.emailSubject,
            emailLeadText: this.properties.emailHeader,
            currentUserEMail: this.context.pageContext.user.email,
            sendConfirmationEMail: this.properties.emailToUser,
            addDataLinkInEMail: this.properties.addDataLinkInEMail,
            enablePrint: this.properties.enablePrint,
            wpContext: this.context,
        });
        //wrap the component with the Fluent UI 9 Provider.
        var fluentElement = React.createElement(FluentProvider, {
            theme: this._appMode === AppMode.Teams || this._appMode === AppMode.TeamsLocal ?
                this._isDarkTheme ? teamsDarkTheme : teamsLightTheme :
                this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal ?
                    this._isDarkTheme ? webDarkTheme : this._theme :
                    this._isDarkTheme ? webDarkTheme : webLightTheme
        }, element);
        ReactDom.render(fluentElement, this.domElement);
    };
    /*protected onInit(): Promise<void> {
      return this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      });
    }*/
    /*private _getEnvironmentMessage(): Promise<string> {
      if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
        return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
          .then(context => {
            let environmentMessage: string = '';
            switch (context.app.host.name) {
              case 'Office': // running in Office
                environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                break;
              case 'Outlook': // running in Outlook
                environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                break;
              case 'Teams': // running in Teams
                environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                break;
              default:
                throw new Error('Unknown host');
            }
  
            return environmentMessage;
          });
      }
  
      return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    }*/
    DynamicFormularGeneratorWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    DynamicFormularGeneratorWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(DynamicFormularGeneratorWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    DynamicFormularGeneratorWebPart.prototype.GetSelectedUrl = function () {
        return this.properties.siteUrl ? this.properties.siteUrl : this.context.pageContext.web.absoluteUrl;
    };
    DynamicFormularGeneratorWebPart.prototype.loadWPConfigInformation = function () {
        if (!this.loadingLists && this.availableLists.length === 0) {
            this.loadAvailableLists();
        }
        return;
    };
    DynamicFormularGeneratorWebPart.prototype.qryListInformation = function () {
        var endpoint = "".concat(this.GetSelectedUrl(), "/_api/web/lists/");
        return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        });
    };
    DynamicFormularGeneratorWebPart.prototype.qryViews4List = function (listID) {
        var endpoint = "".concat(this.GetSelectedUrl(), "/_api/web/lists/getbyid('").concat(listID, "')/views");
        return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        });
    };
    DynamicFormularGeneratorWebPart.prototype.loadAvailableLists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var options, lists, viewData, error_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.loadingLists = true;
                        options = [];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 7, , 8]);
                        return [4 /*yield*/, this.qryListInformation()];
                    case 2:
                        lists = _a.sent();
                        if (!(lists.value.length > 0)) return [3 /*break*/, 5];
                        lists.value.forEach(function (libs) {
                            options.push({ key: libs.Id, text: "".concat(libs.Title, " (").concat(libs.ItemCount, ")") });
                        });
                        if (!(typeof (this.properties.sourceListName) !== "undefined" && this.properties.sourceListName.trim().length > 0)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.qryViews4List(Guid.parse(this.properties.sourceListName))];
                    case 3:
                        viewData = _a.sent();
                        this.viewsInList = [];
                        this.viewData = viewData.value;
                        viewData.value.forEach(function (item) {
                            if (!item.Hidden) {
                                _this.viewsInList.push({
                                    key: item.Id,
                                    text: item.Title
                                });
                            }
                        });
                        this.render();
                        this.context.propertyPane.refresh();
                        this.onPropertyPaneFieldChanged("viewID", null, this.properties.viewID);
                        _a.label = 4;
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        this.loadingLists = false;
                        options = [];
                        _a.label = 6;
                    case 6: return [3 /*break*/, 8];
                    case 7:
                        error_1 = _a.sent();
                        this.loadingLists = false;
                        options = [];
                        return [3 /*break*/, 8];
                    case 8:
                        this.loadingLists = false;
                        this.availableLists = options;
                        this.context.propertyPane.refresh();
                        return [2 /*return*/];
                }
            });
        });
    };
    DynamicFormularGeneratorWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath, oldValue, newValue) {
        var _this = this;
        if (propertyPath === 'currentSite' && newValue) {
            _super.prototype.onPropertyPaneFieldChanged.call(this, "siteUrl", oldValue, "");
        }
        if (propertyPath === 'sourceListName' && newValue) {
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
            delete this.properties.viewXML;
            this.viewsInList = [];
            this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'viewsInList');
            this.qryViews4List(newValue).then(function (viewList) {
                var newListViews = [];
                _this.viewData = viewList.value;
                viewList.value.forEach(function (item) {
                    if (!item.Hidden) {
                        newListViews.push({
                            key: item.Id,
                            text: item.Title
                        });
                    }
                });
                _this.viewsInList = newListViews;
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                _this.context.propertyPane.refresh();
            });
        }
        if (propertyPath === 'viewID' && newValue && Guid.tryParse(newValue) && this.viewData) {
            var temp = this.viewData.filter(function (x) { return x.Id === newValue; })[0];
            if (typeof temp !== "undefined" && typeof temp.ListViewXml !== "undefined") {
                this.properties.viewXML = temp.ListViewXml;
                this.fieldsInView = [];
                Helper.GetViewFields(this.properties.viewXML).forEach(function (fieldName) {
                    _this.fieldsInView.push({ key: fieldName, text: fieldName });
                });
            }
        }
    };
    Object.defineProperty(DynamicFormularGeneratorWebPart.prototype, "getSourceConfiguration", {
        get: function () {
            var _this = this;
            var grp = {
                groupName: strings.GroupListViewData,
                groupFields: [
                    PropertyPaneToggle('currentSite', {
                        label: strings.DataListSourceLabel,
                        onText: strings.DataListSourceCurrentLabel,
                        offText: strings.DataListSourceExternLabel,
                        offAriaLabel: strings.DataListSourceExternLabel,
                        onAriaLabel: strings.DataListSourceCurrentLabel,
                        checked: true
                    }),
                    PropertyPaneTextField('siteUrl', {
                        underlined: true,
                        placeholder: "".concat(this.properties.currentSite ? this.context.pageContext.web.absoluteUrl : strings.URLOfExternalSitePlaceholderLabel),
                        disabled: this.properties.currentSite,
                        onGetErrorMessage: function (value) {
                            if (!_this.properties.currentSite && (value === null || value.trim().length === 0)) {
                                return strings.ErrorMissingSiteText;
                            }
                            return "";
                        }
                    }),
                    PropertyPaneDropdown('sourceListName', {
                        options: this.availableLists,
                        label: strings.ChooseList,
                        disabled: this.availableLists.length === 0,
                        selectedKey: this.properties.sourceListName
                    }),
                    PropertyPaneDropdown('viewID', {
                        options: this.viewsInList,
                        label: strings.ChooseView,
                        disabled: this.viewsInList.length === 0,
                        selectedKey: this.properties.viewID,
                    }),
                    new PropertyPaneFieldRuleEditor('fieldRulesSettings', {
                        label: strings.FieldRulesLabel,
                        disabled: false,
                        stateKey: new Date().toString(),
                        fieldNames: this.fieldsInView,
                        addionalFieldRules: this.properties.addionalFieldRules,
                        onPropertyChange: function (fieldRules) {
                            console.log(fieldRules); // TO JSON AND save
                            _this.properties.addionalFieldRules = fieldRules;
                        },
                    })
                ]
            };
            return grp;
        },
        enumerable: false,
        configurable: true
    });
    DynamicFormularGeneratorWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        this.loadWPConfigInformation();
    };
    DynamicFormularGeneratorWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    displayGroupsAsAccordion: true,
                    groups: [
                        this.getSourceConfiguration,
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel,
                                    multiline: true,
                                    resizable: true
                                }),
                                PropertyPaneTextField('successMessage', {
                                    label: strings.SuccessMessageLabel,
                                    multiline: true,
                                    resizable: true
                                }),
                                PropertyPaneCheckbox('enablePrint', {
                                    text: strings.EnablePrintLabel
                                }),
                                PropertyPaneToggle('emailToUser', {
                                    label: strings.SendEMailWithFormDataLabel,
                                    onText: strings.SendEMailWithFormDataYesLabel,
                                    offText: strings.SendEMailWithFormDataNoLabel,
                                    checked: false
                                }),
                                PropertyPaneTextField('emailSubject', {
                                    label: strings.EmailSubjectLable,
                                    disabled: !this.properties.emailToUser,
                                    multiline: false
                                }),
                                PropertyPaneTextField('emailHeader', {
                                    label: strings.EmailHeaderLabel,
                                    disabled: !this.properties.emailToUser,
                                    multiline: true,
                                    resizable: true
                                }),
                                PropertyPaneCheckbox('addDataLinkInEMail', {
                                    text: strings.AddDataLinkToEMailLabel,
                                    disabled: !this.properties.emailToUser
                                }),
                                PropertyPaneTextField('allowedUploadFileTypes', {
                                    label: strings.AllowedUploadFileTypesLabel,
                                    multiline: false
                                }),
                                PropertyPaneSlider('attachmentFields', {
                                    label: strings.NoOfFileUploads,
                                    min: 0,
                                    max: 3,
                                    value: 0,
                                    showValue: true,
                                    step: 1
                                })
                            ]
                        }
                    ]
                },
            ]
        };
    };
    return DynamicFormularGeneratorWebPart;
}(BaseClientSideWebPart));
export default DynamicFormularGeneratorWebPart;
//# sourceMappingURL=DynamicFormularGeneratorWebPart.js.map