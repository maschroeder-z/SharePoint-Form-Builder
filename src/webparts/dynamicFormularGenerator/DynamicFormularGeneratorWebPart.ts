import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient } from '@microsoft/sp-http';
import { Guid, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneGroup,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'DynamicFormularGeneratorWebPartStrings';
import DynamicFormularGenerator from './components/DynamicFormularGenerator';
import { IDynamicFormularGeneratorProps } from './components/IDynamicFormularGeneratorProps';
import { ISPLists } from '../../Common/ISPLists';
import { ISPView, ISPViews } from '../../Common/ISPListViews';
import { FluentProvider, FluentProviderProps, teamsDarkTheme, teamsLightTheme, webLightTheme, webDarkTheme, Theme } from '@fluentui/react-components';
import { Helper } from '../../Common/Helper';
import { PropertyPaneFieldRuleEditor } from '../Controls/PropertyPaneFieldRuleEditor';
import { IRuleEntry } from '../../Common/IRuleEntry';

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

export interface IDynamicFormularGeneratorWebPartProps {
  description: string;
  successMessage: string;
  currentSite: boolean;
  siteUrl: string;
  sourceListName: string;
  viewID: string;
  viewXML: string;
  emailToUser: boolean;
  attachmentFields: number;
  allowedUploadFileTypes: string;
  addionalFieldRules: { [key: string]: IRuleEntry };
  emailSubject: string;
  emailHeader: string;
  addDataLinkInEMail: boolean;
  enablePrint: boolean;
}

export default class DynamicFormularGeneratorWebPart extends BaseClientSideWebPart<IDynamicFormularGeneratorWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme: Theme = webLightTheme;

  private availableLists: IPropertyPaneDropdownOption[] = [];
  private viewsInList: IPropertyPaneDropdownOption[] = [];
  private viewData: ISPView[] = null;
  private fieldsInView: IPropertyPaneDropdownOption[] = null;
  private loadingLists: boolean = false;

  public render(): void {
    const element: React.ReactElement<IDynamicFormularGeneratorProps> = React.createElement(
      DynamicFormularGenerator,
      {
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
      }
    );

    //wrap the component with the Fluent UI 9 Provider.
    const fluentElement: React.ReactElement<FluentProviderProps> = React.createElement(
      FluentProvider,
      {
        theme: this._appMode === AppMode.Teams || this._appMode === AppMode.TeamsLocal ?
          this._isDarkTheme ? teamsDarkTheme : teamsLightTheme :
          this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal ?
            this._isDarkTheme ? webDarkTheme : this._theme :
            this._isDarkTheme ? webDarkTheme : webLightTheme
      },
      element
    );

    ReactDom.render(fluentElement, this.domElement);
  }

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

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private GetSelectedUrl(): string {
    return this.properties.siteUrl ? this.properties.siteUrl : this.context.pageContext.web.absoluteUrl;
  }

  private loadWPConfigInformation(): void {
    if (!this.loadingLists && this.availableLists.length === 0) {
      this.loadAvailableLists();
    }
    return;
  }

  private qryListInformation(): Promise<ISPLists> {
    const endpoint: string = `${this.GetSelectedUrl()}/_api/web/lists/`;
    return this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then(response => {
        return response.json();
      });
  }
  private qryViews4List(listID: Guid): Promise<ISPViews> {
    const endpoint = `${this.GetSelectedUrl()}/_api/web/lists/getbyid('${listID}')/views`;
    return this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then(response => {
        return response.json();
      });
  }

  private async loadAvailableLists(): Promise<void> {
    this.loadingLists = true;
    let options: IPropertyPaneDropdownOption[] = [];
    try {
      const lists: ISPLists = await this.qryListInformation();
      if (lists.value.length > 0) {
        lists.value.forEach(libs => {
          options.push({ key: libs.Id, text: `${libs.Title} (${libs.ItemCount})` });
        });
        if (typeof (this.properties.sourceListName) !== "undefined" && this.properties.sourceListName.trim().length > 0) {
          const viewData: ISPViews = await this.qryViews4List(Guid.parse(this.properties.sourceListName));
          this.viewsInList = [];
          this.viewData = viewData.value;
          viewData.value.forEach(item => {
            if (!item.Hidden) {
              this.viewsInList.push({
                key: item.Id,
                text: item.Title
              });
            }
          });
          this.render();
          this.context.propertyPane.refresh();
          this.onPropertyPaneFieldChanged("viewID", null, this.properties.viewID);
        }
      }
      else {
        this.loadingLists = false;
        options = [];
      }
    }
    catch (error) {
      this.loadingLists = false;
      options = [];
    }
    this.loadingLists = false;
    this.availableLists = options;
    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'currentSite' && newValue) {
      super.onPropertyPaneFieldChanged("siteUrl", oldValue, "");
    }
    if (propertyPath === 'sourceListName' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      delete this.properties.viewXML;
      this.viewsInList = [];
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'viewsInList');
      this.qryViews4List(newValue).then((viewList: ISPViews): void => {
        const newListViews: IPropertyPaneDropdownOption[] = [];
        this.viewData = viewList.value;
        viewList.value.forEach(item => {
          if (!item.Hidden) {
            newListViews.push({
              key: item.Id,
              text: item.Title
            });
          }
        });
        this.viewsInList = newListViews;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.context.propertyPane.refresh();
      });
    }
    if (propertyPath === 'viewID' && newValue && Guid.tryParse(newValue) && this.viewData) {
      const temp = this.viewData.filter(x => x.Id === newValue)[0];
      if (typeof temp !== "undefined" && typeof temp.ListViewXml !== "undefined") {
        this.properties.viewXML = temp.ListViewXml;
        this.fieldsInView = [];
        Helper.GetViewFields(this.properties.viewXML).forEach(fieldName => {
          this.fieldsInView.push({ key: fieldName, text: fieldName });
        });
      }
    }
  }

  protected get getSourceConfiguration(): IPropertyPaneGroup {

    const grp: IPropertyPaneGroup = {
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
          placeholder: `${this.properties.currentSite ? this.context.pageContext.web.absoluteUrl : strings.URLOfExternalSitePlaceholderLabel}`,
          disabled: this.properties.currentSite,
          onGetErrorMessage: (value: string): string => {
            if (!this.properties.currentSite && (value === null || value.trim().length === 0)) {
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
          onPropertyChange: (fieldRules) => {
            console.log(fieldRules); // TO JSON AND save
            this.properties.addionalFieldRules = fieldRules;
          },
        })
      ]
    };
    return grp;
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.loadWPConfigInformation();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }
}
