declare interface IDynamicFormularGeneratorWebPartStrings {
  PropertyPaneDescription: string;
  GroupListViewData: string;
  DataListSourceLabel: string;
  DataListSourceCurrentLabel: string;
  DataListSourceExternLabel: string;
  ChooseList: string;
  ChooseView: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SuccessMessageLabel: string;
  NoOfFileUploads: string;
  FieldRulesLabel: string;
  URLOfExternalSitePlaceholderLabel: string;
  ErrorMissingSiteText: string;
  AllowedUploadFileTypesLabel: string;
  EmailSubjectLable: string;
  EmailHeaderLabel: string;
  SendEMailWithFormDataLabel: string;
  SendEMailWithFormDataYesLabel: string;
  SendEMailWithFormDataNoLabel: string;
  AddDataLinkToEMailLabel: string;
  AttachmentLabel: string;
  AttachmentIndexLabel: string;
  EnablePrintLabel: string;

  GroupMiscSettings: string;
  DateTimeFieldLabel: string;
  FormValidFromFieldLabel: string;
  FormValidToFieldLabel: string;

  LabelYES: string,
  LabelNO: string,

  VALMsgRequiredField: string;
  VALMsgInvalidFieldData: string;
  VALMsgOnlyNumbersAllowed: string;
  VALMsgDecimalInvalid: string;
  VALMsgvalueRangeOverflow: string;

  CFGHeader: string;
  CFGChooseList: string;
  CFGChooseView: string;
  CFGBTNConfigure: string;

  ChooseContentType: string;

  MAILLinkTodata: string;
  MSGWaiting: string;
  BTNSendFormData: string;
  BTNPrintFormData: string;
  BTNResetFormData: string;
  HEADPrintForm: string;

  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;

  MSGDataSendAlready: string;
}

declare module 'DynamicFormularGeneratorWebPartStrings' {
  const strings: IDynamicFormularGeneratorWebPartStrings;
  export = strings;
}
