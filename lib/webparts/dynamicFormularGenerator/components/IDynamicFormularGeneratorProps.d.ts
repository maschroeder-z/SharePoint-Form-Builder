import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRuleEntry } from '../../../Common/IRuleEntry';
export interface IDynamicFormularGeneratorProps {
    description: string;
    isDarkTheme: boolean;
    hasTeamsContext: boolean;
    userDisplayName: string;
    viewID: string;
    listID: string;
    viewXml: string;
    httpClient: SPHttpClient;
    siteURL: string;
    successMessage: string;
    uploads: number;
    allowedUploadFileTypes: string;
    addionalFieldRules: {
        [key: string]: IRuleEntry;
    };
    emailSubject: string;
    emailLeadText: string;
    currentUserEMail: string;
    sendConfirmationEMail: boolean;
    addDataLinkInEMail: boolean;
    enablePrint: boolean;
    wpContext: WebPartContext;
}
//# sourceMappingURL=IDynamicFormularGeneratorProps.d.ts.map