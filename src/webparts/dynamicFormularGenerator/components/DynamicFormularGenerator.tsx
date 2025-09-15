import * as React from 'react';
import styles from './DynamicFormularGenerator.module.scss';
import { IDynamicFormularGeneratorProps } from './IDynamicFormularGeneratorProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import {  ChoiceValue, ISPListField, ISPListFields} from '../../../Common/ISPListFields'; //await this.qryListFields(this.props.listID);    
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { FormControlFluentUI } from '../../../Common/FormControlFluentUI';
import { FieldTypes } from '../../../Common/FieldTypes';
import { Helper } from '../../../Common/Helper';
import { Button, Spinner } from '@fluentui/react-components';
import { ISPListItem } from '../../../Common/ISPListItem';
import { ISPFormLists } from '../../../Common/ISPFormLists';
import * as strings from 'DynamicFormularGeneratorWebPartStrings';
import { FlashSettings24Regular } from '@fluentui/react-icons';
import { LinkFieldValue } from '../../../Common/LinkFieldValue';

type FormState = {
  errorMessage: string[];
  isFormValid: boolean;
  isProcessing: boolean;
  isAlreadySent: boolean;
  formFields: string[];    
}

export default class DynamicFormularGenerator extends React.Component<IDynamicFormularGeneratorProps, FormState> {
  private availableFields : ISPListFields = null;
  private currentViewXML : string = "";
  private currentListID : string = "";  
  private parser : DOMParser = null;  
  private attachmentCtl : React.ReactNode[] = null;
  private uploadFileList: {[key: string]: File} = {};    
  
  constructor (props:IDynamicFormularGeneratorProps) {
    super(props);
    this.state = {
      errorMessage: new Array<string>(),
      isFormValid : false,
      isProcessing: false,
      isAlreadySent: false,
      formFields: []     
    }
    this.parser = new DOMParser();
  }
  
  private getFieldSchemata(schemaXML:string) : Element
  {
    const xmlDoc = this.parser.parseFromString(schemaXML,"text/xml"); 
    return xmlDoc.getElementsByTagName("Field")[0];
  }
  private getAttributeValue(dom: Element, attributeToRead:string) : string
  {
    if (typeof dom !== "undefined" && dom !== null)
      return dom.getAttribute(attributeToRead);
    return "";
  }

  private async qryFormFields() : Promise<void>
  {    
    if (this.validateConfiguration() && (this.props.viewXml!==this.currentViewXML || this.currentListID!==this.props.listID))
    {
      this.currentViewXML=this.props.viewXml;
      this.currentListID=this.props.listID;
      this.availableFields = await this.qryListFields(this.props.listID);
      const temp : string = this.props.viewXml.replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(temp,"text/xml");    
      const tempFields : string[] = [];    // TODO: replace with helper method       
      xmlDoc.getElementsByTagName("ViewFields")[0].childNodes.forEach((node:HTMLElement, index)=> {        
          const fieldInfo : ISPListField = this.availableFields.value.filter(f=>f.StaticName===node.getAttribute("Name"))[0];
          fieldInfo.IsValid = !fieldInfo.Required;
          tempFields.push(node.getAttribute("Name"));        
      });
      this.setState({formFields: tempFields});        
    }
  } 

  private validateConfiguration() : boolean
  {
    return (typeof this.props.viewXml !== "undefined" && typeof this.props.listID !== "undefined");
  }

  private qryListFields(listID: string): Promise<ISPListFields> {
    const endpoint = `${this.props.siteURL}/_api/web/lists/getbyid('${listID}')/Fields`;
    return this.props.httpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
    .then((response: { json: () => any; }) => {
      return response.json();
    }); 
  }

  private getFieldMetaData(fieldInfo: ISPListField): ISPListField {
    const fieldSchemata = this.getFieldSchemata(fieldInfo.SchemaXml);
    fieldInfo.Decimals=0;    
    if (fieldInfo.FieldTypeKind === FieldTypes.NUMBER)
    {
      fieldInfo.Decimals = parseInt(this.getAttributeValue(fieldSchemata, "Decimals"), 10);
      if (fieldInfo.DefaultValue===null)
        fieldInfo.DefaultValue = "0";
    }
    if (fieldInfo.FieldTypeKind === FieldTypes.CHOICE || fieldInfo.FieldTypeKind === FieldTypes.MULTICHOICE){
      if (typeof fieldInfo.Choices !== "undefined" && fieldInfo.Choices.length>0) {
        fieldInfo.ChoiceUI=this.getAttributeValue(fieldSchemata, "Format");
      } 
    }
    if (fieldInfo.FieldTypeKind === FieldTypes.NOTE) {
      fieldInfo.IsRichTextAllowed = this.getAttributeValue(fieldSchemata, "RichText") === "True";
    }
    if (fieldInfo.FieldTypeKind === FieldTypes.URLORIMAGE) {
      fieldInfo.LinkUI = this.getAttributeValue(fieldSchemata, "Format");
    }
    if (fieldInfo.FieldTypeKind === FieldTypes.LOOKUP)
    {
      fieldInfo.LookupField = {
        DisplayName : this.getAttributeValue(fieldSchemata, "DisplayName"),
        FieldRef : this.getAttributeValue(fieldSchemata, "FieldRef"),
        ID : this.getAttributeValue(fieldSchemata, "ID"),
        List : this.getAttributeValue(fieldSchemata, "List"),
        Name : this.getAttributeValue(fieldSchemata, "Name"),
        ReadOnly : this.getAttributeValue(fieldSchemata, "ReadOnly")==="TRUE",
        ShowField : this.getAttributeValue(fieldSchemata, "ShowField"),
        StaticName : this.getAttributeValue(fieldSchemata, "StaticName"),
        WebId : this.getAttributeValue(fieldSchemata, "WebId"),
        LookupChoices: new Array<ChoiceValue>()          
      };
    }
    return fieldInfo;
  }

  private formComponentFactory(fieldStaticName:string) : React.ReactNode
  {
    if (this.availableFields !== null)
    {
      let fieldInfo : ISPListField = this.availableFields.value.filter(f=>f.StaticName===fieldStaticName)[0];
      if (fieldInfo === undefined)
        return null;     

      if (fieldInfo.IsDependentLookup)
        return;

      if (!fieldInfo.IsUsedInForm) // only once
      {
        if (fieldInfo.FieldTypeKind === FieldTypes.LOOKUP 
            && !fieldInfo.IsDependentLookup 
            && fieldInfo.DependentLookupInternalNames !== null
            && fieldInfo.DependentLookupInternalNames.length>0) 
        {        
          const lookupFieldInfo:string[] = [];
          fieldInfo.DependentLookupInternalNames.forEach((entry, index) => {
            const normalizeFieldName : string[] = entry.split("_x003a_");
            lookupFieldInfo.push(
              normalizeFieldName[normalizeFieldName.length-1]
            );
          });
          fieldInfo.DependentLookupInternalNames=lookupFieldInfo;             
        }        
        fieldInfo.IsUsedInForm=true;
        fieldInfo=this.getFieldMetaData(fieldInfo);
        fieldInfo.httpClient = this.props.httpClient;  
        fieldInfo.SiteUrl = this.props.siteURL;   
        if (fieldInfo.DefaultValue !== null && fieldInfo.DefaultValue.length>0)
        {
          fieldInfo.FormValue=fieldInfo.DefaultValue;
        }
        if (fieldInfo.FieldTypeKind === FieldTypes.BOOLEAN)
        {
          if (fieldInfo.FormValue==="1")
            fieldInfo.FormValue=true;
          else
            fieldInfo.FormValue=false
        }

        if (typeof this.props.addionalFieldRules !== "undefined" && this.props.addionalFieldRules !== null)
        {
          fieldInfo.AddionalRule=this.props.addionalFieldRules[fieldInfo.StaticName];
        }
        // new: ceck for specific default value - override in properties
        if (fieldInfo.AddionalRule !== undefined && fieldInfo.AddionalRule.DefaultValue.length>0)
        {
          fieldInfo.DefaultValue=fieldInfo.AddionalRule.DefaultValue;
          fieldInfo.FormValue=fieldInfo.DefaultValue;
        }        
      }
      return (
        <>{fieldInfo &&                     
          React.createElement(
            FormControlFluentUI,
            {
              ...fieldInfo,
              IsDisabled:this.state.isProcessing || this.state.isAlreadySent,              
              ChangedHandler: (field:ISPListField, value:string|string[]|ChoiceValue|boolean|LinkFieldValue, validationError:string) => {      
                const fieldInfo : ISPListField = this.availableFields.value.filter(f=>f.StaticName===fieldStaticName)[0];                 
                fieldInfo.FormValue = value;                                 
                fieldInfo.IsValid = validationError.length===0;
                this.ValidateCompleteForm();
                //console.log(this.availableFields.value.filter(f=>f.IsUsedInForm));
              },
              key: fieldInfo.StaticName              
            }
          )               
        }        
        </>
      );    
    }    
    return (<p>ERROR</p>);
  }

  private handleAttachment = (eventData:React.ChangeEvent<HTMLInputElement>) : void => { // async Promise<void>    
    // event processing
    const fileInfo : File = eventData.target.files[0];
    this.uploadFileList[eventData.target.id]=fileInfo;   
    this.ValidateCompleteForm();     
    //const result : ArrayBuffer|string = await this.getFileBuffer(eventData.target.files[0]);      
  }

  private ValidateCompleteForm():void
  {
    if (typeof this.props.allowedUploadFileTypes !== "undefined" && this.props.allowedUploadFileTypes.length>0)
    {
      let key: keyof {[key: string]: File};
      for (key in this.uploadFileList) {
        const fileInfo : File = this.uploadFileList[key];
        const parts : string[] = fileInfo.name.split(".");
        const extension = parts[parts.length-1];
        if (this.props.allowedUploadFileTypes.indexOf(extension)===-1)
        {
          // ERROR
          this.setState({isFormValid: false});    
          return;      
        }          
      }            
    }   
    this.setState({
      isFormValid: this.availableFields.value.filter(f=>f.IsUsedInForm && !f.IsValid).length===0
    });                
  }

  
  //https://medium.com/@ian.mundy/async-event-handlers-in-react-a1590ed24399
  private getFileBuffer(file:File) : Promise<ArrayBuffer|string> {    
    const reader = new FileReader();
    return new Promise(
      (resolve, reject) => {        
        reader.onload = function (e) {
          resolve(e.target.result);
        };
        reader.onerror = function (e) {
          reject(e.target.error);
        }
        reader.readAsArrayBuffer(file);
      }
    );            
  } 

  public saveFormData = () : void => {    
    this.setState({isProcessing: true});    
    type DynamicFormatData = {[key: string] : any}
    const fieldToSave : DynamicFormatData = {};
    this.availableFields.value.filter(f=>f.IsUsedInForm && typeof f.FormValue !== "undefined" && f.FormValue !== "").forEach((formEntry, index) => {

      if (formEntry.FieldTypeKind===FieldTypes.LOOKUP)
        fieldToSave[formEntry.InternalName+"Id"]=(formEntry.FormValue as ChoiceValue).Value;
      else
        fieldToSave[formEntry.InternalName]=formEntry.FormValue;

      // override specific
      /*if (formEntry.FieldTypeKind === FieldTypes.BOOLEAN) {
        fieldToSave[formEntry.InternalName]=formEntry.FormValue;
      }
      if (formEntry.FieldTypeKind === FieldTypes.CHOICE) {
        fieldToSave[formEntry.InternalName]=formEntry.FormValue;
      }
      if (formEntry.FieldTypeKind === FieldTypes.MULTICHOICE) {
        fieldToSave[formEntry.InternalName]=formEntry.FormValue;
      }*/
      if (formEntry.FieldTypeKind === FieldTypes.NUMBER) {
        if (formEntry.Decimals===0)
          fieldToSave[formEntry.InternalName] = parseInt(formEntry.FormValue.toString(),10);
        else
          fieldToSave[formEntry.InternalName] = parseFloat(formEntry.FormValue.toString());
      } 
      if (formEntry.FieldTypeKind === FieldTypes.URLORIMAGE) {
        fieldToSave[formEntry.InternalName] = formEntry.FormValue;
      }
      if (formEntry.FieldTypeKind === FieldTypes.DATETIME) {
        fieldToSave[formEntry.InternalName]=(formEntry.FormValue as Date).toISOString();
        /*const result : Date = Helper.parseDateTime(formEntry.FormValue.toString());
        if (result!==null)
        {
          fieldToSave[formEntry.InternalName]=result.toISOString();
        }*/
      }            
    });    
    // Datetime: http://blog.plataformatec.com.br/2014/11/how-to-serialize-date-and-datetime-without-losing-information/
    // https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj246742(v=office.15)
    this.props.httpClient.post(`${this.props.siteURL}/_api/web/lists/getbyid('${this.props.listID}')/items`, 
      SPHttpClient.configurations.v1, 
      { 
        headers: { 
          'Accept': 'application/json;odata=nometadata', 
          'Content-type': 'application/json;odata=nometadata', 
          'odata-version': ''            
        },         
        body: JSON.stringify(fieldToSave)
    })
    .then((x:SPHttpClientResponse) => {                            
      const test = x.json();      
      return test;      
    })
    .then(async (item: ISPListItem): Promise<void> => {
      this.sendConfirmationMail(item);
      this.setState({isProcessing: false, isAlreadySent: true, isFormValid: false});
      await this.uploadAttachments(item);
      alert(typeof this.props.successMessage !== "undefined" ? this.props.successMessage : "Vielen Dank. Die Daten wurden versendet.");
    });
  }

  public async sendConfirmationMail(item: ISPListItem):Promise<void>{    
    if (this.props.sendConfirmationEMail)
    {
      const listFormInfo = await this.props.httpClient.get(`${this.props.siteURL}/_api/web/lists/getbyid('${this.props.listID}')/Forms?$select=ServerRelativeUrl`, SPHttpClient.configurations.v1);
      const resultInfo : ISPFormLists = await listFormInfo.json();    
      const displayFormInfo = resultInfo.value.filter(x=>x.ServerRelativeUrl.indexOf('Dis')!==-1);
      let editLink = "";
      if (this.props.addDataLinkInEMail && displayFormInfo.length>0)
      {
        editLink = `<br /><br /><a href="${this.props.siteURL}/${displayFormInfo[0].ServerRelativeUrl}?ID=${item.Id}">${strings.MAILLinkTodata}</a><br />`;
      }
      const body : string = `<p><strong>${this.props.emailLeadText}</strong></p><table>` + this.availableFields.value.filter(f=>f.IsUsedInForm && typeof f.FormValue !== "undefined").map(entry => {
        return `<tr><td>${entry.Title}</td><td><strong>${Helper.GetFieldValueAsString(entry)}</strong></td></tr>`;
      }).join("")+"</table>"+editLink;                  
      Helper.sendEMail(this.props.currentUserEMail, this.props.emailSubject, body, this.props.siteURL, this.props.httpClient);
    }
  }

  public printFormData = () : void => {    
    const body : string = `<p><strong>${strings.HEADPrintForm}</strong></p><table>` + this.availableFields.value.filter(f=>f.IsUsedInForm && typeof f.FormValue !== "undefined").map(entry => {
      return `<tr><td>${entry.Title}</td><td><strong>${Helper.GetFieldValueAsString(entry)}</strong></td></tr>`;
    }).join("")+"</table>";     

    const wndPrint = window.open("about:blank","_blank");
    wndPrint.document.write(body);
    wndPrint.document.close();
    wndPrint.focus();
    wndPrint.print();
  }

  public resetForm = () : void => {    
    /*if (this.state.isProcessing)
      this.setState({
        "isProcessing": false
      });
    else
      this.setState({
        "isProcessing": true
      });*/        
    this.currentListID=null;
    this.setState({formFields: [], isProcessing: false, isFormValid: false, isAlreadySent: false});           
  }

  private async uploadAttachments(item: ISPListItem) : Promise<void>
  {
    for (const key in this.uploadFileList) {
      const fileObject : File = this.uploadFileList[key];
      const rawFileContent = await this.getFileBuffer(fileObject);
      await this.props.httpClient.post(`${this.props.siteURL}/_api/web/lists/getbyid('${this.props.listID}')/items(${item.Id})/AttachmentFiles/add(FileName='${fileObject.name}')`, 
      SPHttpClient.configurations.v1, 
      { 
        headers: { 
          'Accept': 'application/json', 
          'Content-type': 'application/json'
        },         
        body: rawFileContent
      });
    }
  }

  public componentDidMount () : void {
    this.attachmentCtl = [];
    for (let i = 0; i < this.props.uploads; i++) {
      this.attachmentCtl.push(
        <div>
          <label htmlFor={`FormAttachment${i}`}>{`${i+1}. ${strings.AttachmentIndexLabel}`}</label>
          <input type="file" onChange={this.handleAttachment} id={`FormAttachment${i}`} name={`FormAttachment${i}`} title={strings.AttachmentLabel} />
        </div>
      );
    }
  }

  private _onConfigure = () : void => {
    this.props.wpContext.propertyPane.open();
  }
  
  public render(): React.ReactElement<IDynamicFormularGeneratorProps> {    
    if (!this.validateConfiguration())
    {
      return (        
        <div className={styles.configWrapper}>
          <FlashSettings24Regular />
          <h2>{strings.CFGHeader}</h2>
          <ul>
            <li>{strings.CFGChooseList}</li>
            <li>{strings.CFGChooseView}</li>
          </ul>
          <Button onClick={this._onConfigure}>{strings.CFGBTNConfigure}</Button>
        </div>
      );
    }
    else
    {          
      this.qryFormFields();
      //ref={(el) => this.mainForm = el}
      return (        
        <form className={`${styles.dynamicFormularGenerator}`}>
          {this.state.isAlreadySent && <h3>Folgende Daten haben Sie erfolgreich versendet:</h3>}
          {this.props.description.length>0 && <p>{this.props.description}</p>}
          {this.state && this.state.formFields && this.state.formFields.map((val) => {             
              return this.formComponentFactory(val);
          })} 
          <div className={styles.uploadArea}>
            {this.props.uploads>0 &&  this.attachmentCtl && this.attachmentCtl.map((fileCtl) => {
                return fileCtl;
              })          
            }
          </div>  
          <div className={styles.cmdWrapper}>           
            {this.state.isProcessing ? <Spinner size="extra-small" label={strings.MSGWaiting}/>:<></>}            
            <Button id="btnSaveFormData" name="btnSaveFormData"
                    className={styles.btnSave}                                         
                    disabled={!this.state.isFormValid || this.state.isProcessing } 
                    onClick={this.saveFormData}>{strings.BTNSendFormData}</Button>
            {this.props.enablePrint && <Button id="btnPrintData" 
                    name="btnPrintData"                     
                    className={styles.btnPrint}
                    disabled={!this.state.isAlreadySent}
                    onClick={this.printFormData}>{strings.BTNPrintFormData}</Button>}
            <Button id="btnFormReset" name="btnFormReset" type="reset" onClick={this.resetForm}>{strings.BTNResetFormData}</Button>    
          </div>  
        </form>
      );    
    }
  }
}
