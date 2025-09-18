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
import { Dropdown, Option, Checkbox, Input, Label, Radio, RadioGroup, Textarea, Field, InfoLabel } from '@fluentui/react-components';
import * as React from 'react';
import { FieldTypes } from './FieldTypes';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './FormControlFluentUI.module.scss';
import { ValidationFactory } from './ValidationFactory';
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { ErrorCircle24Filled, NumberSymbolSquare24Regular, TextNumberFormat24Filled, CurrencyDollarEuro24Filled } from '@fluentui/react-icons';
var ValidationMarker = function (props) {
    if (props.errorMessage.length > 0) {
        return (React.createElement("div", { className: styles.errorPanel },
            React.createElement(ErrorCircle24Filled, null),
            React.createElement("span", null, props.errorMessage)));
    }
    return null;
};
var FormControlFluentUI = /** @class */ (function (_super) {
    __extends(FormControlFluentUI, _super);
    function FormControlFluentUI(fieldData) {
        var _this = _super.call(this, fieldData) || this;
        _this.UI_LABELSIZE = "small";
        _this._onDateTimeChanged = function (ev, data) {
            var formTag = ev.target;
            _this.manageFormFieldChanges(formTag.type === "checkbox" ? formTag.checked : formTag.value, formTag.id, formTag);
        };
        _this._onBlurHandlerDatePicker = function (ev) {
            var formTag = ev.target;
            _this.manageFormFieldChanges(new Date(formTag.value), formTag.id, formTag);
        };
        _this._onBlurHandler = function (ev) {
            var formTag = ev.target;
            _this.manageFormFieldChanges(formTag.type === "checkbox" ? formTag.checked : formTag.value, formTag.id, formTag);
        };
        _this._onBlurHandlerTextArea = function (ev) {
            var formTag = ev.target;
            _this.manageFormFieldChanges(formTag.value, formTag.id, formTag);
        };
        _this._onChangedRadioGroupList = function (ev, data) {
            _this.manageFormFieldChanges(data.value, _this.props.InternalName, null);
        };
        _this._onDropdDownSelectionChanged = function (event, data) {
            if (_this.props.FieldTypeKind === FieldTypes.MULTICHOICE) {
                _this.manageFormFieldChanges(data.selectedOptions, _this.props.InternalName, null);
            }
            else {
                _this.manageFormFieldChanges(data.optionValue, _this.props.InternalName, null);
            }
        };
        _this.state = {
            currentFormValue: "",
            errorMessage: ""
        };
        _this.onFormdataChanged = _this.props.ChangedHandler;
        return _this;
    }
    FormControlFluentUI.prototype.manageFormFieldChanges = function (newValue, formID, sourceElement) {
        var fieldValueToSet = newValue;
        var rawNewValue = fieldValueToSet.toString();
        if (this.props.FieldTypeKind === FieldTypes.URLORIMAGE) {
            if (formID.indexOf("Alternate") !== -1) {
                var linkValue = this.state.currentFormValue === "" ? "" : this.state.currentFormValue.Url;
                fieldValueToSet = {
                    Url: linkValue,
                    Description: newValue
                };
            }
            else {
                var labelValue = this.state.currentFormValue === "" ? "" : this.state.currentFormValue.Description;
                fieldValueToSet = {
                    Url: newValue,
                    Description: labelValue
                };
                rawNewValue = fieldValueToSet.Url;
            }
        }
        if (this.props.FieldTypeKind === FieldTypes.LOOKUP) {
            fieldValueToSet = this.state.lookupChoices.filter(function (x) { return x.Value === newValue; })[0];
            rawNewValue = fieldValueToSet.Title;
        }
        var validationResult = ValidationFactory.ValidateFormData(sourceElement, this.props, rawNewValue);
        this.setState({ errorMessage: validationResult, currentFormValue: fieldValueToSet });
        if (!this.onFormdataChanged)
            return;
        this.onFormdataChanged(this.props, fieldValueToSet, validationResult);
    };
    FormControlFluentUI.prototype.componentDidMount = function () {
        if (this.props.FieldTypeKind === FieldTypes.LOOKUP) {
            this.qryLookupListData();
        }
        /*window.addEventListener('dynamic-form-reset', () => {
          const ctl : HTMLFormElement = document.getElementById(this.props.InternalName) as HTMLFormElement;
        });*/
    };
    FormControlFluentUI.prototype.qryLookupListData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var filterFields, endpointFields, response, lookupFieldInfo, endpoint, responseData, data, choices, details;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        filterFields = this.props.DependentLookupInternalNames.map(function (fieldName) {
                            return "StaticName%20eq%20%27".concat(fieldName, "%27%20or%20Title%20eq%20%27").concat(fieldName, "%27");
                        }).join("%20or%20");
                        endpointFields = "".concat(this.props.SiteUrl, "/_api/web/lists/getbyid('").concat(this.props.LookupField.List, "')/Fields?$filter=").concat(filterFields);
                        return [4 /*yield*/, this.props.httpClient.get(endpointFields, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        lookupFieldInfo = _a.sent();
                        endpoint = "".concat(this.props.SiteUrl, "/_api/web/lists/getbyid('").concat(this.props.LookupField.List, "')/Items");
                        return [4 /*yield*/, this.props.httpClient.get(endpoint, SPHttpClient.configurations.v1)];
                    case 3:
                        responseData = _a.sent();
                        return [4 /*yield*/, responseData.json()];
                    case 4:
                        data = _a.sent();
                        choices = new Array();
                        details = "";
                        data.value.forEach(function (item) {
                            details = "";
                            if (_this.props.DependentLookupInternalNames !== null && _this.props.DependentLookupInternalNames.length > 0) {
                                details = _this.props.DependentLookupInternalNames.map(function (fieldName) {
                                    return _this.formatFieldValue(item[fieldName], lookupFieldInfo.value.filter(function (x) { return x.StaticName === fieldName || x.Title === fieldName; })[0]);
                                }).filter(function (x) { return x.length > 0; }).join(" | ");
                            }
                            choices.push({
                                Title: item[_this.props.LookupField.ShowField],
                                Value: item.ID,
                                Details: details
                            });
                        });
                        this.setState({
                            lookupChoices: choices
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    FormControlFluentUI.prototype.formatFieldValue = function (rawValue, field) {
        if (field !== null && rawValue !== null) {
            if (field.FieldTypeKind === FieldTypes.DATETIME) {
                var dateObj = new Date(rawValue);
                return dateObj.toLocaleDateString();
            }
        }
        return typeof rawValue === "undefined" || rawValue === null ? "" : rawValue;
    };
    FormControlFluentUI.prototype.render = function () {
        return (React.createElement("div", { className: styles.formFieldContainer },
            (this.props.FieldTypeKind !== FieldTypes.BOOLEAN || (this.props.Description.length > 0)) &&
                React.createElement("div", { className: this.props.FieldTypeKind === FieldTypes.BOOLEAN ? styles.inline : null },
                    this.props.FieldTypeKind !== FieldTypes.BOOLEAN &&
                        React.createElement(Label, { size: this.UI_LABELSIZE, htmlFor: this.props.InternalName, required: this.props.Required, id: "".concat(this.props.StaticName, "Label") }, this.props.Title),
                    this.props.FieldTypeKind !== FieldTypes.BOOLEAN && this.props.Description.length > 0 &&
                        React.createElement(InfoLabel, { id: "".concat(this.props.StaticName, "Info"), size: "medium", "aria-labelledby": "".concat(this.props.StaticName, "Label ").concat(this.props.StaticName, "Info"), info: this.props.Description }),
                    this.props.FieldTypeKind === FieldTypes.BOOLEAN && this.props.Description.length > 0 &&
                        React.createElement(InfoLabel, { id: "".concat(this.props.StaticName, "Info"), size: "large", info: this.props.Description })),
            this.renderFormControl(),
            React.createElement(ValidationMarker, { errorMessage: this.state.errorMessage, currentFormValue: this.state.currentFormValue })));
    };
    //onChange={this._onChangeHandler} 
    FormControlFluentUI.prototype.renderFormControl = function () {
        var _this = this;
        if (this.props.FieldTypeKind === FieldTypes.CHOICE || this.props.FieldTypeKind === FieldTypes.MULTICHOICE) {
            if (typeof this.props.Choices !== "undefined" && this.props.Choices.length > 0) {
                if (this.props.ChoiceUI === "RadioButtons") {
                    return (React.createElement("div", null,
                        React.createElement(RadioGroup, { "aria-labelledby": this.props.InternalName, name: this.props.InternalName, id: this.props.InternalName, onChange: this._onChangedRadioGroupList }, this.props.Choices.map(function (option) { return (React.createElement(Radio, { name: _this.props.InternalName, label: option, value: option, key: _this.props.InternalName, disabled: _this.props.ReadOnlyField || _this.props.IsDisabled })); }))));
                }
                return (React.createElement("div", null,
                    React.createElement(Dropdown, { "aria-labelledby": this.props.InternalName, name: this.props.InternalName, multiselect: this.props.FieldTypeKind === FieldTypes.MULTICHOICE, id: this.props.InternalName, inlinePopup: true, defaultValue: this.props.DefaultValue, disabled: this.props.ReadOnlyField || this.props.IsDisabled, onOptionSelect: this._onDropdDownSelectionChanged }, this.props.Choices.map(function (option) { return (React.createElement(Option, { key: option }, option)); }))));
            }
        }
        if (this.props.FieldTypeKind === FieldTypes.LOOKUP) {
            return (React.createElement("div", null,
                React.createElement(Dropdown, { disabled: this.props.ReadOnlyField || this.props.IsDisabled, name: this.props.InternalName, id: this.props.InternalName, multiselect: false, onOptionSelect: this._onDropdDownSelectionChanged }, this.state.lookupChoices && this.state.lookupChoices.map(function (option) { return (React.createElement(Option, { key: option.Value, value: option.Value, text: option.Title },
                    React.createElement("span", null,
                        option.Title,
                        option.Details && React.createElement("span", null,
                            React.createElement("br", null),
                            option.Details)))); }))));
        }
        if (this.props.FieldTypeKind === FieldTypes.BOOLEAN) {
            return (React.createElement("div", { className: this.props.Description.length > 0 ? styles.inline : null },
                React.createElement(Checkbox, { label: this.props.Title, disabled: this.props.ReadOnlyField || this.props.IsDisabled, name: this.props.InternalName, id: this.props.InternalName, defaultChecked: this.props.DefaultValue === "1", onChange: this._onBlurHandler })));
        }
        if (this.props.FieldTypeKind === FieldTypes.NOTE) {
            if (this.props.IsRichTextAllowed) {
                return (React.createElement("div", null,
                    React.createElement("p", null, "Not supported")));
            }
            else {
                return (React.createElement(Field, { size: 'small' },
                    React.createElement(Textarea, { disabled: this.props.ReadOnlyField || this.props.IsDisabled, name: this.props.InternalName, id: this.props.InternalName, resize: 'both', onBlur: this._onBlurHandlerTextArea })));
            }
        }
        if (this.props.FieldTypeKind === FieldTypes.URLORIMAGE) {
            if (this.props.LinkUI === "Hyperlink") {
                return (React.createElement(React.Fragment, null,
                    React.createElement("div", null,
                        React.createElement(Input, { name: this.props.InternalName, id: this.props.InternalName, placeholder: 'URL eingeben', type: "url", className: styles.textInput, disabled: this.props.ReadOnlyField || this.props.IsDisabled, onBlur: this._onBlurHandler })),
                    React.createElement("div", null,
                        React.createElement(Label, { size: this.UI_LABELSIZE, htmlFor: "".concat(this.props.InternalName, "Alternate"), required: false }, "Link Text"),
                        React.createElement(Input, { name: this.props.InternalName + 'Alternate', id: this.props.InternalName + 'Alternate', type: "text", className: styles.textInput, placeholder: 'alternativen Text', disabled: this.props.ReadOnlyField || this.props.IsDisabled, onBlur: this._onBlurHandler }))));
            }
            else {
                return (React.createElement("div", null,
                    React.createElement("p", null, "Not supported")));
            }
        }
        if (this.props.FieldTypeKind === FieldTypes.DATETIME) {
            return (React.createElement("div", null,
                React.createElement(DatePicker, { name: this.props.InternalName, id: this.props.InternalName, allowTextInput: true, onBlur: this._onBlurHandlerDatePicker, onChange: this._onDateTimeChanged, disabled: this.props.ReadOnlyField || this.props.IsDisabled })));
        }
        if (this.props.FieldTypeKind === FieldTypes.NUMBER || this.props.FieldTypeKind === FieldTypes.CURRENCY) {
            var placeHolder = "";
            if (Math.abs(this.props.MinimumValue) !== Number.MAX_VALUE)
                placeHolder = "Min. Wert ".concat(this.props.MinimumValue);
            if (this.props.MaximumValue !== Number.MAX_VALUE)
                placeHolder = "".concat(placeHolder.length > 0 ? placeHolder + " bis max. " : "Max. Wert: ", " ").concat(this.props.MaximumValue);
            return (React.createElement("div", null,
                React.createElement(Input, { name: this.props.InternalName, id: this.props.InternalName, onBlur: this._onBlurHandler, disabled: this.props.ReadOnlyField || this.props.IsDisabled, placeholder: placeHolder, contentBefore: this.props.FieldTypeKind === FieldTypes.NUMBER ? React.createElement(NumberSymbolSquare24Regular, null) : React.createElement(CurrencyDollarEuro24Filled, null), type: 'number' })));
        }
        return (React.createElement("div", null,
            React.createElement(Input, { name: this.props.InternalName, id: this.props.InternalName, defaultValue: this.props.DefaultValue, onBlur: this._onBlurHandler, disabled: this.props.ReadOnlyField || this.props.IsDisabled, className: styles.textInput, contentBefore: React.createElement(TextNumberFormat24Filled, null), type: this.props.FieldTypeKind === FieldTypes.NUMBER ? 'number' : 'text' })));
    };
    return FormControlFluentUI;
}(React.Component));
export { FormControlFluentUI };
//# sourceMappingURL=FormControlFluentUI.js.map