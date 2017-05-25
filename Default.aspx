<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="CIB" Namespace="Common.SharePoint.Web.UserControls" Assembly="Common.SharePoint.Web, Version=1.0.0.0, Culture=neutral, PublicKeyToken=68d03adf96d84a47" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="/_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.requestexecutor.js"></script>

    <!-- CIB common frameworks -->
    <CIB:CommonStyleSheet runat="server" Path="lib/bootstrap.min.css" />
    <CIB:CommonStyleSheet runat="server" Path="lib/bootstrap-theme.min.css" />
    <CIB:CommonStyleSheet runat="server" Path="Common/App.css" />
    <CIB:CommonScript runat="server" Path="lib/jquery-1.8.2.min.js" />
    <CIB:CommonScript runat="server" Path="lib/bootstrap.min.js" />
    <CIB:CommonScript runat="server" Path="Common/Utilities.js" />
    <CIB:CommonScript runat="server" Path="Common/Installer.js" />
    <CIB:CommonScript runat="server" Path="Common/Logger.js" />

	    <!-- Add your JavaScript to the following file -->
<%--     <script type="text/javascript" src="Install.js"></script> --%>

<script type="text/javascript" >


"use strict";
var FeeMan = FeeMan || {};
FeeMan.app = FeeMan.app || {};
FeeMan.app.appdisplay = function () {

    var globalContext;
    var context;
    var hostContext;
    var hostWeb;

    /////////////////////////////////////////
    // Config
    /////////////////////////////////////////
    ////// ************ //////////////////////////
    var testAddId = 'D';
    var testAddName = 'D';
    ////// ************ //////////////////////////
    var groupName = 'CIB DE FeeMan';
    var siteColumns = {
        // Currency
        currencyCode: {
            name: 'fmCode' + testAddName,
            displayName: 'Code' + testAddName,
            id: '{3d9abb16-d9f6-4381-8e30-db1812e6c6e' + testAddId + '}',
            type: 'Text',
            maxLength: '3',
            required: true,
            indexed: true,
            enforceUnique: true,
            linkToItem: 'Required',
            group: groupName
        },
        // Team
        teamName: {
            name: 'fmName' + testAddName,
            displayName: 'Name' + testAddName,
            id: '{62cb9781-a886-4641-ac7d-cc3d7b13239' + testAddId + '}',
            type: 'Text',
            maxLength: '20',
            required: true,
            indexed: true,
            enforceUnique: true,
            linkToItem: 'Required',
            group: groupName
        },
        teamMembers: {
            name: 'fmMembers' + testAddName,
            displayName: 'Members' + testAddName,
            id: '{8618d2f1-e739-4bc0-a2c4-1dfb1557880' + testAddId + '}',
            type: 'UserMulti',
            multi: true,
            userSelectionMode: 'PeopleOnly',
            group: groupName
        },
        // PriceList
        plvPlvDescription: {
            name: 'fmPLVDescription' + testAddName,
            displayName: 'PLV Description' + testAddName,
            id: '{c746fea0-5e55-4772-a9c4-96b180ff352' + testAddId + '}',
            type: 'Text',
            maxLength: '255',
            required: true,
            indexed: true,
            enforceUnique: true,
            linkToItem: 'Required',
            group: groupName
        },
        plvCode: {
            name: 'fmPLVCode' + testAddName,
            displayName: 'Code' + testAddName,
            id: '{3797cebb-f263-47ab-8ec5-32277e654c0' + testAddId + '}',
            type: 'Text',
            maxLength: '10',
            required: true,
            indexed: true,
            group: groupName
        },
        plvStandardPrice: {
            name: 'fmStandardPrice' + testAddName,
            displayName: 'Standard Price' + testAddName,
            id: '{ac671946-6460-4aa7-abce-23ca5f4a7de' + testAddId + '}',
            type: 'Number',
            decimals: '2',
            lcid: '1031',
            required: true,
            group: groupName
        },
        plvDescription: {
            name: 'fmDescription' + testAddName,
            displayName: 'Description' + testAddName,
            id: '{65afe706-62af-40eb-bf7c-4c8db89ed2b' + testAddId + '}',
            type: 'Text',
            maxLength: '255',
            group: groupName
        },
        plvTeams: {
            name: 'fmTeams' + testAddName,
            displayName: 'Teams' + testAddName,
            id: '{0f990193-2ffe-4612-901f-0d0a7558e42' + testAddId + '}',
            type: 'LookupMulti',
            lookupList: 'Teams' + testAddName,
            lookupField: 'fmName' + testAddName,
            // EnforceUniqueValues="FALSE" RelationshipDeleteBehavior="None"
            group: groupName
        },
        feeTeam: {
            name: 'fmTeam' + testAddName,
            displayName: 'Team' + testAddName,
            id: '{71a21909-9ac6-42c9-8d55-9935ba2dd18' + testAddId + '}',
            type: 'Lookup',
            lookupList: 'Teams' + testAddName,
            lookupField: 'fmName' + testAddName,
            required: true,
            // EnforceUniqueValues="FALSE" RelationshipDeleteBehavior="None"
            group: groupName
        },
        feeUnits: {
            name: 'fmUnits' + testAddName,
            displayName: 'Units' + testAddName,
            id: '{957d4ece-bc38-4f75-a08d-6b5f962b42e' + testAddId + '}',
            type: 'Number',
            decimals: '0',
            required: true,
            // EnforceUniqueValues="FALSE" RelationshipDeleteBehavior="None"
            group: groupName
        },
        feePlvDescription: {
            name: 'fmFeePLVDescription' + testAddName,
            displayName: 'PLV Description' + testAddName,
            id: '{67c0479c-2ca9-4ecf-8f01-62fb53aa617' + testAddId + '}',
            type: 'Lookup',
            lookupList: 'PriceList' + testAddName,
            lookupField: 'fmPLVDescription' + testAddName,
            required: true,
            // EnforceUniqueValues="FALSE" RelationshipDeleteBehavior="None"
            additionalFields: [ { target: 'fmPLVCode' + testAddName, displayName: 'Operation Code' + testAddName } ],
            group: groupName
        },
        feeAccount: {
            name: 'fmAccount' + testAddName,
            displayName: 'Customer Account Number' + testAddName,
            id: '{6f4338d7-d741-4058-8d09-c2dfeec3049' + testAddId + '}',
            type: 'Text',
            maxLength: '30',
            required: true,
            linkToItem: 'Required',
            //        <Validation Message="Account number must be like:&#xD;&#xA;5 digits BLANK 6 digits BLANK 3 digits BLANK 2 digits BLANK then three characters for the currency"
            //           Script="function(x){return SP.Exp.Calc.valid(SP.Exp.Node.f(&#39;EQ&#39;,[SP.Exp.Node.f(&#39;LEN&#39;,[SP.Exp.Node.a(0)]),SP.Exp.Node.v(23)]),x)}">=(LEN(Account)=23)
            //        </Validation>
            group: groupName
        },
        feeClientName: {
            name: 'fmClientName' + testAddName,
            displayName: 'Client Name' + testAddName,
            id: '{a3e4f8a4-5ae7-4099-8f9e-7cc8201922f' + testAddId + '}',
            type: 'Text',
            maxLength: '255',
            required: true,
            group: groupName
        },
        feeTransactionAmount: {
            name: 'fmTransactionAmount' + testAddName,
            displayName: 'Transaction Amount' + testAddName,
            id: '{f6c96dac-dc06-4fce-9198-96faed821b5' + testAddId + '}',
            type: 'Number',
            decimals: '2',
            group: groupName
        },
        feeCurrency: {
            name: 'fmCurrency' + testAddName,
            displayName: 'Currency' + testAddName,
            id: '{883328a7-e9c2-41ff-ab74-ab749fc7a0f' + testAddId + '}',
            type: 'Lookup',
            lookupList: 'Currencies' + testAddName,
            lookupField: 'fmCode' + testAddName,
            // EnforceUniqueValues="FALSE" RelationshipDeleteBehavior="None"
            group: groupName
        },
        feeAmount: {
            name: 'fmAmount' + testAddName,
            displayName: 'Amount' + testAddName,
            id: '{825e77bc-68d4-4a9f-a0ef-e051b0e4ce2' + testAddId + '}',
            type: 'Number',
            decimals: '2',
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        feeComments: {
            name: 'fmComments' + testAddName,
            displayName: 'Comments' + testAddName,
            id: '{05be0f26-1830-42c0-a427-492766be736' + testAddId + '}',
            type: 'Note',
            numLines: 10,
            richText: false,
            group: groupName
        },
        feeStatus: {
            name: 'fmStatus' + testAddName,
            displayName: 'Status' + testAddName,
            id: '{25e3c5c9-46c1-4c00-9e1c-f86c2b7999a' + testAddId + '}',
            type: 'Choice',
            format: SP.ChoiceFormatType.dropdown,
            fillInChoice: false,
            showInEditForm: false,
            showInNewForm: false,
            choices: ["Draft/Invalid", "Requested", "Approved", "InputInvalid"],
            group: groupName
        },
        feeStatus_: {
            name: 'fmStatus_' + testAddName,
            displayName: 'Status' + testAddName,
            id: '{21a72381-785a-464e-9647-971aeb78b59' + testAddId + '}',
            type: 'Calculated',
            formula: '=fmStatus' + testAddName,
            resultType: 'Text',
            showInEditForm: false,
            showInNewForm: false,
            readOnly: true,
            fieldRefs: ['fmStatus' + testAddName],
            group: groupName
        },
        feeRequest: {
            name: 'fmRequest' + testAddName,
            displayName: 'Request' + testAddName,
            id: '{9a559b34-d1ce-4739-820f-c6a7085fd43' + testAddId + '}',
            type: 'Choice',
            format: SP.ChoiceFormatType.dropdown,
            fillInChoice: false,
            showInEditForm: false,
            showInNewForm: false,
            choices: ["Draft/Invalid", "Request", "Cancel"],
            defaultValue: "Request",
            required: true,
            group: groupName
        },
        feeRequestedBy: {
            name: 'fmRequestedBy' + testAddName,
            displayName: 'RequestedBy' + testAddName,
            id: '{6d941747-a2ee-483d-95ce-8f7cfc1cef9' + testAddId + '}',
            type: 'User',
            userSelectionMode: 'PeopleOnly',
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        feeApprove: {
            name: 'fmApprove' + testAddName,
            displayName: 'Approve' + testAddName,
            id: '{a300fd56-bb6a-479e-a117-ceb1469ecb9' + testAddId + '}',
            type: 'Choice',
            format: SP.ChoiceFormatType.dropdown,
            fillInChoice: false,
            showInEditForm: false,
            showInNewForm: false,
            choices: ["-", "Approve", "Invalid", "Reject"],
            defaultValue: "-",
            required: true,
            group: groupName
        },
        feeApprovedAt: {
            name: 'fmApprovedAt' + testAddName,
            displayName: 'ApprovedAt' + testAddName,
            id: '{32c0d2ea-b2cc-4a98-9950-f16e7c98359' + testAddId + '}',
            type: 'DateTime',
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        feeApprovedBy: {
            name: 'fmApprovedBy' + testAddName,
            displayName: 'ApprovedBy' + testAddName,
            id: '{e90a6e93-a56d-4d67-aeeb-33b7da1a35f' + testAddId + '}',
            type: 'User',
            userSelectionMode: 'PeopleOnly',
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        feeValidate: {
            name: 'fmValidate' + testAddName,
            displayName: 'Validate' + testAddName,
            id: '{0544c533-7652-43dc-abd2-0eb38072aa5' + testAddId + '}',
            type: 'Choice',
            format: SP.ChoiceFormatType.dropdown,
            fillInChoice: false,
            showInEditForm: false,
            showInNewForm: false,
            choices: ["-", "Validate", "InputInvalid"],
            defaultValue: "-",
            required: true,
            group: groupName
        },
        feeStatusHistory: {
            name: 'fmStatusHistory' + testAddName,
            displayName: 'Status History' + testAddName,
            id: '{baa9c5de-9291-436f-a0fd-cd7a23003a5' + testAddId + '}',
            type: 'Note',
            numLines: 10,
            richText: false,
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        feeWorkflowActive: {
            name: 'fmWorkflowActive' + testAddName,
            displayName: 'Workflow Active' + testAddName,
            id: '{4f7c0748-9b64-411f-ac47-88605aac648' + testAddId + '}',
            type: 'Choice',
            format: SP.ChoiceFormatType.dropdown,
            fillInChoice: false,
            showInEditForm: false,
            showInNewForm: false,
            choices: ["-", "Active", "Error"],
            defaultValue: "-",
            required: true,
            group: groupName
        },
    };

    var siteCT = {
        config: {
            name: 'CIB Feem Config' + testAddName,
            id: '0x01009FC47BE0766A4455B10560438A54CBD' + testAddId,
            group: groupName,
            columns: [siteColumns.teamMembers],
            columnNames: [siteColumns.teamMembers.name]
        },
        currency: {
            name: 'CIB Feem Currencies' + testAddName,
            id: '0x0100F34E70C088364B1B85B9462BA830A28' + testAddId,
            group: groupName,
            columns: [siteColumns.currencyCode],
            columnNames: [siteColumns.currencyCode.name]
        },
        team: {
            name: 'CIB Feem Teams' + testAddName,
            id: '0x01007DDEF9578DB3451495FF451DF3B9539' + testAddId,
            group: groupName,
            columns: [siteColumns.teamName, siteColumns.teamMembers],
            columnNames: [siteColumns.teamName.name, siteColumns.teamMembers.name],
            newColumns: [siteColumns.teamName]
        },
        priceList: {
            name: 'CIB Feem Price List' + testAddName,
            id: '0x0100DB215693108F478FA58825DE78BEB51' + testAddId,
            group: groupName,
            columns: [siteColumns.plvPlvDescription, siteColumns.plvCode, siteColumns.plvStandardPrice, siteColumns.plvDescription, siteColumns.plvTeams],
            columnNames: [siteColumns.plvPlvDescription.name, siteColumns.plvCode.name, siteColumns.plvStandardPrice.name, siteColumns.plvDescription.name, siteColumns.plvTeams.name]
        },
        fee: {
            name: 'CIB Feem Fees' + testAddName,
            id: '0x010091DD1E064DBC4B9AA0D9B0E1000FDB1' + testAddId,
            group: groupName,
/*            columns3: [siteColumns.feeTeam, siteColumns.feeStatus, siteColumns.feeStatus_, siteColumns.feeRequest, siteColumns.feeApprove, siteColumns.feeValidate],
            columnNames3: [siteColumns.feeTeam.name, siteColumns.feeStatus.name, siteColumns.feeStatus_.name, siteColumns.feeRequest.name, siteColumns.feeApprove.name, siteColumns.feeValidate.name],
*/
            columns: [siteColumns.feeTeam, siteColumns.feeUnits, siteColumns.feePlvDescription, siteColumns.feeAccount, siteColumns.feeClientName,
                siteColumns.feeTransactionAmount, siteColumns.feeCurrency, siteColumns.feeAmount, siteColumns.feeComments, siteColumns.feeStatus, siteColumns.feeStatus_,
                siteColumns.feeRequest, siteColumns.feeRequestedBy, siteColumns.feeApprove, siteColumns.feeApprovedAt, siteColumns.feeApprovedBy, siteColumns.feeValidate,
                siteColumns.feeStatusHistory, siteColumns.feeWorkflowActive],
            columnNames: [siteColumns.feeTeam.name, siteColumns.feeUnits.name, siteColumns.feePlvDescription.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.feeStatus.name, siteColumns.feeStatus_.name,
                siteColumns.feeRequest.name, siteColumns.feeRequestedBy.name, siteColumns.feeApprove.name, siteColumns.feeApprovedAt.name, siteColumns.feeApprovedBy.name, siteColumns.feeValidate.name,
                siteColumns.feeStatusHistory.name, siteColumns.feeWorkflowActive.name]
        }                           
    };

    var siteLists = {
        config: {
            name: 'Config' + testAddName,
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.config.id],
            views: {
                default: {
                    name: 'default' + testAddName, columns: siteCT.config.columnNames
                }
            }
        },
        currency: {
            name: 'Currencies' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.currency.id],
            views: {
                default: {
                    name: 'default' + testAddName, columns: siteCT.currency.columnNames, query: '<OrderBy><FieldRef Name="' + siteColumns.currencyCode.name + '" Ascending="TRUE"></FieldRef></OrderBy>'
                }
            }
        },
        team: {
            name: 'Teams' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.team.id],
            views: {
                default: {
                    name: 'default' + testAddName, columns: siteCT.team.columnNames, query: '<OrderBy><FieldRef Name="' + siteColumns.teamName.name + '" Ascending="TRUE"></FieldRef></OrderBy>'
                }
            }
        },
        priceList: {
            name: 'PriceList' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.priceList.id],
            views: {
                default: {
                    name: 'default' + testAddName, columns: siteCT.priceList.columnNames, query: '<OrderBy><FieldRef Name="' + siteColumns.plvPlvDescription.name + '" Ascending="TRUE"></FieldRef></OrderBy>'
                }
            }
        },
        fee: {
            name: 'Fees' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.fee.id],
            views: {
                default: {
                    name: 'default' + testAddName,
                    columns: siteCT.fee.columnNames,
/*                    columns: [siteColumns.feeTeam.name, siteColumns.feeUnits.name, siteColumns.feePlvDescription.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.feeStatus.name, siteColumns.feeStatus_.name,
                        siteColumns.feeRequest.name, siteColumns.feeRequestedBy.name, siteColumns.feeApprove.name, siteColumns.feeApprovedAt.name, siteColumns.feeApprovedBy.name, siteColumns.feeValidate.name,
                        siteColumns.feeStatusHistory.name, siteColumns.feeWorkflowActive.name, 'ID'],*/
                    query: '<OrderBy><FieldRef Name="ID" Ascending="TRUE"></FieldRef></OrderBy>'
                }
            }
        }
    };
    var TITLE = 'Title';
    var ITEM = 'Item';
    var ALL_ITEMS = 'All Items';

    /////////////////////////////////////////
    // Retrive subsites to install app
    /////////////////////////////////////////
    var retrieveSubsites = function () {
        var hostWeb = hostContext.get_web();
        var hostSubSites = hostWeb.get_webs();
        context.load(hostWeb);
        context.load(hostSubSites);
        return context.executeQueryAsyncPromise().done(function () {
            var subs = hostSubSites.getEnumerator();
            var subSite;
            var subTitles = hostWeb.get_title();
            var subTitlesUrl = $.getHostWebUrl();
//             var subTitlesUrl = $.getServerRealtiveHostWebUrl();
            $('#drp_subsite').append($('<option>', { value: subTitlesUrl }).text(subTitles + ' (' + subTitlesUrl + ')'));
            while (subs.moveNext()) {
                subSite = subs.get_current();
                var webTitle = subSite.get_title();
                if (webTitle && subSite.get_webTemplate() !== 'APP') {
                    subTitles = subSite.get_title();
//                     subTitlesUrl = utils.getServerRelativeUrl(subSite.get_url());
                    subTitlesUrl = subSite.get_url();
                    $('#drp_subsite').append($('<option>', { value: subTitlesUrl }).text(subTitles + ' (' + subTitlesUrl + ')'));
                } 
            }
         })
         .fail(function (message) {
             CIB.installer.message('Error in retrieving subsites: ' + message, 'error');
         });
    };

    var installer = function () {
        var listIds = {};
        // This multidimentional array will be storing list name - list view - view ID data
        var listToViewIds = {};

        var helper = function () {
            return {
                getBoolean: function (s) {
                    var ret = Boolean(s);
                    if (ret) {
                        if (s.toString().toLowerCase() == 'false') { return false; }
                    }
                    return ret;
                },

                getServerRelativeUrl: function (n) {
                    return n.replace("http://", "").replace("https://", "").indexOf("/") < 0 && (n += "/"), "/" + n.replace(/^(?:\/\/|[^\/]+)*\//, "");
                },

                /*
                    Executes a query containing a series of scopes for SharePoint. 
                    Will resolve or reject a promise based on the results
                */
                executeQuery: function (scopes, promise, value) {
                    context.executeQueryAsync(
                        function () {
                            var messages = [];
                            var handled = true;
                            $.each(scopes, function () {
                                var scope = this;
                                if (scope.get_hasException()) {
                                    var error = helper.handleError(this, scope);
                                    handled &= error.handled;
                                    messages.push(error.message);
                                }
                                else {
                                    helper.message(scope.successMessage, 'success');
                                    messages.push(scope.successMessage);
                                }
                            });
                            if (handled) { if (value) { promise.resolve(messages, value); } else { promise.resolve(messages); } }
                            else { promise.reject(messages); }

                        }, function (sender, args) {
                            var error = helper.handleError(sender, args);
                            if (error.handled) { promise.resolve(error.message); }
                            else { promise.reject(error.message); }
                        });
                },

                /*
                    Due to SharePoint's CSOM, in cases the most efficient way of detecting if something is already is provisioned
                    is to create it and handle the exception. The following determines if exceptions are expected.
                */
                handleError: function (sender, args) {
                    var message = args.get_message ? args.get_message() : args.get_errorMessage();
                    var expectedErrorMessages = [
                        ' is already activated at scope ',
                        'A duplicate field name "',
                        'A duplicate content type "',
                        'A file or folder with the name ',
                        'The specified name is already in use.',
                        'A list, survey, discussion board, or document library with the specified title already exists in this Web site.'];
                    var type = 'error';
                    for (var i = 0; i < expectedErrorMessages.length; i++) {
                        if (message.slice(0, expectedErrorMessages[i].length) == expectedErrorMessages[i] || message.indexOf(expectedErrorMessages[i]) > -1) {
                            type = 'info';
                            message += ' (expected if provisioned already)';
                            break;
                        }
                    }

                    helper.message(message, type);

                    return { handled: type == 'info', message: message };
                },

                /*
                    Writes colour coded messages to the user, this method assumes the existence of an element with id 'install-status'
                */
                message: function (text, type) {
                    if (!type) type = 'message';
                    if (console && console.log) console.log(text + ' [' + type + ']');
                    if (type == 'error') CIB.logging.logError('Provisioning', text, window.location.href);
                    var colour = type == 'success' ? 'green' : (type == 'error' ? 'red' : (type == 'info' ? 'orange' : 'gray'));
                    $('#install-status').append('<span style="color:' + colour + '">' + text + '</span>');
                    var elem = document.getElementById('install-status');
                    if (elem) elem.scrollTop = elem.scrollHeight;
                },

                /*
                    Get all list ids to for use by lookup columns
                */
                updateListIds: function () {
                    var listIdsUpdated = new jQuery.Deferred();

                    // Popuate list
                    //if ($.isEmptyObject(listIds))
                    //{
                    var lists = hostContext.get_web().get_lists();
                    context.load(lists, 'Include(Title, Id)');
                    context.executeQueryAsync(function () {
                        var listEnumerator = lists.getEnumerator();
                        while (listEnumerator.moveNext()) {
                            var list = listEnumerator.get_current();
                            listIds[list.get_title()] = list.get_id();
                        }
                        listIdsUpdated.resolve();
                    }, function (sender, args) {
                        var error = helper.handleError(sender, args);
                        listIdsUpdated.reject(error.message);
                    });
                    /*}
                    else
                    {
                        listIdsUpdated.resolve();
                    }*/

                    return listIdsUpdated.promise();
                },

                /*
                    Get views by list name
                */
                getViewsForList: function (promises, listName) {
                    var dfd = new $.Deferred();

                    var web = hostContext.get_web();
                    var list = web.get_lists().getByTitle(listName);
                    var views = list.get_views();

                    context.load(views, 'Include(Id, Title)');

                    context.executeQueryAsync(function () {
                        var viewEnumerator = views.getEnumerator();

                        while (viewEnumerator.moveNext()) {
                            var existingView = viewEnumerator.get_current();
                            var viewName = existingView.get_title();
                            var viewId = existingView.get_id().toString().toLowerCase();

                            if ($.isEmptyObject(listToViewIds[listName]))
                            {
                                listToViewIds[listName] = [];
                            }

                            listToViewIds[listName][viewName] = viewId;
                        }

                        dfd.resolve();
                    },
                    function (sender, args) {
                        helper.handleError(sender, args);
                        dfd.fail();
                    }
                    );

                    promises.push(dfd);
                },

                /*
                    Get all view ids for each  list and store them in multidimentional array
                */
                updateViewIds: function () {
                    var promises = [];

                    // Populate object with data only if it's empty
                    /*if ($.isEmptyObject(listToViewIds))
                    {*/
                    for(var listName in listIds)
                    {
                        helper.getViewsForList(promises, listName);
                    }
                    /*}
                    else
                    {
                        var dfd = new $.Deferred();
                        promises.push(dfd);
                        dfd.resolve();
                    }*/

                    // Wait for all async operations to complete before moving on to the next step
                    return $.when.apply($, promises).promise();
                }
            };
        }();

        return {
            updateListIds: function () {
                return helper.updateListIds()
            },

            createSiteColumns: function (columns) {
                var fields = hostContext.get_web().get_fields();
                return installer.createColumns(columns, fields);
            },

            /*
                Create a site column in the host web
                @columns { name: 'cmppYear', id: '{F4605722-C180-46B0-8AAE-0C0BC0EA4EC3}', displayName: 'Year', type: 'Number', group: 'Test' }
                @fields a field collection from a web or list object
                Addtional parameters are supported for lookups, calculated, datetime and choice fields
            */
            createColumns: function (columns, fields) {
                var scopes = [];
                var columns = CIB.utilities.ensureArray(columns);
                var columnsCreated = new jQuery.Deferred();

                if (!fields)
                    throw new Error('Field collection not provided, use createSiteColumns or createListColumns instead.');

                var createColumns = function () {
                    $.each(columns, function () {
                        var column = this;

                        if (!column.id || !column.name || !column.type || !column.displayName || !column.group)
                            throw new Error('Column object must have id, name, type, group and displayName attributes');

                        var scope = $.handleExceptionsScope(context, function () {
                            CIB.installer.message('Creating column \'' + column.displayName + '\'');

                            var multi = (helper.getBoolean(column.multi) || (column.type.toLowerCase() == "lookupmulti") || (column.type.toLowerCase() == "usermulti"));
                            var indexed = (helper.getBoolean(column.indexed) || helper.getBoolean(column.enforceUnique));

                            var fieldXml = "<Field ID='" + column.id + "' Type='" + column.type + "' DisplayName='" + column.name +
                                "' Name='" + column.name + "' Group='" + column.group + "' Required='" + helper.getBoolean(column.required).toString().toUpperCase() + "' />";

                            if ((column.type.toLowerCase() == "user")  || (column.type.toLowerCase() == "usermulti")) {
                                fieldXml = fieldXml.replace(" />", " List='UserInfo' ShowField='ImnName' " +
                                    (column.hasOwnProperty('userSelectionMode') ? ("UserSelectionMode='" + column.userSelectionMode + "'") : '') +
                                    " UserSelectionScope='0' />"); // TODO
                            }
                            if (column.hasOwnProperty('maxLength')) { fieldXml = fieldXml.replace(" />", " MaxLength='" + column.maxLength + "' />"); }
                            if (column.hasOwnProperty('numLines')) { fieldXml = fieldXml.replace(" />", " NumLines='" + column.numLines + "' />"); }
                            if (column.hasOwnProperty('richText')) {
                                fieldXml = fieldXml.replace(" />", " RichText='" + helper.getBoolean(column.richText).toString().toUpperCase() + "' />");
                            }
                            if (column.hasOwnProperty('enforceUnique')) {
                                fieldXml = fieldXml.replace(" />", " AllowDuplicateValues='" + (!helper.getBoolean(column.enforceUnique)).toString().toUpperCase() + "' EnforceUniqueValues='" + helper.getBoolean(column.enforceUnique).toString().toUpperCase() + "' />");
                            }
                            if (indexed) { fieldXml = fieldXml.replace(" />", " Indexed='TRUE' />"); }
                            if (multi) { fieldXml = fieldXml.replace(" />", " Mult='TRUE' />") };

                            if (column.hasOwnProperty('readOnly')) {
                                fieldXml = fieldXml.replace(" />", " ReadOnly='" + helper.getBoolean(column.readOnly).toString().toUpperCase() + "' />");
                            }

                            if (column.hasOwnProperty('showInDisplayForm')) {
                                fieldXml = fieldXml.replace(" />", " ShowInEditForm='" + helper.getBoolean(column.showInDisplayForm).toString().toUpperCase() + "' />");
                            }
                            if (column.hasOwnProperty('showInNewForm')) {
                                fieldXml = fieldXml.replace(" />", " ShowInNewForm='" + helper.getBoolean(column.showInNewForm).toString().toUpperCase() + "' />");
                            }
                            if (column.hasOwnProperty('showInEditForm')) {
                                fieldXml = fieldXml.replace(" />", " ShowInEditForm='" + helper.getBoolean(column.showInEditForm).toString().toUpperCase() + "' />");
                            }

                            if (column.linkToItem) {
                                fieldXml = fieldXml.replace(" />", " LinkToItem='TRUE' LinkToItemAllowed='" + column.linkToItem + "' ListItemMenu='TRUE' />");
                            }
                            if ((column.type.toLowerCase() == 'number') && column.hasOwnProperty('decimals')) {
                                fieldXml = fieldXml.replace(" />", " Decimals='" + column.decimals + "' />");
                            }
                            if ((column.type.toLowerCase() == 'choice') && column.hasOwnProperty('format')) {
                                var format = '';
                                if (column.format == SP.ChoiceFormatType.radioButtons) { format = 'RadioButtons' }
                                else if (column.format == SP.ChoiceFormatType.registerEnum) { format = 'RegisterEnum' }
                                else { format = 'Dropdown' }
                                fieldXml = fieldXml.replace(" />", " Format='" + format + "' />");
                            }
                            if ((column.type.toLowerCase() == 'choice') && column.hasOwnProperty('fillInChoice')) {
                                fieldXml = fieldXml.replace(" />", " FillInChoice='" + helper.getBoolean(column.fillInChoice).toString().toUpperCase() + "' />");
                            }

                            if (column.hasOwnProperty('lcid')) {
                                fieldXml = fieldXml.replace(" />", " LCID='" + column.lcid + "' />");
                            }
                            if (column.type.toLowerCase() == 'calculated') {
                                if (!column.formula || !column.resultType)
                                    throw new Error('Calculated columns must have a formula and resultType set');
                                var formulaXml = '<Formula>' + column.formula + '</Formula>';
                                if (column.hasOwnProperty('fieldRefs')) {
                                    formulaXml += '<FieldRefs>'
                                    var fieldRefs = CIB.utilities.ensureArray(column.fieldRefs);
                                    $.each(fieldRefs, function () {
                                        var fieldRef = this;
                                        formulaXml += "<FieldRef Name='" + fieldRef + "'/>"
                                    });
                                    formulaXml += '</FieldRefs>'
                                }
                                fieldXml = fieldXml.replace(' />', ' ResultType="' + column.resultType + '">' + formulaXml + '</Field>');
                            }

                            var field = fields.addFieldAsXml(fieldXml, false, SP.AddFieldOptions.AddToNoContentType);

                            if (column.hasOwnProperty('hidden')) {
                                field.set_hidden(helper.getBoolean(column.hidden));
                            }

                            field.set_title(column.displayName);
                            field.set_required(helper.getBoolean(column.required));
/*
                            if (column.hasOwnProperty('showInDisplayForm')) {
                                field.setShowInDisplayForm(helper.getBoolean(column.showInDisplayForm));
                            }
                            if (column.hasOwnProperty('showInNewForm')) {
                                field.setShowInNewForm(helper.getBoolean(column.showInNewForm));
                            }
                            if (column.hasOwnProperty('showInEditForm')) {
                                field.setShowInEditForm(helper.getBoolean(column.showInEditForm));
                            }
*/
                            if (column.defaultValue)
                                field.set_defaultValue(column.defaultValue);

                            context.load(field);

                            if (column.type.toLowerCase() == 'lookup') {
                                if (!listIds[column.lookupList]) {
                                    var message = 'The id for the list ' + column.lookupList + ' has not been loaded. updateListIds must be called before creating lookup fields';
                                    columnsCreated.reject(message);
                                    throw new Error(message);
                                }
                                var fieldLookup = context.castTo(field, SP.FieldLookup);
                                fieldLookup.set_lookupList(listIds[column.lookupList]);
                                fieldLookup.set_lookupField(column.lookupField);
                                fieldLookup.update();

                                if (column.additionalFields) {
                                    $.each(column.additionalFields, function () {
                                        var additionalColumn = this;
                                        fields.addDependentLookup(additionalColumn.displayName, field, additionalColumn.target);
                                    });
                                }
                            }
                            // below code is to handle MultiLookup fields
                            else if (column.type.toLowerCase() == 'lookupmulti') {
                                if (!listIds[column.lookupList]) {
                                    var message = 'The id for the list ' + column.lookupList + ' has not been loaded. updateListIds must be called before creating lookup fields';
                                    columnsCreated.reject(message);
                                    throw new Error(message);
                                }
                                var fieldLookup = context.castTo(field, SP.FieldLookup);
                                fieldLookup.set_lookupList(listIds[column.lookupList]);
                                fieldLookup.set_lookupField(column.lookupField);
                                fieldLookup.set_allowMultipleValues(true);
                                fieldLookup.update();

                                if (column.additionalFields) {
                                    $.each(column.additionalFields, function () {
                                        var additionalColumn = this;
                                        fields.addDependentLookup(additionalColumn.displayName, field, additionalColumn.target);
                                    });
                                }
                            }
                            else if (column.type.toLowerCase() == 'currency' && column.locale) {
                                var fieldCurrency = context.castTo(field, SP.FieldCurrency);
                                fieldCurrency.set_currencyLocaleId(column.locale);
                                fieldCurrency.update();
                            }
                            else if (column.type.toLowerCase() == 'number') {
                                var fieldNumber = context.castTo(field, SP.FieldNumber);
                                if (column.minimumValue)
                                    fieldNumber.set_minimumValue(column.minimumValue);
                                if (column.maximumValue)
                                    fieldNumber.set_maximumValue(column.maximumValue);
                                fieldNumber.update();
                            }
                            else if (column.type.toLowerCase() == 'choice') {
                                var fieldChoice = context.castTo(field, SP.FieldChoice);
                                if (column.choices) {
                                    fieldChoice.set_choices($.makeArray(column.choices));
                                }
                                if (column.format) {
                                    fieldChoice.set_editFormat(column.format);
                                }
                                fieldChoice.update();
                            }
                            else if (column.type.toLowerCase() == 'multichoice' && column.choices) {
                                var fieldChoice = context.castTo(field, SP.FieldMultiChoice);
                                fieldChoice.set_choices($.makeArray(column.choices));
                                fieldChoice.update();
                            }
                                /*else if (column.type.toLowerCase() == 'calculated' && column.formula) {
                                    var fieldCalculated = context.castTo(field, SP.FieldCalculated);
                                    fieldCalculated.set_formula(column.formula);
                                    fieldCalculated.update();
                                }*/
                            else if (column.type.toLowerCase() == 'datetime' && column.dateOnly) {
                                var fieldDateTime = context.castTo(field, SP.FieldDateTime);
                                fieldDateTime.set_displayFormat(SP.DateTimeFieldFormatType.dateOnly);
                                fieldDateTime.update();
                            }
                            else if (column.type.toLowerCase() == 'taxonomyfieldtypemulti') {
                                var fieldTaxonomy = context.castTo(field, SP.Taxonomy.TaxonomyField);
                                fieldTaxonomy.set_allowMultipleValues(true);
                                fieldTaxonomy.update();
                            }
                            else {
                                field.update();
                            }

                        });
                        scope.successMessage = 'Column ' + column.displayName + ' created';
                        scopes.push(scope);
                    });

                    helper.executeQuery(scopes, columnsCreated);
                };
                if (columns.filter(function (e) { return e.type == 'lookup'; }).length > 0) {
                    helper.updateListIds()
                    .done(CIB.installer.updateListIds())
                    .done(createColumns);
                }
                else {
                    createColumns();
                }

                return columnsCreated.promise();
            },

            hideContentTypeField: function (contentTypeId, fieldName) {
                var executepromise = $.Deferred();
                var scopes = [];
                var web = hostContext.get_web();
                var contentTypeCollection = web.get_contentTypes();
                context.load(contentTypeCollection);
                context.executeQueryAsyncPromise()
                    .done(function () {
                        var contentTypeEnumerator = contentTypeCollection.getEnumerator();
                        while (contentTypeEnumerator.moveNext()) {
                            var content = contentTypeEnumerator.get_current();
                            if (content.get_id() == contentTypeId)
                            {
                                var fields = content.get_fieldLinks();
                                context.load(fields);
                                context.executeQueryAsyncPromise().done(function () {
                                    var fieldEnumerator = fields.getEnumerator();
                                    while (fieldEnumerator.moveNext()) {
                                        var field = fieldEnumerator.get_current();
                                        var fid = field.get_id();
                                        var fname = field.get_name();
                                        if (field.get_name().toLowerCase() == fieldName.toLowerCase()) {

                                            var scope = $.handleExceptionsScope(context, function () {
                                                field.set_hidden(true);
                                                field.set_required(false);
                                                content.update();
                                            });

                                            scope.successMessage = 'Field: ' + fieldName + ' is hided from content type= ' + contentTypeId + '.';
                                            scopes.push(scope);
                                        }
                                    }
                                })
                                .fail(function (message) {
                                    CIB.installer.message('Error in content type: ' + message, 'error')
                                });
                                break;
                            }
                        }
                        executepromise.resolve();
                    })
                .fail(function (message) {
                    executepromise.reject();
                    CIB.installer.message('Error in outer query : ' + message, 'error')
                });

                helper.executeQuery(scopes, executepromise);

                return executepromise.promise();
            },

            /*
            allowManagementCT: function (listTitle) {
                var dfd = $.Deferred();
                var web = hostContext.get_web();
                var list = web.get_lists().getByTitle(listTitle);
                context.load(list);
                list.set_contentTypesEnabled(true);
                list.update();  //update operation is required to apply list changes
                context.executeQueryAsync(
                  function () {
                      dfd.resolve();
                  },
                  function (sender, args) {
                      dfd.reject(args.get_message());
                  }
                );
                return dfd.promise();
            };
            */

            /*
                Adds existing content types to a list
                @listTitle 'Documents'
                @contentTypeIds [ '0x0100C4AE7CEF4055486987E22766C23F7F35' ]
            */
            addContentTypesToList: function (listTitle, contentTypeIds) {
                var scopes = [];
                contentTypeIds = CIB.utilities.ensureArray(contentTypeIds);

                var contentTypesAdded = new jQuery.Deferred();

                helper.message('Adding content types to list \'' + listTitle + '\'');

                var web = hostContext.get_web();
                var list = web.get_lists().getByTitle(listTitle);

                list.set_contentTypesEnabled(true);

                var contentTypes = web.get_availableContentTypes(); // web.get_contentTypes();
                var listContentTypes = list.get_contentTypes();

                $.each(contentTypeIds, function () {
                    var contentTypeId = this;

                    var scope = $.handleExceptionsScope(context, function () {
                        var existingContentType = contentTypes.getById(contentTypeId);
                        listContentTypes.addExistingContentType(existingContentType);
                    });

                    scope.successMessage = 'Content type ' + contentTypeId + ' added to list';
                    scopes.push(scope);
                });

                helper.executeQuery(scopes, contentTypesAdded);

                return contentTypesAdded.promise();
            },
/*
            removeContentTypesFromList: function (listName, contentTypes) {
                var executepromise = $.Deferred();
                CIB.installer.message("Removing content types from list '" + listName + "'");
                contentTypes = CIB.utilities.ensureArray(contentTypes);

                var web = hostContext.get_web();
                var list = web.get_lists().getByTitle(listName);
                list.set_contentTypesEnabled(true);
                list.update();
                var listContentTypes = list.get_contentTypes();
                context.load(listContentTypes);
                context.executeQueryAsyncPromise().done(function () {
                    var objects = [];
                    $.each(contentTypes, function () {
                        var contentType = this;
                        var found = false;
                        var contentTypeEnumerator = listContentTypes.getEnumerator();
                        while (contentTypeEnumerator.moveNext()) {
                            var content = contentTypeEnumerator.get_current();
                            var name = null;
                            try { name = content.get_name().toLowerCase() } catch (ex) { }
                            if (name == contentType.toLowerCase()) {
                                objects.push(content);
                                found = true;
                            }
                        }
                        if (!found) {
                            CIB.installer.message('Could not find content type ' + contentType + ' in list.', 'info');
                        }
                    });

                    var scopes = [];
                    $.each(objects, function () {
                        var content = this;
                        var scope = $.handleExceptionsScope(context, function () {
                            content.deleteObject();
                        });
                        scope.successMessage = "Content type removed from list";
                        scopes.push(scope);
                    });
                    helper.executeQuery(scopes, executepromise);
                })
                .fail(function (message) {
                    executepromise.reject();
                    CIB.installer.message('Error removing content types from list: ' + message, 'error');
                });
                return executepromise.promise();
            },
*/
            setFieldVisibility: function(listTitle, fieldName) {
                var web = hostContext.get_web();
                var list = web.get_lists().getByTitle(listTitle);
                var field = list.get_fields().getByTitle(fieldName);
                field.setShowInDisplayForm(false);
                field.setShowInNewForm(false);
                field.setShowInEditForm(false);
                field.set_hidden(true);
                field.update();
                return context.executeQueryAsyncPromise()
                    .done(function () {
                        CIB.installer.message('Disabled ' + fieldName + ' in ' + listTitle);
                    })
                    .fail(function (message) {
                        CIB.installer.message('Error disabling ' + +fieldName + ' in ' + listTitle + ': ' + message);
                    });
            },

            removeView: function (listTitle, viewTitle) {
                var executepromise = $.Deferred();
                var list = hostContext.get_web().get_lists().getByTitle(listTitle);
                var views = list.get_views();
                context.load(views);
                return context.executeQueryAsyncPromise().done(function () {
                    var view;
                    var viewExists = false;
                    var viewsEnumerator = views.getEnumerator();
                    while (viewsEnumerator.moveNext()) {
                        view = viewsEnumerator.get_current();
                        if (view.get_title().toLowerCase() == viewTitle.toLowerCase()) {
                            viewExists = true;
                        }
                    }

                    if (!viewExists) {
                        CIB.installer.message('"' + viewTitle + '" view was not found. (expected if deleted already)', 'info');
                    }
                    else {
                        var view = list.get_views().getByTitle(viewTitle);
                        view.deleteObject();
                        return context.executeQueryAsyncPromise()
                            .done(function () {
                                CIB.installer.message(listTitle + ' - view ' + viewTitle + ' has been removed', 'success');
                                executepromise.resolve();
                            })
                            .fail(function (message) {
                                CIB.installer.message(listTitle + ' - error remove view ' + viewTitle + ' : ' + message, 'error');
                                executepromise.reject();
                            });
                    }
                })
                .fail(function (message) {
                    CIB.installer.message(listTitle + ' - error remove view ' + viewTitle + ' : ' + message, 'error');
                    executepromise.reject();
                });
                return executepromise.promise();
            }
        }
    }()

    ////////////////////////////////////////////////////
    // Content Types
    ////////////////////////////////////////////////////

    /*
        var createCurrencyContentType3 = function () {
            var executepromise = $.Deferred();
    //            var s = hostContext.get_web().get_fields();
                var web = hostContext.get_web();
                context.load(web);
                var hcv = hostContext.get_web().get_contentTypes();
                var ocv = hcv.getById(ctCurrency.id);
    //            var e = o.get_fieldLinks();
                context.load(hcv);
                context.load(ocv);
    //            context.load(e);
                context.executeQueryAsyncPromise().done(function ()
                {
                    var uf = web.get_url();
    //                if (ocv) {
    //                    var o2 = ocv.get_id();
    //                }
    
                    var contentTypeEnumerator = hcv.getEnumerator();
                    while (contentTypeEnumerator.moveNext()) {
                        var content = contentTypeEnumerator.get_current();
                        CIB.installer.message(content.get_name(), 'success');
                    }
    
    //                var c = e.getEnumerator();
    //                while (c.moveNext())
    //                {
    //                    var fi = c.get_current();
    //                    var id = fi.get_id();
    //                    var nm = fi.get_name();
    //                }
                    executepromise.resolve();
                })
                .fail(function (message) {
                    executepromise.reject();
                    CIB.installer.message('Error in outer query : '  + message, 'error')
                });
            return executepromise.promise();
        };
    
    
        var createCurrencyContentTypeSafe = function () {
            var executepromise = $.Deferred();
            if (context != undefined && context != null) {
                var web = hostContext.get_web();
                var contentTypeCollection = web.get_contentTypes();
                var contentType = null;
                context.load(contentTypeCollection);
                context.executeQueryAsyncPromise()
                   .done(function () {
                       var contentTypeEnumerator = contentTypeCollection.getEnumerator();
                       while (contentTypeEnumerator.moveNext()) {
                           var content = contentTypeEnumerator.get_current();
                           if (content.get_name() == siteCT.currency.name) {
                               var id = content.get_id();
                               contentType = contentTypeCollection.getById(content.get_id());
                               break;
                           }
                       }
    
                       if (contentType != null) {
                           window.alert('content type already exists');
                       }
                       else {
                           window.alert('create content type');
                           return $.whenSync(function () {
                              return CIB.installer.createContentTypes(siteCT.currency)
                              .then(function () {
                                  return CIB.installer.addColumnsToContentType(siteCT.currency.id, siteCT.currency.columnNames);
                              })
                              .fail(function (message) {
                                  CIB.installer.message('Error in creating the workflow Task content type: ' + message, 'error');
                                  executepromise.reject();
                              })
                              .done(function () {
                                  CIB.installer.message('workflow Task content type created');
                                  executepromise.resolve();
                              });
                          })
                       }
                   })
                .fail(function (message) {
                    executepromise.reject();
                    CIB.installer.message('Error in outer query : ' + message, 'error')
                });
            }
            return excutepromise.promise();
        };
    */

    var createCustomList = function (contentType, list) {
        return installer.createSiteColumns(contentType.newColumns ? contentType.newColumns : contentType.columns)
        .then(function () {
            return CIB.installer.createContentTypes(contentType);
        })
        .then(function () {
            return CIB.installer.addColumnsToContentType(contentType.id, contentType.columnNames);
        })
        .then(function () {
            return installer.hideContentTypeField(contentType.id, TITLE);
        })
        .then(function () {
            return CIB.installer.createLists(list)
        })
        .then(function () {
            return CIB.installer.updateListIds();
        })
        .then(function () {
            return installer.updateListIds();
        })
        .then(function () {
            return installer.setFieldVisibility(list.name, TITLE);
        })
        .then(function () {
            return installer.addContentTypesToList(list.name, list.contentTypes);
        })
        .then(function () {
            return CIB.installer.removeContentTypesFromList(list.name, ITEM);
        })
        .then(function () {
            return CIB.installer.createView(list.name, list.views.default.name, list.views.default.columns, list.views.default.query);
        })
        .then(function () {
            return installer.removeView(list.name, ALL_ITEMS);
        })
        .done(function () {
            CIB.installer.message(list.name + ' created');
        })
        .fail(function (message) {
            CIB.installer.message('Error creating list ' + list.name  + ': ' + message, 'error');
        });
    }

    var createConfigList = function () {
        return createCustomList(siteCT.config, siteLists.config);
    }

    var createCurrenciesList = function () {
        return createCustomList(siteCT.currency, siteLists.currency);
    }

    var createTeamsList = function () {
        return createCustomList(siteCT.team, siteLists.team);
    }

    var createPriceList = function () {
        return createCustomList(siteCT.priceList, siteLists.priceList);
    }

    var createFeesList = function () {
        return createCustomList(siteCT.fee, siteLists.fee);
    }

/*

    var createWorkflowHistoryList = function () {
        return CIB.installer.createLists({
            name: 'Workflow History', type: 140, hidden: true
        }).then(function () {
            return CIB.installer.updateListIds();
        })
                 .fail(function (message) {
                     CIB.installer.message('Error in creating the workflow history list: ' + message, 'error');
                 })
            .done(function () {
                CIB.installer.message('Workflow history list created.');
            });
    };

    //create tasks list
    var createTasksList = function () {
        return CIB.installer.createLists({ name: 'Tasks', type: 107 })
            .then(function () {
                return CIB.installer.updateListIds();
            })
           .fail(function (message) {
               CIB.installer.message('Error in creating the workflow tasks list: ' + message, 'error');
           })
            .done(function () {

                CIB.installer.message('Workflow tasks list created.');
            });
    };
*/


/* 
C:\TFS\CIB.Apps\LivelinkMigration\Tool\OFF\CIB.LiveLinkProvision\CIB.LiveLinkProvision\ExitProcess\ExitScript\

    var setActiveEmployeeGIDLeftValidationFormula = function () {
        var dfd = $.Deferred();
        setFieldValidationFormula(CIB.LiveLinkEXIT.Framework.ListNames.EXITActiveEmployees.Name,
            CIB.LiveLinkEXIT.Framework.ColumnNames.ExitsTransfers_x002d_GID.Name,
            "=IF(OR(LEN(GID)=6,LEN(GID)=7),ISNUMBER(RIGHT(GID,5)*1),FALSE)",
            "Invalid Global ID"
            ).done(function () {
                dfd.resolve();
            }).fail(function (error) {
                console.error(error);
                dfd.reject();
            });

        return dfd.promise();
    };

    var updatelistContentTypeAllowManagement = function () {
        var dfd = $.Deferred();
        allowManagementCT(CIB.LiveLinkEXIT.Framework.ListNames.EXITActiveEmployees.Name).done(function () {
            dfd.resolve();
        }).fail(function (context, sender, args) {
            console.error(args.get_message());
            dfd.reject();
        });

        allowManagementCT(CIB.LiveLinkEXIT.Framework.ListNames.EXITTransfers.Name).done(function () {
            dfd.resolve();
        }).fail(function (context, sender, args) {
            console.error(args.get_message());
            dfd.reject();
        });

        allowManagementCT(CIB.LiveLinkEXIT.Framework.ListNames.EXITPOTasks.Name).done(function () {
            dfd.resolve();
        }).fail(function (context, sender, args) {
            console.error(args.get_message());
            dfd.reject();
        });

        allowManagementCT(CIB.LiveLinkEXIT.Framework.ListNames.ProcessOwners.Name).done(function () {
            dfd.resolve();
        }).fail(function (context, sender, args) {
            console.error(args.get_message());
            dfd.reject();
        });


        return dfd.promise();
    };

    var enableVersion = function (listTitle) {
        var dfd = $.Deferred();
        var web = hostContext.get_web();
        var list = web.get_lists().getByTitle(listTitle);
        context.load(list);
        list.set_enableVersioning(true);
        list.update();  //update operation is required to apply list changes
        context.load(list);
        context.executeQueryAsync(
          function () {
              dfd.resolve();
          },
          function (sender, args) {
              dfd.reject(args.get_message());
          }
        );
        return dfd.promise();
    };

    var updateListEnableVersion = function () {
        var dfd = $.Deferred();
        enableVersion(CIB.LiveLinkEXIT.Framework.ListNames.EXITActiveEmployees.Name).done(function () {
            dfd.resolve();
        }).fail(function (error) {
            console.error(error);
            dfd.reject();
        });

        enableVersion(CIB.LiveLinkEXIT.Framework.ListNames.EXITTransfers.Name).done(function () {
            dfd.resolve();
        }).fail(function (error) {
            console.error(error);
            dfd.reject();
        });

        enableVersion(CIB.LiveLinkEXIT.Framework.ListNames.ProcessOwners.Name).done(function () {
            dfd.resolve();
        }).fail(function (error) {
            console.error(error);
            dfd.reject();
        });
        return dfd.promise();
    };

    var updateListFieldTitle = function () {
        var dfd = $.Deferred();
        UpdateFieldTitle(CIB.LiveLinkEXIT.Framework.ListNames.EXITPOTasks.Name,
            CIB.LiveLinkEXIT.Framework.ColumnNames.Status.Name,
            CIB.LiveLinkEXIT.Framework.ColumnNames.Status.DisplayName
            ).done(function () {
                dfd.resolve();
            }).fail(function (error) {
                console.error(error);
                dfd.reject();
            });
        return dfd.promise();
    }

    var UpdateFieldTitle = function (listTitle, oldFieldTitle, newFieldTitle) {
        var dfd = $.Deferred();
        var web = hostContext.get_web();
        var list = web.get_lists().getByTitle(listTitle);
        context.load(list);
        var fieldCollection = list.get_fields();
        var onField = fieldCollection.getByInternalNameOrTitle(oldFieldTitle);
        onField.set_title(newFieldTitle);
        onField.update();
        context.load(onField);
        context.executeQueryAsync(
          function () {
              CIB.installer.message('List field title updated', 'success');
              dfd.resolve();
          },
          function (sender, args) {
              CIB.installer.message('List field title update failed', error);
              dfd.reject(args.get_message());
          }
        );
        return dfd.promise();
    };

    var UpdateFieldDefaultValue = function (listTitle, oldFieldTitle, defaultValue) {
        var dfd = $.Deferred();
        var web = hostContext.get_web();
        var list = web.get_lists().getByTitle(listTitle);
        context.load(list);
        var fieldCollection = list.get_fields();
        var onField = fieldCollection.getByInternalNameOrTitle(oldFieldTitle);
        onField.set_defaultValue(defaultValue);
        onField.update();
        context.load(onField);
        context.executeQueryAsync(
          function () {
              CIB.installer.message('List field default value updated', 'success');
              dfd.resolve();
          },
          function (sender, args) {
              CIB.installer.message('List field default value update failed', error);
              dfd.reject(args.get_message());
          }
        );
        return dfd.promise();
    };


    var updateListFieldDefaultValue = function () {
        var dfd = $.Deferred();
        UpdateFieldDefaultValue(CIB.LiveLinkEXIT.Framework.ListNames.EXITPOTasks.Name,
            CIB.LiveLinkEXIT.Framework.ColumnNames.PercentComplete.Name,
            "0"
            ).done(function () {
                dfd.resolve();
            }).fail(function (error) {
                console.error(error);
                dfd.reject();
            });
        return dfd.promise();
    }


*/

    var loadContext = function () {
        var url;
        console.log("Entering LoadContext");
        url = $.getServerRealtiveHostWebUrl();
        var newHostContext = new SP.ClientContext(url);
        hostWeb = newHostContext.get_web();
        newHostContext.load(hostWeb);
        console.log("Leaving LoadContext");
        return newHostContext.executeQueryAsyncPromise().
            then(function () {
                var excutepromise = $.Deferred();
                globalContext = CIB.utilities.getContext();
                context = globalContext.context;
                hostContext = globalContext.hostContext;
                excutepromise.resolve();
                return excutepromise.promise();
            });
    };

    //reinitialize context for subsites - on change of dropdown values
    var refreshContext = function () {
        globalContext = CIB.utilities.getContext();
        context = globalContext.context;
        hostContext = globalContext.hostContext;
        jQuery.getScript(document.querySelector('script[src$="Installer.js"]').getAttribute('src'), function () {
        }, true);
    }

    $(document).ready(function () {

        globalContext = CIB.utilities.getContext();
        context = globalContext.context;
        hostContext = globalContext.hostContext;

        // Retrives all the sub sites
        retrieveSubsites();

        $("#drp_subsite").change(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                document.getElementById("validate-msg").style.visibility = "hidden";
                $('#CIBAppFrameWorkSubWebUrl').html($("#drp_subsite option:selected").val());
            }
            refreshContext();
        });

        $('#install-lists').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    createConfigList,
                    createCurrenciesList,
                    createTeamsList,
                    createPriceList,
                    createFeesList)
            .done(function () { CIB.installer.message('lists are installed.', 'success'); })
            .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });
    });
}();

</script>


</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    CommandPoint App Installation Page
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <div>
        Please select a site you would like to install this application on:
       <br />
        <br />
        <select id="drp_subsite" style="width: 150px">
            <option selected="selected">-select site-</option>
        </select>
        <br />
        <div id="validate-msg"></div>
        <br />
        <br />
        <button id="install-lists" type="button" class="btn btn-success" data-loading-text="Install Lists">1. Install Lists</button>
        <br />
        <br />
<%--
        <button id="install-forms" type="button" class="btn btn-success" data-loading-text="Installing forms...">2. Install list forms</button>
        <br />
        <br />
        <button id="install-views" type="button" class="btn btn-success" data-loading-text="Installing views...">3. Install list views</button>
        <br />
        <br />
        <button id="installWorkflows" type="button" class="btn btn-success" data-loading-text="Installing workflows...">4. Install Workflows</button>
--%>		
        <br />
        <div id="install-status">
        </div>
        <div id="CIBAppFrameWorkSubWebUrl" style="display: none"></div>
    </div>
</asp:Content>
