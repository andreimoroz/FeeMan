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
<%--	
    <CIB:CommonStyleSheet runat="server" Path="lib/bootstrap.min.css" />
    <CIB:CommonStyleSheet runat="server" Path="lib/bootstrap-theme.min.css" />
--%>	
    <CIB:CommonStyleSheet runat="server" Path="Common/App.css" />
<%--	
    <CIB:CommonScript runat="server" Path="lib/jquery-1.8.2.min.js" />
    <CIB:CommonScript runat="server" Path="lib/bootstrap.min.js" />
--%>	

    <script type="text/javascript" src="../SiteAssets/Utilities.js"></script>
	<script type="text/javascript" src="../SiteAssets/Installer.js"></script>
	<script type="text/javascript" src="../SiteAssets/Logger.js"></script>

	    <!-- Add your JavaScript to the following file -->
<%--     <script type="text/javascript" src="Install.js"></script> --%>

<script type="text/javascript" >


"use strict";
var FeeMan = FeeMan || {};
FeeMan.app = FeeMan.app || {};
FeeMan.app.appdisplay = function () {

    var ctx;

    /////////////////////////////////////////
    // Config
    /////////////////////////////////////////
    ////// ************ //////////////////////////
    var testAddId = '0';
    var testAddName = '';
    ////// ************ //////////////////////////
    var TITLE = 'Title';
    var ITEM = 'Item';
    var ALL_ITEMS = 'All Items';
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
        feePlvOperationCode: {
            name: 'Operation Code' + testAddName,
            displayName: 'Operation Code' + testAddName,
            type: 'Other',
        },
        feeAccount: {
            name: 'fmAccount' + testAddName,
            displayName: 'Customer Account Number' + testAddName,
            id: '{6f4338d7-d741-4058-8d09-c2dfeec3049' + testAddId + '}',
            type: 'Text',
            maxLength: '23',
            required: true,
            linkToItem: 'Required',
            validationMessage: "Account number must be like: digits BLANK 6 digits BLANK 3 digits BLANK 2 digits BLANK then three characters for the currency",
            validationFormula: "=(LEN([Customer Account Number" + testAddName + "])=23)",
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
        feeUnits: {
            name: 'fmUnits' + testAddName,
            displayName: 'Units' + testAddName,
            id: '{957d4ece-bc38-4f75-a08d-6b5f962b42e' + testAddId + '}',
            type: 'Number',
            decimals: '0',
            required: true,
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
            choices: ["Draft", "Invalid", "Requested", "Approved", "Exporting", "Exported"],
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
            choices: ["Draft", "Request", "Cancel"],
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
        feeExportedAt: {
            name: 'fmExportedAt' + testAddName,
            displayName: 'ExportedAt' + testAddName,
            id: '{169b6611-7a62-461a-987f-b062d1e0026' + testAddId + '}',
            type: 'DateTime',
            showInEditForm: false,
            showInNewForm: false,
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
        archiveTitle: {
            name: 'fmArchiveTitle' + testAddName,
            displayName: 'Customer Account Number / Title' + testAddName,
            id: '{0bde72da-df12-42b6-bb3e-4b14cdaa6bb' + testAddId + '}',
            type: 'Calculated',
            formula: '=IF(ISBLANK(Title),[fmAccount' + testAddName + '],Title)',
            resultType: 'Text',
            showInEditForm: false,
            showInNewForm: false,
            readOnly: true,
            fieldRefs: ['Title', 'fmAccount' + testAddName],
            linkToItem: 'Required',
            group: groupName
        },
        archiveStatus: {
            name: 'fmArchiveStatus' + testAddName,
            displayName: 'Status' + testAddName,
            id: '{bb0e6664-2312-450c-94a2-56f41d0c21a' + testAddId + '}',
            type: 'Choice',
            format: SP.ChoiceFormatType.dropdown,
            fillInChoice: false,
            showInEditForm: false,
            showInNewForm: false,
            choices: ["Validated", "Cancelled", "Rejected"],
            required: true,
            group: groupName
        },
        archiveValidatedBy: {
            name: 'fmValidatedBy' + testAddName,
            displayName: 'ValidatedBy' + testAddName,
            id: '{d7bd4bd2-c992-4c05-8e27-5434691b734' + testAddId + '}',
            type: 'User',
            userSelectionMode: 'PeopleOnly',
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        archiveCreatedAt: {
            name: 'fmCreatedAt' + testAddName,
            displayName: 'CreatedAt' + testAddName,
            id: '{aadb733e-f040-4559-8fb1-6e0a2b1ef77' + testAddId + '}',
            type: 'DateTime',
            showInEditForm: false,
            showInNewForm: false,
            group: groupName
        },
        archiveItemID: {
            name: 'fmItemID' + testAddName,
            displayName: 'ItemID' + testAddName,
            id: '{a475331e-40cf-485e-bcc0-f8eb0097f4f' + testAddId + '}',
            type: 'Number',
            decimals: '0',
            required: true,
            group: groupName
        }
    };

    var siteCT = {
        config: {
            name: 'CIB Feem Config' + testAddName,
            id: '0x01009FC47BE0766A4455B10560438A54CBD' + testAddId,
            group: groupName,
            columns: [siteColumns.teamMembers],
            columnNames: [siteColumns.teamMembers.name],
            deleteTitle: true
        },
        currency: {
            name: 'CIB Feem Currencies' + testAddName,
            id: '0x0100F34E70C088364B1B85B9462BA830A28' + testAddId,
            group: groupName,
            columns: [siteColumns.currencyCode],
            columnNames: [siteColumns.currencyCode.name],
            deleteTitle: true
        },
        team: {
            name: 'CIB Feem Teams' + testAddName,
            id: '0x01007DDEF9578DB3451495FF451DF3B9539' + testAddId,
            group: groupName,
            columns: [siteColumns.teamName, siteColumns.teamMembers],
            columnNames: [siteColumns.teamName.name, siteColumns.teamMembers.name],
            newColumns: [siteColumns.teamName],
            deleteTitle: true
        },
        priceList: {
            name: 'CIB Feem Price List' + testAddName,
            id: '0x0100DB215693108F478FA58825DE78BEB51' + testAddId,
            group: groupName,
            columns: [siteColumns.plvPlvDescription, siteColumns.plvCode, siteColumns.plvStandardPrice, siteColumns.plvDescription, siteColumns.plvTeams],
            columnNames: [siteColumns.plvPlvDescription.name, siteColumns.plvCode.name, siteColumns.plvStandardPrice.name, siteColumns.plvDescription.name, siteColumns.plvTeams.name],
            deleteTitle: true
        },
        fee: {
            name: 'CIB Feem Fees' + testAddName,
            id: '0x010091DD1E064DBC4B9AA0D9B0E1000FDB1' + testAddId,
            group: groupName,
            columns: [siteColumns.feeTeam, siteColumns.feePlvDescription, siteColumns.feeAccount, siteColumns.feeClientName,
                siteColumns.feeTransactionAmount, siteColumns.feeCurrency, siteColumns.feeUnits, siteColumns.feeAmount, siteColumns.feeComments, siteColumns.feeStatus, siteColumns.feeStatus_,
                siteColumns.feeRequest, siteColumns.feeRequestedBy, siteColumns.feeApprove, siteColumns.feeApprovedAt, siteColumns.feeApprovedBy, siteColumns.feeValidate, siteColumns.feeExportedAt,
                siteColumns.feeStatusHistory, siteColumns.feeWorkflowActive],
            columnNames: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.feeStatus.name, siteColumns.feeStatus_.name,
                siteColumns.feeRequest.name, siteColumns.feeRequestedBy.name, siteColumns.feeApprove.name, siteColumns.feeApprovedAt.name, siteColumns.feeApprovedBy.name, siteColumns.feeValidate.name, siteColumns.feeExportedAt.name,
                siteColumns.feeStatusHistory.name, siteColumns.feeWorkflowActive.name],
            deleteTitle: true
        },
        archive: {
            name: 'CIB Feem Archive' + testAddName,
            id: '0x0100CADB99084D394B7B91DB7222C02B934' + testAddId,
            group: groupName,
            columns: [siteColumns.feeTeam, siteColumns.feePlvDescription, siteColumns.feeAccount, siteColumns.feeClientName,
                siteColumns.feeTransactionAmount, siteColumns.feeCurrency, siteColumns.feeUnits, siteColumns.feeAmount, siteColumns.feeComments, siteColumns.archiveStatus,
                siteColumns.feeRequestedBy, siteColumns.feeApprovedAt, siteColumns.feeApprovedBy, siteColumns.archiveValidatedBy, siteColumns.feeExportedAt, siteColumns.feeStatusHistory, siteColumns.archiveCreatedAt, siteColumns.archiveItemID, siteColumns.archiveTitle],
            columnNames: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.archiveStatus.name,
                siteColumns.feeRequestedBy.name, siteColumns.feeApprovedAt.name, siteColumns.feeApprovedBy.name, siteColumns.archiveValidatedBy.name, siteColumns.feeExportedAt.name, siteColumns.feeStatusHistory.name, siteColumns.archiveCreatedAt.name, siteColumns.archiveItemID.name, siteColumns.archiveTitle.name],
            newColumns: [siteColumns.archiveStatus, siteColumns.archiveValidatedBy, siteColumns.archiveCreatedAt, siteColumns.archiveItemID, siteColumns.archiveTitle],
            deleteTitle: false
        }                           
    };

    var siteLists = {
        config: {
            name: 'Config' + testAddName,
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.config.id],
            views: [ { name: 'Default', columns: siteCT.config.columnNames, paged: true, setAsDefaultView: true } ],
            items: { columns: [ siteColumns.teamMembers.name ], values: [['10;#Andrei MOROZ;#20;#DEV Andrei MOROZ;#11;#Jochen FUNK;#16;#Grzegorz KUBACKI;#19;#Gabriel OLTEANU;#17;#Malgorzata JANKOWSKA;#18;#Aleksandra JAROSINSKA-CHOLEWA']]}
        },
        currency: {
            name: 'Currencies' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.currency.id],
            views: [{ name: 'Default', columns: siteCT.currency.columnNames, paged: true, query: '<OrderBy><FieldRef Name="' + siteColumns.currencyCode.name + '" Ascending="TRUE"></FieldRef></OrderBy>', setAsDefaultView: true }],
            items: { columns: [ siteColumns.currencyCode.name ], values: [['EUR'], ['USD'], ['GBP']] }
        },
        team: {
            name: 'Teams' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.team.id],
            views: [{ name: 'Default', columns: siteCT.team.columnNames, paged: true, query: '<OrderBy><FieldRef Name="' + siteColumns.teamName.name + '" Ascending="TRUE"></FieldRef></OrderBy>', setAsDefaultView: true }],
            items: { columns: [siteColumns.teamName.name, siteColumns.teamMembers.name], 
                values: [
                    ['CMTO', '10;#Andrei MOROZ;#20;#DEV Andrei MOROZ;#11;#Jochen FUNK;#16;#Grzegorz KUBACKI;#19;#Gabriel OLTEANU;#17;#Malgorzata JANKOWSKA;#18;#Aleksandra JAROSINSKA-CHOLEWA'],
                    ['CSD', '10;#Andrei MOROZ;#11;#Jochen FUNK;#16;#Grzegorz KUBACKI;#19;#Gabriel OLTEANU;#17;#Malgorzata JANKOWSKA;#18;#Aleksandra JAROSINSKA-CHOLEWA'],
                    ['CR', null], ['DISPO', null], ['FX/MM/CHEQUES', null]
                ]
            }
        },
        priceList: {
            name: 'PriceList' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.priceList.id],
            views: [{ name: 'Default', columns: siteCT.priceList.columnNames, paged: true, query: '<OrderBy><FieldRef Name="' + siteColumns.plvPlvDescription.name + '" Ascending="TRUE"></FieldRef></OrderBy>', setAsDefaultView: true }],
            items: { columns: [siteColumns.plvPlvDescription.name, siteColumns.plvCode.name, siteColumns.plvStandardPrice.name, siteColumns.plvDescription.name, siteColumns.plvTeams.name],
                values: [
                    ['CMTO - Connexis Cash - Implementation fee - Delegated user administration', 'CXCIMPDG', '1000', 'Connexis Cash - Implementation', '1;#CMTO'],
                    ['CMTO - Connexis Cash - Implementation fee', 'CXCIMPLE', '2000', 'Connexis Cash - Implementation', '1;#CMTO'],
                    ['CMTO - Connexis Cash - Other services - Security token', 'CXCTOKEN', '15', 'Connexis Cash - Other services', '1;#CMTO'],
                    ['CMTO - Connexis Cash - Additional training session - Remote', 'CXCTRARM', '200', 'Connexis Cash - Implementation', '1;#CMTO'],
                    ['CMTO - Connexis Cash - Additional training session - At client site', 'CXCTRASI', '500', 'Connexis Cash - Implementation', '1;#CMTO'],
                    ['CMTO - Connexis Cash - Other services - Block / unblock user', 'CXCUBLCK', '15', 'Connexis Cash - Other services', '1;#CMTO'],
                    ['CMTO - Connexis Cash - Other services - Update user token', 'CXCUPDTK', '15', 'Connexis Cash - Other services', '1;#CMTO'],
                    ['CMTO - Connexis Gateway - Implementation fee', 'CXGAIMPL', '3000', 'Connexis Gateway', '1;#CMTO'],
                    ['CSD - Production of duplicate document upon customer request', 'DUPLICAT', '15', 'Statements and mailing', '2;#CSD'],
                    ['CMTO - Account statement with 3rd party bank - Implementation fee', 'DX3PIMP', '150', 'Account statement', '1;#CMTO'],
                    ['CMTO - Debit / Credit advice', 'DXDCADV', '30', 'Transaction reporting', '1;#CMTO'],
                    ['CMTO - E-Link - Other services - Electronic certificate', 'ELINCERT', '100', 'E-Link - Other services', '1;#CMTO'],
                    ['CMTO - E-Link - Implementation fee', 'ELINKIMP', '2000', 'E-Link implementation', '1;#CMTO'],
                    ['CMTO - Global EBICS - implementation fee', 'GEBICIMP', '2000', 'Global EBICS implementation', '1;#CMTO'],
                    ['CMTO - Local channel - Implementation fee', 'LOCACIMP', '250', 'EBICS', '1;#CMTO'],
                    ['CMTO - Local channel – Implementation Fee (at client site)', 'LOCAPT', '', 'Local channel Implementation', '1;#CMTO'],
                    ['CMTO - Notional pooling - Interest optimisation - Implementation fee', 'NPIIMPLE', '150', 'Interest enhancement / optimisation', '1;#CMTO'],
                    ['CMTO - Notional pooling - Interest optimisation - Structural adjustment', 'NPILADJU', '500', 'Interest enhancement / optimisation', '1;#CMTO'],
                    ['CMTO - Notional pooling - Interest optimisation - Small adjustment', 'NPISADJU', '100', 'Interest enhancement / optimisation', '1;#CMTO'],
                    ['CMTO - Physical pooling - Implementation fee', 'PPCCIMPL', '150', 'Physical pooling - Cash concentration', '1;#CMTO'],
                    ['CMTO - Physical pooling - Structural adjustment', 'PPCCLADJ', '500', 'Physical pooling - Cash concentration', '1;#CMTO'],
                    ['CMTO - Physical pooling - Small adjustment', 'PPCCSADJ', '100', 'Physical pooling - Cash concentration', '1;#CMTO'],
                    ['CMTO - SWIFTNET - Implementation fee - FileAct', 'SWFACIMP', '2000', 'SWIFTNET implementation', '1;#CMTO'],
                    ['CMTO - SWIFTNET - Other services - Electronic certificate', 'SWFECERT', '100', 'SWIFTNET - Other services', '1;#CMTO'],
                    ['CMTO - SWIFTNET - Implementation fee - FIN and FileAct', 'SWFFAIMP', '3000', 'SWIFTNET implementation', '1;#CMTO'],
                    ['CMTO - SWIFTNET - Implementation fee - FIN', 'SWFINIMP', '2000', 'SWIFTNET implementation', '1;#CMTO'],
                    ['CR - Audit certificate upon customer request', 'CERTAUDI', '150', 'Certification', '3;#CR'],
                    ['CR - Other certificate upon customer request', 'CERTOTHR', '100', 'Certification', '3;#CR'],
                    ['CR - Virtual IBAN - Implementation fee', 'CLVACIMP', '1', 'Virtual IBAN implementation', '3;#CR'],
                    ['CSD - Account investigation / update upon customer request', 'ACCTINVG', '25', 'Account', '2;#CSD'],
                    ['CSD - Corporate card - Implementation fee', 'COCDIMPL', '250', 'Corporate card - Card fees', '2;#CSD'],
                    ['CSD - Investigation upon customer request', 'MAINVSGT', '25', 'Client services support', '2;#CSD'],
                    ['CSD - Client services support upon customer request (per hour)', 'SUPORTLG', '75', 'Client services support', '2;#CSD'],
                    ['CSD - Other manual services outside of account maintenance and payment services upon customer request', 'SUPOTHER', '10', 'Client services support', '2;#CSD'],
                    ['CSD/ DISPO - Investigation of the customer´s new address due to account maintenance and payment services', 'ACCINADD', '25', 'Account', '2;#CSD;#4;#DISPO'],
                    ['DISPO - Manual intervention - Non-formatted order', 'MANOFORM', '25', 'Repair and non-STP services', '4;#DISPO'],
                    ['DISPO - Cancel', 'MACANCEL', '15', 'Repair and non-STP services', '4;#DISPO'],
                    ['CSD - Standing order - Implementation fee', 'TSPSOIMP', '10', 'Processing services', '2;#CSD'],
                    ['DISPO - Notification about legitimate refusal to execute a payment order', 'TSPRFNOT', '5', 'Processing services', '4;#DISPO'],
                    ['FX/MM/CHEQUES - Cheque/Promissory note - Manual intervention - Non formatted order', 'CHQMANNF', '25', 'heque and promissory note payment services', '5;#FX/MM/CHEQUES'],
                    ['FX/MM/CHEQUES - Cheque returned to the bank - Notification (amount equal to or higher than 6000 EUR)', 'CHQRETNO', '25', 'heque and promissory note payment services', '5;#FX/MM/CHEQUES'],
                    ['FX/MM/CHEQUES - Cheque returned to the bank - Processing fee', 'CHQRET', '25', 'heque and promissory note payment services', '5;#FX/MM/CHEQUES'],
                    ['FX/MM/CHEQUES - Cheque / Promissory note - Stop payment upon customer request', 'CHQSTCUS', '25', 'heque and promissory note payment services', '5;#FX/MM/CHEQUES'],
                    ['CSD - Confirmation upon customer request', 'TSPCONFI', '25', 'Processing services', '2;#CSD'],
                    ['FX/MM/CHEQUES - Issuing of cheque book', 'PYCHBOOK', '25', 'Cheque', '5;#FX/MM/CHEQUES'],
                    ['FX/MM/CHEQUES - Lockbox - Implementation fee', 'CLLBXIMP', '250', 'Cheque', '5;#FX/MM/CHEQUES'],
                    ['FX/MM/CHEQUES - Lockbox - Processing fee', 'CLLBXPRO', '5', 'Cheque', '5;#FX/MM/CHEQUES']
                ]
            }
        },
        fee: {
            name: 'Fees' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.fee.id],
            views: [ 
                {
                    name: 'Default',
                    columns: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feePlvOperationCode.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.feeStatus.name, siteColumns.feeStatus_.name,
                        siteColumns.feeRequest.name, siteColumns.feeRequestedBy.name, siteColumns.feeApprove.name, siteColumns.feeApprovedAt.name, siteColumns.feeApprovedBy.name, siteColumns.feeValidate.name, siteColumns.feeExportedAt.name,
                        siteColumns.feeStatusHistory.name, siteColumns.feeWorkflowActive.name, 'ID'],
                    query: '<OrderBy><FieldRef Name="ID" Ascending="FALSE"></FieldRef></OrderBy>',
                    paged: true,
                    setAsDefaultView: true
                },
                {
                    name: 'Request View',
                    columns: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feePlvOperationCode.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.feeStatus_.name,
                        siteColumns.feeRequestedBy.name, siteColumns.feeRequest.name],
                    query: '<Where><And><Eq><FieldRef Name="fmWorkflowActive' + testAddName + '"/><Value Type="Text">-</Value></Eq>' +
                        '<Or><IsNull><FieldRef Name="fmStatus' + testAddName + '"/></IsNull>' +
                          '<Or><Or><Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Draft</Value></Eq>' +
                          '<Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Invalid</Value></Eq>' +
                        '</Or><Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Requested</Value></Eq></Or></Or>' +
                        '</And></Where>' +
                    '<OrderBy><FieldRef Name="ID" Ascending="FALSE"></FieldRef></OrderBy>',
                    paged: true
                },
                {
                    name: 'Approve View',
                    columns: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feePlvOperationCode.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name,
                        siteColumns.feeRequestedBy.name, siteColumns.feeStatus_.name, siteColumns.feeApprove.name],
                    query: '<Where><And>' +
                            '<Eq><FieldRef Name="fmWorkflowActive' + testAddName + '"/><Value Type="Text">-</Value></Eq>' +
                            '<Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Requested</Value></Eq>' +
                        '</And></Where>' + 
                        '<OrderBy><FieldRef Name="ID" Ascending="FALSE"></FieldRef></OrderBy>',
                    paged: true,
                    viewType: SP.ViewType.grid
                },
                {
                    name: 'Export View',
                    columns: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feePlvOperationCode.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name,
                        siteColumns.feeRequestedBy.name, siteColumns.feeStatus_.name, siteColumns.feeApprove.name],
                    query: '<Where><And><Eq><FieldRef Name="fmWorkflowActive' + testAddName + '"/><Value Type="Text">-</Value></Eq>' +
                              '<Or><Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Approved</Value></Eq>' +
                              '<Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Exporting</Value></Eq>' +
                              '</Or></And></Where>' +
                          '<OrderBy><FieldRef Name="ID" Ascending="FALSE"></FieldRef></OrderBy>',
                    paged: true
                },
                {
                    name: 'Validate View',
                    columns: [siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.feePlvOperationCode.name, siteColumns.feeAccount.name, siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name,
                        siteColumns.feeRequestedBy.name, siteColumns.feeApprovedBy.name, siteColumns.feeStatus_.name, siteColumns.feeValidate.name],
                    query: '<Where><And><Eq><FieldRef Name="fmWorkflowActive' + testAddName + '"/><Value Type="Text">-</Value></Eq>' +
                        '<Eq><FieldRef Name="fmStatus' + testAddName + '"/><Value Type="Text">Exported</Value></Eq></And></Where>' +
                        '<OrderBy><FieldRef Name="ID" Ascending="FALSE"></FieldRef></OrderBy>',
                    paged: true
                }
            ],
            workflows: {
                feeUpdated: {
                    name: "FeeUpdated",
                    startOnCreate: true,
                    startOnChange: true,
                    startManually: true,
                    taskList: "Tasks",
                    historyList: "Workflow History",
                    xaml: "/SiteAssets/FeeMan/FeeUpdated.xamlw"
                }
            }
        },
        archive: {
            name: 'Archive' + testAddName, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            enableFolderCreation: true,
            contentTypes: [siteCT.archive.id, '0x0120'],
            views: [ 
                {
                    name: 'Default',
                    columns: ['DocIcon', /*'LinkTitle',*/ siteColumns.feeTeam.name, siteColumns.feePlvDescription.name, siteColumns.archiveTitle.name, /*siteColumns.feeAccount.name,*/ siteColumns.feeClientName.name,
                        siteColumns.feeTransactionAmount.name, siteColumns.feeCurrency.name, siteColumns.feeUnits.name, siteColumns.feeAmount.name, siteColumns.feeComments.name, siteColumns.archiveStatus.name,
                        siteColumns.feeRequestedBy.name, /*siteColumns.feeApprovedAt.name,*/ siteColumns.feeApprovedBy.name, siteColumns.archiveValidatedBy.name, siteColumns.archiveItemID.name, 'ID'],
                    query: '<OrderBy><FieldRef Name="Title" Ascending="FALSE" /><FieldRef Name="ID" Ascending="FALSE" /></OrderBy>',
                    setAsDefaultView: true
                }
            ]
        }
    };

    ////////////////////////////////////////////////////
    // Copy files
    ////////////////////////////////////////////////////
    var copyFiles = function () {
		return CIB.DE.installer.createFolders([{ name: 'FeeMan', list: 'Site Assets', path: 'SiteAssets' }])
		.then(function() {
			return CIB.DE.installer.copyFiles([
				{ name: 'App.js', url: 'SiteAssets/FeeMan', path: 'SiteAssets/FeeMan/App.js', publish: false, binary: true },
				{ name: 'Blob.js', url: 'SiteAssets/FeeMan', path: 'SiteAssets/FeeMan/Blob.js', publish: false, binary: true },
				{ name: 'FileSaver.js', url: 'SiteAssets/FeeMan', path: 'SiteAssets/FeeMan/FileSaver.js', publish: false, binary: true },
				{ name: 'xlsx.full.min.js', url: 'SiteAssets/FeeMan', path: 'SiteAssets/FeeMan/xlsx.full.min.js', publish: false, binary: true, useRequestor: true }
			]);
		})
        .done(function () {
            CIB.DE.installer.message('Files successfully copied.', 'success')
        })
        .fail(function (message) {
            CIB.DE.installer.message('Error copying files to Site Assets: ' + message, 'error')
        });
    };

    ////////////////////////////////////////////////////
    // Lists
    ////////////////////////////////////////////////////
    var createCustomList = function (contentType, list) {
        return CIB.DE.installer.createSiteColumns(contentType.newColumns ? contentType.newColumns : contentType.columns)
        .then(function () {
            return CIB.DE.installer.createContentTypes(contentType);
        })
        .then(function () {
            return CIB.DE.installer.addColumnsToContentType(contentType.id, contentType.columnNames);
        })
        .then(function () {
            return contentType.deleteTitle ? CIB.DE.installer.removeContentTypeField(contentType.id, TITLE) : (new $.Deferred()).resolve();
        })
        .then(function () {
            return CIB.DE.installer.createLists(list);
        })
        .then(function () {
            return list.enableFolderCreation ? CIB.DE.installer.enableFolders(list.name, list.enableFolderCreation) : (new $.Deferred()).resolve();
        })
        .then(function () {
            return CIB.DE.installer.updateListIds();
        })
        .then(function () {
            return contentType.deleteTitle ? CIB.DE.installer.hideFieldFromList(list.name, TITLE) : (new $.Deferred()).resolve();
        })
        .then(function () {
            return CIB.DE.installer.addContentTypesToList(list.name, list.contentTypes);
        })
        .then(function () {
            return CIB.DE.installer.removeContentTypesFromList(list.name, ITEM);
        })
        .then(function () {
            return CIB.DE.installer.createViews(list.name, list.views);
        })
        .then(function () {
            return CIB.DE.installer.removeView(list.name, ALL_ITEMS);
        })/*
        .then(function () {
            return CIB.DE.installer.enableQuickLaunch(list.name);
        })*/
        .done(function () {
            CIB.DE.installer.message(list.name + ' created');
        })
        .fail(function (message) {
            CIB.DE.installer.message('Error creating list ' + list.name  + ': ' + message, 'error');
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

    var createArchiveList = function () {
        return createCustomList(siteCT.archive, siteLists.archive);
    }

    var setFeeFieldValidationFormula = function () {
        return CIB.DE.installer.setFieldValidationFormula(siteLists.fee.name, siteColumns.feeAccount.name, siteColumns.feeAccount.validationFormula, siteColumns.feeAccount.validationMessage);
    }

    var setArchiveFieldValidationFormula = function () {
        return CIB.DE.installer.setFieldValidationFormula(siteLists.archive.name, siteColumns.feeAccount.name, siteColumns.feeAccount.validationFormula, siteColumns.feeAccount.validationMessage);
    }

    var createWorkflowHistoryList = function () {
        return CIB.DE.installer.createLists({ name: 'Workflow History', type: 140, hidden: true })
        .then(function () {
            return CIB.DE.installer.updateListIds();
        })
        .done(function () {
            CIB.DE.installer.message('Workflow history list created.');
        })
        .fail(function (message) {
            CIB.DE.installer.message('Error in creating the workflow history list: ' + message, 'error');
        });
    };

    // create tasks list
    var createTasksList = function () {
        return CIB.DE.installer.createLists({ name: 'Tasks', type: 107 })
        .then(function () {
            return CIB.DE.installer.updateListIds();
        })
        .fail(function (message) {
            CIB.DE.installer.message('Error in creating the workflow tasks list: ' + message, 'error');
        })
        .done(function () {
            CIB.DE.installer.message('Workflow tasks list created.');
        });
    };

    ////////////////////////////////////////////////////
    // Install workflow
    ////////////////////////////////////////////////////
    var createWorkflow = function () {
		return CIB.DE.installer.updateListIds()
			.then(function () {
                return CIB.DE.installer.installWorkflowFromXaml(siteLists.fee.name, siteLists.fee.workflows.feeUpdated)
            })        
            .done(function () {
                CIB.DE.installer.message('Workflow installed successfully.');
            })
            .fail(function (message) {
                CIB.DE.installer.message('Error installing the workflow', message);
            });
    };

    ////////////////////////////////////////////////////
    // Provision Data
    ////////////////////////////////////////////////////
    var provisionListData = function (listDef) {
        var dfd = $.Deferred();
        var list = ctx.host.get_web().get_lists().getByTitle(listDef.name);
        ctx.context.load(list);
        return ctx.context.executeQueryAsyncPromise()
          .done(function () {
              $.each(listDef.items.values, function () {
                  var val = this;
                  var itc = new SP.ListItemCreationInformation();
                  var li = list.addItem(itc);
                  $.each(val, function (i, f) {
                      li.set_item(listDef.items.columns[i], f);
                  });
                  li.update();
                  ctx.context.load(li);
              });

              return ctx.context.executeQueryAsyncPromise()
                  .done(function () {
                      CIB.DE.installer.message('Item added to the list');
                      dfd.resolve();
                  })
                  .fail(function (message) {
                      CIB.DE.installer.message('Error adding entry to list ' + list.name + ': ' + message, 'error');
                      dfd.reject();
                  });              
          })
          .fail(function (message) {
              CIB.DE.installer.message('Error loading list ' + list.name + ': ' + message, 'error');
              dfd.reject();
          });

        return dfd.promise();
    };

    var provisionData = function () {
        return provisionListData(siteLists.config)
        .then(function () {
            return provisionListData(siteLists.currency);
        })
        .then(function () {
            return provisionListData(siteLists.team);
        })
        .then(function () {
            return provisionListData(siteLists.priceList);
        })
		.fail(function (message) {
			CIB.DE.installer.message('Error in provision data: ' + message, 'error');
		});
    };

    function addExportScriptWebPart() {
        return CIB.DE.installer.addWebPartsToPage([{
            url: 'Lists/Fees/Export%20View.aspx',
            title: 'Export Buttons',
            assembly: 'Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c',
            type: 'Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart',
            zone: 'Main',
            index: 2,
			properties: '<property name="Content" type="string">' + 
                '&lt;script type="text/javascript" src="../../SiteAssets/FeeMan/App.js"&gt;&lt;\/script&gt;' +
				'&lt;div&gt;' + 
				'&lt;button id="export-xls" type="button" class="btn btn-success" onclick="saveExcel();"&gt;Export to Excel&lt;/button&gt;' + 
				'     ' + // '&nbsp;&nbsp;&nbsp;' + 
				'&lt;button id="confrm-export" type="button" class="btn btn-success" onclick="alert(\'Confirm\');"&gt;Confirm export&lt;/button&gt;' + 
				'&lt;/div&gt;' +
				'</property>'
        }])
        .done(function () {
            CIB.DE.installer.message('Export Script WebPart added to page.', 'success')
        })
        .fail(function (message) {
            CIB.DE.installer.message('Error adding Export Script WebPart:' + message, 'error');
            CIB.DE.logging.logError('error', message);
        });
    }

    ////////////////////////////////////////////////////
    // Context
    ////////////////////////////////////////////////////
    var loadContext = function () {
		return CIB.DE.installer.refreshContext()
		.then(function () {
			// ctx = CIB.DE.utilities.getContext();
			CIB.DE.utilities.getContext().then(function(result) { 
				ctx = result;
			});
        })
		.fail(function (message) {
			CIB.DE.installer.message('Error in loading context: ' + message, 'error');
		});
    };

    /////////////////////////////////////////
    // Retrive subsites to install app
    /////////////////////////////////////////
    var retrieveSubsites = function () {
		var retrieveSubs = function (web, addWeb) {
			var context = web.get_context();
			if (addWeb)
				context.load(web);
			var subSites = web.get_webs();
			context.load(subSites);
			return context.executeQueryAsyncPromise().done(function () {
				var subs = subSites.getEnumerator();
				if (addWeb) {
					var webTitle = web.get_title();
					var webUrl = web.get_serverRelativeUrl();
					$('#drp_subsite').append($('<option>', { value: webUrl }).text(webTitle + ' (' + webUrl + ')'));
				}
				while (subs.moveNext()) {
					var subSite = subs.get_current();
					webTitle = subSite.get_title();
					if (webTitle && subSite.get_webTemplate() !== 'APP') {
						webUrl = subSite.get_serverRelativeUrl();
						$('#drp_subsite').append($('<option>', { value: webUrl }).text(webTitle + ' (' + webUrl + ')'));
						retrieveSubs(subSite, false);
					} 
				}
			 })
			 .fail(function (message) {
				 CIB.DE.installer.message('Error in retrieving subsites: ' + message, 'error');
			 });
		};

		var rootWeb = ctx.host.get_site().get_rootWeb();
		retrieveSubs(rootWeb, true);
    };
	
    $(document).ready(function () {

        CIB.DE.utilities.getContext().then(function(result) { 
			ctx = result;
			// Retrives all the sub sites
			retrieveSubsites();
		});

        $("#drp_subsite").change(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                document.getElementById("validate-msg").style.visibility = "hidden";
                $('#CIBAppFrameWorkSubWebUrl').html($("#drp_subsite option:selected").val());
            }
//          loadContext();
        });

        $('#copy-files').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    copyFiles)
            .done(function () { CIB.DE.installer.message('Files are copied.', 'success'); })
            .fail(function (message) { CIB.DE.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-lists').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    createConfigList,
                    createCurrenciesList,
                    createTeamsList,
                    createPriceList,
                    createFeesList,
                    setFeeFieldValidationFormula,
                    createArchiveList,
                    setArchiveFieldValidationFormula,
                    createWorkflowHistoryList,
                    createTasksList)
            .done(function () { CIB.DE.installer.message('lists are installed.', 'success'); })
            .fail(function (message) { CIB.DE.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-webparts').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    addExportScriptWebPart)
            .done(function () { CIB.DE.installer.message('webparts are added.', 'success'); })

            .fail(function (message) { CIB.DE.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });
		
        $('#install-workflows').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    createWorkflow)
            .done(function () { CIB.DE.installer.message('workflow is installed.', 'success'); })
            .fail(function (message) { CIB.DE.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-data').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    provisionData)
            .done(function () { CIB.DE.installer.message('data is provisioned.', 'success'); })

            .fail(function (message) { CIB.DE.installer.message(message, 'error'); });
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
    FeeMan App Installation Page
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
        <button id="copy-files" type="button" class="btn btn-success" data-loading-text="Copy Files">1. Copy Files</button>
        <br />
        <br />
        <button id="install-lists" type="button" class="btn btn-success" data-loading-text="Install Lists">2. Install Lists</button>
        <br />
        <br />
        <button id="install-webparts" type="button" class="btn btn-success" data-loading-text="Adding webparts...">3. Add Webparts</button>
        <br />
        <br />
        <button id="install-workflows" type="button" class="btn btn-success" data-loading-text="Installing workflows...">4. Install Workflows</button>
        <br />
        <br />
        <button id="install-data" type="button" class="btn btn-success" data-loading-text="Installing data...">5. Provision data</button>
        <br />
        <div id="install-status">
        </div>
        <div id="CIBAppFrameWorkSubWebUrl" style="display: none"></div>
    </div>
</asp:Content>
