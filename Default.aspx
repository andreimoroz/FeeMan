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
    var testAdd = '9';
    ////// ************ //////////////////////////
    var groupName = 'CIB DE FeeMan';
    var siteColumns = {
        currencyCode: {
            name: 'fmCode' + testAdd,
            id: '{3d9abb16-d9f6-4381-8e30-db1812e6c6e' + testAdd + '}',
            displayName: 'Code' + testAdd,
            type: 'Text',
            maxLength: '3',
            required: true,
            enforceUnique: true,
            linkToItem: 'Required',
            group: groupName
        },
        teamName: {
            name: 'fmName' + testAdd,
            id: '{62cb9781-a886-4641-ac7d-cc3d7b13239' + testAdd + '}',
            displayName: 'Name' + testAdd,
            type: 'Text',
            maxLength: '20',
            required: true,
            enforceUnique: true,
            linkToItem: 'Required',
            group: groupName
        },
        teamMembers: {
            name: 'fmMembers' + testAdd,
            id: '{8618d2f1-e739-4bc0-a2c4-1dfb1557880' + testAdd + '}',
            displayName: 'Members' + testAdd,
            type: 'UserMulti',
            multi: true,
            userSelectionMode: 'PeopleOnly',
            linkToItem: 'Required',
            group: groupName
        }
        /*
      <Field ID="{62cb9781-a886-4641-ac7d-cc3d7b132399}" Name="Name" DisplayName="Name" Type="Text" MaxLength="20" Required="TRUE" Group="Feeman Columns"
             AllowDuplicateValues="FALSE" EnforceUniqueValues="TRUE" Indexed="TRUE" LinkToItem="TRUE" LinkToItemAllowed="Required" ListItemMenu="TRUE" ></Field>
      <Field ID="{8618d2f1-e739-4bc0-a2c4-1dfb1557880e}" Name="Members" DisplayName="Members" Type="UserMulti" Mult="TRUE" List="UserInfo"
             ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0" Required="FALSE" Group="Feeman Columns" />
*/

    };
    var siteColumnArray = [
        siteColumns.currencyCode, siteColumns.teamName, siteColumns.teamMembers
    ];

    var siteCT = {
        currency: {
            id: '0x0100F34E70C088364B1B85B9462BA830A28' + testAdd,
            name: 'CIB Feem Currencies' + testAdd,
            group: groupName,
            columns: [siteColumns.currencyCode],
            columnNames: [siteColumns.currencyCode.name]
        },
        team: {
            id: '0x01007DDEF9578DB3451495FF451DF3B9539' + testAdd,
            name: 'CIB Feem Teams' + testAdd,
            group: groupName,
            columns: [siteColumns.teamName, siteColumns.teamMembers],
            columnNames: [siteColumns.teamName.name, siteColumns.teamMembers.name]
        }
                           
    };

    var siteLists = {
        currency: {
            name: 'Currencies' + testAdd, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.currency.id],
            views: {
                default: {
                    name: 'Currencies' + testAdd, columns: [siteColumns.currencyCode.name], query: '<OrderBy><FieldRef Name="' + siteColumns.currencyCode.name + '" Ascending="TRUE"></FieldRef></OrderBy>'
                }
            }
        },
        team: {
            name: 'Teams' + testAdd, 
            type: SP.ListTemplateType.genericList,
            onQuickLaunch: true,
            contentTypes: [siteCT.team.id],
            views: {
                default: {
                    name: 'Teams' + testAdd, columns: [siteColumns.teamName.name], query: '<OrderBy><FieldRef Name="' + siteColumns.teamName.name + '" Ascending="TRUE"></FieldRef></OrderBy>'
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
//                     subTitlesUrl = helper.getServerRelativeUrl(subSite.get_url());
                    subTitlesUrl = subSite.get_url();
                    $('#drp_subsite').append($('<option>', { value: subTitlesUrl }).text(subTitles + ' (' + subTitlesUrl + ')'));
                } 
            }
         })
         .fail(function () {
             CIB.installer.message('Error in retrieveing subsites', 'error');
         });
    };

    var helper = function () {
        return {

            getBoolean: function(s) {
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
                        if (handled) {
                            if (value) { promise.resolve(messages, value); } else { promise.resolve(messages); }
                            
                        }
                        else {
                            promise.reject(messages);
                        }
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

                return { handled: type == 'info', message: message }
            },

            /*
                Writes colour coded messages to the user, this method assumes the existence of an element with id 'install-status'
            */
            message: function (text, type) {
                if (!type) type = 'message';
                if (console && console.log) console.log(text + ' [' + type + ']');
                // --------------------------------------------------------
                // TODO: remove the comment when logging app is deployed to IHC
                //if (type == 'error') CIB.logging.logError('Provisioning', text, window.location.href);
                var colour = type == 'success' ? 'green' : (type == 'error' ? 'red' : (type == 'info' ? 'orange' : 'gray'));
                $('#install-status').append('<span style="color:' + colour + '">' + text + '</span>');
                var elem = document.getElementById('install-status');
                if (elem) elem.scrollTop = elem.scrollHeight;
            },

            arrayBufferToBase64: function (buffer) {
                var binary = '';
                var bytes = new Uint8Array(buffer);
                //var bytes = new SP.Base64EncodedByteArray(buffer)
                var len = bytes.byteLength;
                for (var i = 0; i < len; i++) {
                    binary += String.fromCharCode(bytes[i]);
                }
                return window.btoa(binary);
                //return binary;
            },
            /*
                Get all list ids to for use by lookup columns
            */
            updateListIds: function () {
                var listIdsUpdated = new jQuery.Deferred();
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
                return listIdsUpdated.promise();
            }

            /*
                function sleepFor(sleepDuration) {
                    var now = new Date().getTime();
                    while (new Date().getTime() < now + sleepDuration) {  }
                }
            */
        }
    }()

    var installer = function () {
        return {
            createColumns: function (columns) {
                var executepromise = $.Deferred();
                columns = CIB.utilities.ensureArray(columns);
                var fields = hostContext.get_web().get_fields();
                context.load(fields);
                context.executeQueryAsyncPromise().done(function () {
                    $.each(columns, function () {
                        var column = this;
                        if (!column.id || !column.name || !column.type || !column.displayName || !column.group) throw new Error("Column object must have id, name, type, group and displayName attributes");
            
                        CIB.installer.message("Creating column '" + column.displayName + "'");
                        var hidden = helper.getBoolean(column.hidden) ? 'true' : 'false';
                        var required = helper.getBoolean(column.required) ? 'true' : 'false';
                        var multi = helper.getBoolean(column.multi) ? 'true' : 'false';
                        var xml = "<Field ID='" + column.id + "' Type='" + column.type + "' DisplayName='" + column.name + "' Name='" + column.name + "' Group='" + column.group + "' Required='" + required + "' />";
                        if (column.type.toLowerCase() == "usermulti") {
                            multi = 'true';
                            xml = xml.replace(" />", " List='UserInfo' ShowField='ImnName' " + 
                                (column.hasOwnProperty('userSelectionMode') ? ("UserSelectionMode='" + column.userSelectionMode + "'") : '') + 
                                " UserSelectionScope='0' />"); // TODO
                        }
                        if (column.maxLength) {
                            xml = xml.replace(" />", " MaxLength='" + column.maxLength + "' />");
                        }
                        if (column.hasOwnProperty('enforceUnique')) {
                            xml = xml.replace(" />", " AllowDuplicateValues='" + (!helper.getBoolean(column.enforceUnique)).toString().toUpperCase() + "' EnforceUniqueValues='" + helper.getBoolean(column.enforceUnique).toString().toUpperCase() + "' />");
                        }
                        if (column.linkToItem) {
                            xml = xml.replace(" />", " LinkToItem='TRUE' LinkToItemAllowed='" + column.linkToItem + "' ListItemMenu='TRUE' />");
                        }
                        if (column.type.toLowerCase() == "calculated") {
                            if (!column.formula || !column.resultType) throw new Error("Calculated columns must have a formula and resultType set");
                            var formula = "<Formula>" + column.formula + "<\/Formula>";
                            xml = xml.replace(" />", ' ResultType="' + column.resultType + '">' + formula + "<\/Field>")
                        }
                        if (multi == 'true') { xml = xml.replace(" />", ' Mult="TRUE" />') };
                        var field = fields.addFieldAsXml(xml, false, SP.AddFieldOptions.AddToNoContentType);
                        if (hidden == 'true') { field.set_hidden(hidden) }
                        field.set_title(column.displayName);
                        field.set_required(required);
                        if (column.defaultValue) { field.set_defaultValue(column.defaultValue) }
                        context.load(field);
                        if (column.type.toLowerCase() == "lookup") {
                            var lookupField = context.castTo(field, SP.FieldLookup);
                            // TODO
                            lookupField.set_lookupList(r[column.lookupList]);
                            lookupField.set_lookupField(column.lookupField);
                            lookupField.update();
                            if (column.additionalFields) {
                                $.each(column.additionalFields, function () {
                                    var f = this;
                                    fields.addDependentLookup(f.displayName, field, f.target)
                                })
                            }
                        }
                        else if (column.type.toLowerCase() == "lookupmulti") {
                            var lookupField = context.castTo(field, SP.FieldLookup);
                            // TODO
                            lookupField.set_lookupList(r[column.lookupList]);
                            lookupField.set_lookupField(column.lookupField);
                            lookupField.set_allowMultipleValues(true);
                            lookupField.update();
                            if (column.additionalFields) {
                                $.each(column.additionalFields, function () {
                                    var n = this;
                                    fields.addDependentLookup(f.displayName, field, f.target)
                                })
                            }
                        }
                        else if (column.type.toLowerCase() == "currency" && column.locale) {
                            var currencyField = context.castTo(field, SP.FieldCurrency);
                            currencyField.set_currencyLocaleId(column.locale);
                            currencyField.update();
                        }
                        else if (column.type.toLowerCase() == "number") {
                            var numberField = context.castTo(field, SP.FieldNumber);
                            if (column.minimumValue) numberField.set_minimumValue(column.minimumValue);
                            if (column.maximumValue) numberField.set_maximumValue(column.maximumValue);
                            numberField.update();
                        }
                        else if (column.type.toLowerCase() == "choice" && column.choices) {
                            var choiceField = context.castTo(field, SP.FieldChoice);
                            choiceField.set_choices($.makeArray(column.choices));
                            choiceField.update()
                        }
                        else if (column.type.toLowerCase() == "multichoice" && column.choices) {
                            var choiceField = context.castTo(field, SP.FieldMultiChoice)
                            choiceField.set_choices($.makeArray(column.choices));
                            choiceField.update();
                        }
                        else if (column.type.toLowerCase() == "datetime" && column.dateOnly) {
                            var dateField = context.castTo(field, SP.FieldDateTime);
                            dateField.set_displayFormat(SP.DateTimeFieldFormatType.dateOnly);
                            dateField.update();
                        }
                        else if (column.type.toLowerCase() == "taxonomyfieldtypemulti") {
                            var taxField = context.castTo(field, SP.Taxonomy.TaxonomyField);
                            taxField.set_allowMultipleValues(true);
                            taxField.update();
                        }
                        else {
                            field.update()
                        }
            
                        context.executeQueryAsyncPromise()
                            .done(function (message) {
                                CIB.installer.message('Column ' + column.displayName + ' created');
                            })
                            .fail(function (message) {
                                if (message.match('duplicate')) {
                                    CIB.installer.message(message + ' (expected if provisioned already)', 'info');
                                }
                                else {
                                    CIB.installer.message('Error adding column ' + column.displayName + ': ' + message, 'error');
                                }
                            });
                    });

                    executepromise.resolve();
                })
                .fail(function (message) {
                    executepromise.reject();
                    CIB.installer.message('Error adding columns: ' + message, 'error');
                });
                return executepromise.promise();
            },

            hideContentTypeField: function (contentTypeId, fieldName) {
                var executepromise = $.Deferred();
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
                                            field.set_hidden(true);
                                            field.set_required(false);
                                            content.update();
                                            context.executeQueryAsyncPromise().done(function () {
                                                CIB.installer.message('Field: ' + fieldName + ' is hided.');
                                            })
                                            .fail(function (message) {
                                                CIB.installer.message('Error hide field: ' + fieldName + ': ' + message, 'error');
                                            });
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

            addContentTypesToList: function (listName, contentTypes) {
                var executepromise = $.Deferred();
                CIB.installer.message("Adding content types to list '" + listName + "'");
                var contentTypes = CIB.utilities.ensureArray(contentTypes);

                var web = hostContext.get_web();
                var list = web.get_lists().getByTitle(listName);
                list.set_contentTypesEnabled(true);
                list.update();
                var availableContentTypes = web.get_availableContentTypes();
                var listContentTypes = list.get_contentTypes();
                context.load(listContentTypes);
                context.executeQueryAsyncPromise().done(function () {
                    var scopes = [];
                    $.each(contentTypes, function () {
                        var contentType = this;
                        var scope = $.handleExceptionsScope(context, function () {
                            var ct = availableContentTypes.getById(contentType);
                            listContentTypes.addExistingContentType(ct);
                        });
                        scope.successMessage = "Content type " + contentType + " added to list";
                        scopes.push(scope);
                    });
                    helper.executeQuery(scopes, executepromise);
                })
                .fail(function (message) {
                    executepromise.reject();
                    CIB.installer.message('Error adding content types to list: ' + message, 'error');
                });
                return executepromise.promise();
            },

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
        return installer.createColumns(contentType.columns)
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
            return helper.updateListIds();
        })
        .then(function () {
            return installer.setFieldVisibility(list.name, TITLE);
        })
        .then(function () {
            return installer.addContentTypesToList(list.name, list.contentTypes);
        })
        .then(function () {
            return installer.removeContentTypesFromList(list.name, ITEM);
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

    var createCurrenciesList = function () {
        return createCustomList(siteCT.currency, siteLists.currency);
    }

    var createTeamsList = function () {
        return createCustomList(siteCT.team, siteLists.team);
    }

/*
    var createCustomList = function (contentType, list) {
        var executepromise = $.Deferred();
        // currency
        return installer.createColumns(siteCT.currency.columns)
        .then(function () {
            return CIB.installer.createContentTypes(siteCT.currency);
        })
        .then(function () {
            return CIB.installer.addColumnsToContentType(siteCT.currency.id, siteCT.currency.columnNames);
        })
        .then(function () {
            return installer.hideContentTypeField(siteCT.currency.id, TITLE);
        })
        .then(function () {
            return CIB.installer.updateListIds();
        })
        .then(function () {
            return installer.setFieldVisibility(siteLists.currency.name, TITLE);
        })
        .then(function () {
            return installer.addContentTypesToList(siteLists.currency.name, siteLists.currency.contentTypes);
        })
        .then(function () {
            return installer.removeContentTypesFromList(siteLists.currency.name, ITEM);
        })
        .then(function () {
            return CIB.installer.createView(siteLists.currency.name, siteLists.currency.views.default.name, siteLists.currency.views.default.columns, siteLists.currency.views.default.query);
        })
        .then(function () {
            return installer.removeView(siteLists.currency.name, ALL_ITEMS);
        })
        // team
        .done(function () {
            CIB.installer.message('currency content type created');
        })
        .fail(function (message) {
            CIB.installer.message('Error in creating the currency content type: ' + message, 'error');
        });
    };



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

    ////////////////////////////////////////////////////
    // Currencies Handlers
    ////////////////////////////////////////////////////
    var currencyHandler = function () {
        return $.whenSync(
           function () {
               return CIB.installer.createLists(siteLists.currency)
                   .then(function () {
                       return CIB.installer.updateListIds();
                   })
                   .then(function () {
                       return installer.setFieldVisibility(siteLists.currency.name, TITLE);
                   })
                   .then(function () {
                       return installer.addContentTypesToList(siteLists.currency.name, siteLists.currency.contentTypes);
                   })
                   .then(function () {
                       return installer.removeContentTypesFromList(siteLists.currency.name, ITEM);
                   })
                   .then(function () {
                       return CIB.installer.createView(siteLists.currency.name, siteLists.currency.views.default.name, siteLists.currency.views.default.columns, siteLists.currency.views.default.query);
                   })
                   .then(function () {
                       return installer.removeView(siteLists.currency.name, ALL_ITEMS);
                   })
                   .fail(function (message) {
                       CIB.installer.message('Error creating ' + siteLists.currency.name + ' list: ' + message, 'error');
                       CIB.logging.logError('error', message);
                   });
           });
    };



    var currencyCreateList = function () {
        return CIB.installer.createLists(siteLists.currency)
            .then(function () {
                CIB.installer.updateListIds();
            })
            .then(function () {
                return installer.setFieldVisibility(siteLists.currency.name, TITLE);
            })
            .done(function () {
                CIB.installer.message(siteLists.currency.name + ' list created');
            })
            .fail(function (message) {
                CIB.installer.message('Error creating ' + siteLists.currency.name + ' list: ' + message, 'error');
                CIB.logging.logError('error', message);
            });
    };

    var currencyContentTypesAdd = function () {
        return installer.addContentTypesToList(siteLists.currency.name, siteLists.currency.contentTypes)
            .done(function () {
                CIB.installer.message(siteLists.currency.name + ': content type added');
            })
            .fail(function (message) {
                CIB.installer.message(siteLists.currency.name  + ': error adding content type: ' + message, 'error');
                CIB.logging.logError('error', message);
            });
    };

    var currencyContentTypesRemove = function () {
        return installer.removeContentTypesFromList(siteLists.currency.name, ITEM)
            .done(function () {
                CIB.installer.message(siteLists.currency.name + ': ' + ITEM + ' content type removed');
            })
            .fail(function (message) {
                CIB.installer.message(siteLists.currency.name + ': error removing content type: ' + message, 'error');
                CIB.logging.logError('error', message);
            });
    };

    var currencyViewAdd = function () {
        return CIB.installer.createView(siteLists.currency.name, siteLists.currency.views.default.name, siteLists.currency.views.default.columns, siteLists.currency.views.default.query)
            .done(function () {
                CIB.installer.message(siteLists.currency.name + ': view added');
            })
            .fail(function (message) {
                CIB.installer.message(siteLists.currency.name + ': error creating view: ' + message, 'error');
                CIB.logging.logError('error', message);
            });
    };

    var currencyViewRemove = function () {
        return installer.removeView(siteLists.currency.name, ALL_ITEMS)
            .done(function () {
                CIB.installer.message(siteLists.currency.name + ': view removed');
            })
            .fail(function (message) {
                CIB.installer.message(siteLists.currency.name + ': error removing view: ' + message, 'error');
                CIB.logging.logError('error', message);
            });
    };

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

/*
    var refreshContext2 = function () {
    }
*/


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
                    createCurrenciesList,
                    createTeamsList)
            .done(function () { CIB.installer.message('lists are installed.', 'success'); })
            .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-contenttypes').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    createContentTypes)
            .done(function () { CIB.installer.message('Content types are installed.', 'success'); })
            .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-currency').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    currencyHandler)
                .done(function () { CIB.installer.message('Done.', 'success'); })
                .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-currency-create').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    currencyCreateList)
                .done(function () { CIB.installer.message('Done.', 'success'); })
                .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-currency-contenttype-add').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    currencyContentTypesAdd)
                .done(function () { CIB.installer.message('Done.', 'success'); })
                .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-currency-contenttype-remove').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    currencyContentTypesRemove)
                .done(function () { CIB.installer.message('Done.', 'success'); })
                .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-currency-view-add').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                    $.whenSync(
                    loadContext,
                    currencyViewAdd)
                .done(function () { CIB.installer.message('Done.', 'success'); })
                .fail(function (message) { CIB.installer.message(message, 'error'); });
            }
            else {
                $('#validate-msg').html('Please select site to install an app').css({ 'color': 'red', 'font-size': '100%' });
                document.getElementById("validate-msg").style.visibility = "visible";
            }
        });

        $('#install-currency-view-remove').click(function () {
            if (document.getElementById("drp_subsite").selectedIndex != 0) {
                $.whenSync(
                    loadContext,
                    currencyViewRemove)
                .done(function () { CIB.installer.message('Done.', 'success'); })
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
        <button id="install-lists" type="button" class="btn btn-success" data-loading-text="Install Content Types">1. Install Lists</button>
        <br />
        <br />
        <button id="install-contenttypes" type="button" class="btn btn-success" data-loading-text="Install Content Types">1. Install Content Types</button>
        <br />
        <br />
        <table><tbody>
            <tr>
                <td>Currencies</td>
                <td><button id="install-currency" type="button" class="btn btn-success" data-loading-text="Create currencies">List Handler</button></td>
                <td><button id="install-currency-create" type="button" class="btn btn-success" data-loading-text="Create currencies">Create List</button></td>
                <td><button id="install-currency-contenttype-add" type="button" class="btn btn-success" data-loading-text="Set Content Type">Add Content Type</button></td>
                <td><button id="install-currency-contenttype-remove" type="button" class="btn btn-success" data-loading-text="Remove Item Content Type">Remove Item Content Type</button></td>
                <td><button id="install-currency-view-add" type="button" class="btn btn-success" data-loading-text="Create View">Create View</button></td>
                <td><button id="install-currency-view-remove" type="button" class="btn btn-success" data-loading-text="Remove View">Remove Default View</button></td>
            </tr>
        </tbody></table>
        
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
