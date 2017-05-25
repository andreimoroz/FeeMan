'use strict';

/*

    Installer.js
    Provides a framework to provision SharePoint artefacts on host web

*/

var CIB = CIB || {};

CIB.installer = function () {

    var globalContext;
    var context;
    var hostContext;

    var listIds = {};

    // This multidimentional array will be storing list name - list view - view ID data
    var listToViewIds = {};

    $(document).ready(function () {
        //get context after the document is ready; to get the reference of "CIBSubWebURL" element in "Utilities.js".
        globalContext = CIB.utilities.getContext()
        context = globalContext.context;
        hostContext = globalContext.hostContext;

        if (!$.isInternetExplorer() && !$.hasAppWeb()) {
            if ($('#install-status').length > 0) {
                $('#install-status').append(
                    $('<div class="alert alert-danger" role="alert" style="width:580px;">' +
                        '<strong>Unsupported browser</strong>' +
                        '<span>The provisioning wizard will only work with internet explorer for provider hosted apps.</span>' +
                    '</div>'));
            }
            else {
                throw new Error('Unsupported browser: The provisioning wizard will only work with internet explorer for provider hosted apps.');
            }
        }
    });

    /*
        Helper namespace contains utility functions for the installer
    */
    var helper = function () {
        return {
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

    window.onerror = function (errorMsg, url, lineNumber) {
        CIB.logging.logError('Unhandled JavaScript Error', errorMsg, 'Line: ' + lineNumber + '\r\n' + url);
        helper.message(errorMsg, 'error');
    };
    return {

        message: function (text, type) {
            helper.message(text, type);
        },

        getListIds: function () {
            var getListIds = new jQuery.Deferred();

            helper.updateListIds()
            .done(function () {
                getListIds.resolve(listIds);
            })
            .fail(function (message) {
                getListIds.reject(message);
            });

            return getListIds.promise();
        },

        /*
            Activates site or web features
        */
        activateFeatures: function (features) {
            var scopes = [];
            var features = CIB.utilities.ensureArray(features);

            var featuresActivated = new jQuery.Deferred();

            $.each(features, function () {
                var feature = this;

                if (!feature.id || !feature.name || !feature.scope)
                    throw new Error('Feature object must had id, name and scope attributes');

                if (feature.scope != 'site' && feature.scope != 'web')
                    throw new Error('Feature scope must be either site or web');

                var scope = $.handleExceptionsScope(context, function () {
                    helper.message('Activateg feature \'' + feature.name + '\'');

                    var activatedFeatures = feature.scope == 'site' ? hostContext.get_site().get_features() : hostContext.get_web().get_features();
                    var featureDefinition = activatedFeatures.add(new SP.Guid(feature.id), false, SP.FeatureDefinitionScope.farm);
                });

                scope.successMessage = 'Feature \'' + feature.name + '\' activated.';
                scopes.push(scope);
            });

            helper.executeQuery(scopes, featuresActivated);

            return featuresActivated.promise();
        },

        /*
            Create lists on the host web
            @lists { name: 'Example', type: 100 }
        */
        createLists: function (lists) {
            var scopes = [];
            var lists = CIB.utilities.ensureArray(lists);

            var listsCreated = new jQuery.Deferred();

            $.each(lists, function () {
                var list = this;

                if (!list.name || !list.type)
                    throw new Error('List object must have name and type attributes');

                var scope = $.handleExceptionsScope(context, function () {
                    helper.message('Creating list \'' + list.name + '\'');

                    var lists = hostContext.get_web().get_lists();

                    var newList = new SP.ListCreationInformation();
                    newList.set_title(list.name);
                    newList.set_templateType(list.type);

                    if (list.feature) {
                        newList.set_templateFeatureId(list.feature);
                    }

                    var updateNeeded = false;
                    var newList = lists.add(newList);
                    /*check for CIB document library template */
                    if (list.type == "10002" || list.type == "10000" || list.type == "10001") {
                        var listRootFolder = newList.get_rootFolder();
                        var rootFolderProperties = listRootFolder.get_properties();
                        if (list.type == "10002") //Public
                            rootFolderProperties.set_item('InformationSecurityLevel', 0);
                        else if (list.type == "10001") //Confidential
                            rootFolderProperties.set_item('InformationSecurityLevel', 2);
                        else if(list.type == "10000") //Restricted
                            rootFolderProperties.set_item('InformationSecurityLevel', 1);
                        listRootFolder.update();
                    }

                    if (list.hasOwnProperty('hidden')) {
                        newList.set_hidden(list.hidden);
                        updateNeeded = true;
                    }

                    if(list.hasOwnProperty('onQuickLaunch')) {
                        newList.set_onQuickLaunch(list.onQuickLaunch);
                        updateNeeded = true;
                    }

                    if(updateNeeded)
                        newList.update();
                });
                scope.successMessage = 'List \'' + list.name + '\' created.';
                scopes.push(scope);
            });

            helper.executeQuery(scopes, listsCreated);

            return listsCreated.promise();

        },

        /*
            Create folders in pre-existing lists on the host web
            @folders { name: 'CIB', list: 'Style Library', path: 'Style Library' }
        */
        createFolders: function (folders) {
            var scopes = [];
            var folders = CIB.utilities.ensureArray(folders);

            var foldersCreated = new jQuery.Deferred();

            $.each(folders, function () {
                var folder = this;

                if (!folder.name || !folder.list || !folder.path)
                    throw new Error('Folder object must have name, list and path attributes');

                var scope = $.handleExceptionsScope(context, function () {
                    helper.message('Creating folder ' + folder.name + ' in list ' + folder.list);

                    var list = hostContext.get_web().get_lists().getByTitle(folder.list);

                    var folderInfo = new SP.ListItemCreationInformation();
                    folderInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
                    folderInfo.set_leafName(folder.name);
                    folderInfo.set_folderUrl($.getHostWebUrl() + '/' + folder.path);
                    var folderItem = list.addItem(folderInfo);
                    folderItem.set_item('Title', folder.name);
                    folderItem.update();

                });
                scope.successMessage = 'Folder ' + folder.name + ' created in list ' + folder.list;
                scopes.push(scope);
            });

            helper.executeQuery(scopes, foldersCreated);

            return foldersCreated.promise();
        },

        // Token should be in the following format: {\$([^:]*):([^}]*)} - {$List:Workflow History}
        updateFileTokens: function (content){

            // Declare helper functions
            var getListGuid = function (listIds, listName) {
                if (listIds[listName])
                {
                    return listIds[listName];
                }
            };

            var getViewGuid = function (listName, viewName) {
                //	Example: listToViewIds['Alert Configuration List']['All Items']
                if (listToViewIds[listName])
                {
                    return listToViewIds[listName][viewName];
                }
            };

            var processContent = function (content) {

                var tokens = new RegExp('\{\\$([^\:]*)\:([^\}]*)\}', 'gi');
                content = content.replace(tokens, function (x, tokenType, tokenName) {
                    if (tokenType === 'List')
                    {
                        return getListGuid(listIds, tokenName);
                    }
                    return x;
                });

                // Replace view token like: {\$([^:]*):([^}]*):([^}]*)} - {$ListView:Stakeholder:All Items}
                tokens = new RegExp('\{\\$([^\:]*)\:([^\}]*)\:([^\}]*)\}', 'gi');
                content = content.replace(tokens, function (x, tokenType, listToken, viewToken) {
                    if (tokenType === 'ListView')
                    {
                        return getViewGuid(listToken, viewToken);
                    }
                    return x;
                });


                return content;
            };

            content = processContent(content);
            return content;
        },

        /*
            Copies html, js, css and other text based content to the host web
            Due to browser limitations binary files are not currently supported, see inline comments for more details
            @files { name: 'App.css', url: 'Style Library/CIB/CSS/Common', path: 'Style Library/CIB/CSS/Common/App.css' }
        */
        copyFiles: function (files) {

            var files = CIB.utilities.ensureArray(files);

            var counter = 0;
            var filesCopied = new jQuery.Deferred();

            $.each(files, function () {
                var scopes = [];
                var file = this;

                if (!file.name || !file.url || !file.path)
                    throw new Error('File object must have name, url and path attributes');

                helper.message('Copying file ' + file.name);

                var getFileUsingRequestExecutor = function (file) {
                    var binary = file.binary ? true : false;
                    // https://msdn.microsoft.com/en-us/library/office/dn450841.aspx
                    // If you're not making cross-domain requests, remove SP.AppContextSite(@target) and ?@target='<host web url>' from the endpoint URI.
                    // var fileContentUrl = "_api/SP.AppContextSite(@target)/web/GetFileByServerRelativeUrl('" + ($.getServerRealtiveApptWebUrl() + "/" + file.path).replace('//', '/') + "')/$value?@target='" + $.getHostWebUrl() + "'";
                    var fileContentUrl = "_api/web/GetFileByServerRelativeUrl('" + ($.getServerRealtiveApptWebUrl() + "/" + file.path).replace('//', '/') + "')/$value";
                    var executor = new SP.RequestExecutor($.getAppWebUrl());
                    var info = {
                        url: fileContentUrl,
                        method: "GET",
                        binaryStringResponseBody: binary,
                        success: function (data) {
                            uploadFileToHostWeb(file, data.body);
                        },
                        error: function (err) {
                            // Resort to using AJAX
                            getFileUsingAjax(file);
                        }
                    };
                    executor.executeAsync(info);
                };

                var getFileUsingAjax = function (file) {
                    $.support.cors = true;
                    var fileContentUrl = $.getAppWebUrl() + "/_api/SP.AppContextSite(@target)/web/GetFileByServerRelativeUrl('" + ($.getServerRealtiveApptWebUrl() + "/" + file.path).replace('//', '/') + "')/$value?@target='" + $.getHostWebUrl() + "'";
                    $.ajax({ url: fileContentUrl, cache: false })
                        .done(function (content) {
                            uploadFileToHostWeb(file, content);
                        })
                        .fail(function (sender, status) {
                            filesCopied.reject(sender.statusText);
                        });
                };

                var uploadFileToHostWeb = function (file, fileContents) {

                    // Token replacement should never happen for Installer.js itself. This may occur when installer installs itself
                    if (file.name != "Installer.js") {
                        // Do token replacement if necessary. Currently supported tokens are {#HostWebURL#} and {#ServerRelativeHostWebURL#}
                        fileContents = fileContents.replace(/{#HostWebURL#}/g, $.getHostWebUrl());
                        fileContents = fileContents.replace(/{#ServerRelativeHostWebURL#}/g, $.getServerRealtiveHostWebUrl());
                        fileContents = CIB.installer.updateFileTokens(fileContents);
                    }

                    var destinationUrl = ($.getServerRealtiveHostWebUrl() + '/' + file.url).replace('//', '/');

                    var scope = $.handleExceptionsScope(context, function () {

                        if (file.publish) {
                            // Attempt to checkout the document and ignore any errors.
                            // The file upload code will throw an error if something is unexpected
                            $.handleExceptionsScope(context, function () {
                                var existingFile = hostContext.get_web().getFileByServerRelativeUrl(destinationUrl + '/' + file.name);
                                existingFile.checkOut();
                            });
                        }

                        var createInfo = new SP.FileCreationInformation();
                        createInfo.set_content(new SP.Base64EncodedByteArray());
                        for (var i = 0; i < fileContents.length; i++) {
                            createInfo.get_content().append(fileContents.charCodeAt(i));
                        }
                        createInfo.set_overwrite(true);
                        createInfo.set_url(file.name);
                        var files = hostContext.get_web().getFolderByServerRelativeUrl(destinationUrl).get_files();
                        var newFile = files.add(createInfo);
                        if (file.publish) {
                            newFile.checkIn('Checked in by provisioning framework.', SP.CheckinType.majorCheckIn);
                            newFile.publish('Published by provisioning framework.');
                        }
                    });

                    scope.successMessage = 'File ' + file.name + ' created at ' + file.url;
                    scopes.push(scope);

                    if (++counter == files.length)
                        helper.executeQuery(scopes, filesCopied);
                };

                if ($.isAppWeb()) {
                    getFileUsingRequestExecutor(file);
                }
                else {
                    getFileUsingAjax(file);
                }

            });

            return filesCopied.promise();
        },

        /*
            Helper function to populate a mapping of lists ids to be used when creating lookup columns
        */
        updateListIds: function () {
            return helper.updateListIds();
        },

        updateViewIds: function () {
            return helper.updateViewIds();
        },

        /*
            Create a site column in the host web
            @columns { name: 'cmppYear', id: '{F4605722-C180-46B0-8AAE-0C0BC0EA4EC3}', displayName: 'Year', type: 'Number', group: 'Test' }
            Addtional parameters are supported for lookups, calculated, datetime and choice fields
        */
        createSiteColumns: function (columns) {
            var fields = hostContext.get_web().get_fields();
            return CIB.installer.createColumns(columns, fields);
        },

        /*
            Create a site column in the specified list
            @listTitle 'Documents'
            @columns { name: 'cmppYear', id: '{F4605722-C180-46B0-8AAE-0C0BC0EA4EC3}', displayName: 'Year', type: 'Number', group: 'Test' }
            Addtional parameters are supported for lookups, calculated, datetime and choice fields
        */
        createListColumns: function (listTitle, columns) {
            var list = hostContext.get_web().get_lists().getByTitle(listTitle);
            var fields = list.get_fields();
            return CIB.installer.createColumns(columns, fields);
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
                        helper.message('Creating column \'' + column.displayName + '\'');

                        var hidden = column.hidden ? 'true' : 'false';
                        var required = column.required ? 'true' : 'false';
                        var multi = column.multi ? 'true' : 'false';

                        var fieldXml = "<Field ID='" + column.id + "' Type='" + column.type + "' DisplayName='" + column.name +
                            "' Name='" + column.name + "' Group='" + column.group + "' Required='" + required + "' />";

                        if (column.type.toLowerCase() == 'calculated') {

                            if (!column.formula || !column.resultType)
                                throw new Error('Calculated columns must have a formula and resultType set');

                            var forumulaXml = '<Formula>' + column.formula + '</Formula>';
                            fieldXml = fieldXml.replace(' />', ' ResultType="' + column.resultType + '">' + forumulaXml + '</Field>');
                        }

                        if (multi == 'true')
                        {
                            fieldXml = fieldXml.replace(' />', ' Mult="TRUE" />');
                        }

                        var field = fields.addFieldAsXml(fieldXml, false, SP.AddFieldOptions.AddToNoContentType);

                        if (hidden != null) {
                            field.set_hidden(hidden);
                        }

                        field.set_title(column.displayName);
                        field.set_required(required);

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
                        else if (column.type.toLowerCase() == 'choice' && column.choices) {
                            var fieldChoice = context.castTo(field, SP.FieldChoice);
                            fieldChoice.set_choices($.makeArray(column.choices));
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
                    .done(createColumns);
            }
            else {
                createColumns();
            }

            return columnsCreated.promise();
        },

        /*
            Create a content type on the host web
            @contentTypes { name: 'Test Content Type', id: '0x0100C4AE7CEF4055486987E22766C23F7F35', group: 'Test' }
        */
        createContentTypes: function (contentTypes) {
            var scopes = [];
            var contentTypes = CIB.utilities.ensureArray(contentTypes);

            var contentTypesCreated = new jQuery.Deferred();

            $.each(contentTypes, function () {
                var contentType = this;

                if (!contentType.name || !contentType.id || !contentType.group)
                    throw new Error('Content Type object must have id, name and group attributes');

                var scope = $.handleExceptionsScope(context, function () {
                    helper.message('Creating content type \'' + contentType.name + '\'');

                    var contentTypes = hostContext.get_web().get_contentTypes();

                    var newContentType = new SP.ContentTypeCreationInformation();
                    newContentType.set_id(contentType.id);
                    newContentType.set_name(contentType.name);
                    newContentType.set_group(contentType.group);

                    contentTypes.add(newContentType);

                });
                scope.successMessage = 'Content type ' + contentType.name + ' created';
                scopes.push(scope);
            });

            helper.executeQuery(scopes, contentTypesCreated);

            return contentTypesCreated.promise();
        },

        /*
            Add a list of site columns to a content type
            @contentTypeId '0x0100C4AE7CEF4055486987E22766C23F7F35'
            @columns [ 'cmppMonth', 'cmppYear' ]
        */
        addColumnsToContentType: function (contentTypeId, columns) {
            var columns = CIB.utilities.ensureArray(columns);

            var columnsAdded = new jQuery.Deferred();

            var fields = hostContext.get_web().get_fields();
            var contentTypes = hostContext.get_web().get_contentTypes();
            var contentType = contentTypes.getById(contentTypeId);
            var fieldLinks = contentType.get_fieldLinks();

            context.load(fieldLinks);

            context.executeQueryAsync(function () {
                var scopes = [];
                var existingColumnNames = [];
                var existingColumnIds = [];
                var fieldLinkEnumerator = fieldLinks.getEnumerator();
                while (fieldLinkEnumerator.moveNext()) {
                    var fieldLink = fieldLinkEnumerator.get_current();
                    existingColumnIds.push(fieldLink.get_id().toString().toLowerCase());
                    existingColumnNames.push(fieldLink.get_name());
                }
                $.each(columns, function () {
                    var column = this.toString();
                    if ($.inArray(column, existingColumnNames) >= 0) {
                        helper.message('Column already added to content type \'' + column +
                            '\'. (expected if provisioned already)', 'info');
                        return;
                    }
                    var scope = $.handleExceptionsScope(context, function () {
                        var field = fields.getByInternalNameOrTitle(column);
                        var fieldRef = new SP.FieldLinkCreationInformation();
                        var contentTypeField = fieldRef.set_field(field);
                        fieldLinks.add(fieldRef);
                    });
                    scope.successMessage = 'Added column ' + column + ' to content type';
                    scopes.push(scope);
                });

                contentType.update(true);

                helper.executeQuery(scopes, columnsAdded);

            }, function (sender, args) {
                var error = helper.handleError(sender, args);
                if (error.handled) { columnsAdded.resolve(error.message); }
                else { columnsAdded.reject(error.message); }
            });

            return columnsAdded.promise();
        },

        /*
            Hide columns in a list from the default list edit forms
            @listTitle 'Documents'
            @columns [ 'cmppMonth', 'cmppYear' ]
        */
        hideColumnsFromEditForm: function (listTitle, columns) {
            var scopes = [];
            var columns = CIB.utilities.ensureArray(columns);

            var columnsHid = new jQuery.Deferred();

            helper.message('Hiding columns in list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);
            var listfields = list.get_fields();

            context.load(listfields);

            context.executeQueryAsync(function () {
                $.each(columns, function () {
                    var column = this;

                    var scope = $.handleExceptionsScope(context, function () {
                        var field = listfields.getByInternalNameOrTitle(column);
                        field.setShowInEditForm(false);
                        field.update();
                    });

                    scope.successMessage = 'Column "' + column + '" hidden from edit view';
                    scopes.push(scope);
                });
            }, function (sender, args) {
                var error = helper.handleError(sender, args);
                if (error.handled) { columnsHid.resolve(error.message); }
                else { columnsHid.reject(error.message); }
            });

            helper.executeQuery(scopes, columnsHid);

            return columnsHid.promise();
        },

        /*
           Hide columns in a list from the default list edit forms
           @listTitle 'Documents'
           @viewName 'Important Documents'
           @viewField ['Title']
           @query '<Where></Where>
           @viewType SP.ViewType.calendar
       */
        createView: function (listTitle, viewName, viewFields, query, viewType, rowLimit, paged) {
            var scopes = [];

            var viewCreated = new jQuery.Deferred();

            helper.message('Creating view ' + viewName + ' for list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);
            var views = list.get_views();
            var viewFields = $.ensureArray(viewFields);

            context.load(views, 'Include(Title, ViewFields)');

            context.executeQueryAsync(function () {

                var currentView;
                var viewEnumerator = views.getEnumerator();

                while (viewEnumerator.moveNext()) {
                    var existingView = viewEnumerator.get_current();
                    if (viewName == existingView.get_title()) {
                        helper.message('View \'' + viewName + '\' already exists for list ' + listTitle + '.', 'info');
                        currentView = existingView;
                        break;
                    }
                }

                var scope = $.handleExceptionsScope(context, function () {

                    //If paged is true,"Display items in batches of the specified size" is checked
                    //If paged is false, "Limit the total number of items returned to the specified amount" is checked                                         
                    if (currentView) {
                        var currentViewFields = [];
                        var fields = currentView.get_viewFields();
                        var viewFieldEnumerator = fields.getEnumerator();

                        while (viewFieldEnumerator.moveNext()) {
                            currentViewFields.push(viewFieldEnumerator.get_current());
                        }

                        viewFields.forEach(function (fieldName, index) {
                            if (currentViewFields.indexOf(fieldName) < 0)
                                fields.add(fieldName);
                        });

                        if (query)
                            currentView.set_viewQuery(query);
                        //set row limit for the view                       
                        if (rowLimit)
                            currentView.set_rowLimit(rowLimit);
                        if ((paged != undefined) && (paged != null))
                            currentView.set_paged(paged);

                        currentView.update();
                    }
                    else {
                        var view = new SP.ViewCreationInformation();
                        view.set_title(viewName);
                        view.set_viewFields(viewFields);
                        view.set_query(query);
                        //set row limit for the view                       
                        if (rowLimit)
                            view.set_rowLimit(parseInt(rowLimit));
                        if (viewType)
                            view.set_viewTypeKind(viewType);                                                  
                        if ((paged != undefined) && (paged != null))
                            view.set_paged(paged);

                        views.add(view);
                    }
                });

                scope.successMessage = viewName + (currentView ? ' updated' : ' created ') + 'for list \'' + listTitle + '\'';
                scopes.push(scope);

                helper.executeQuery(scopes, viewCreated);

            }, function (sender, args) {
                var error = helper.handleError(sender, args);
                if (error.handled) { viewCreated.resolve(error.message); }
                else { viewCreated.reject(error.message); }
            });

            return viewCreated.promise();
        },

        /*
            Adds existing content types to a list
            @listTitle 'Documents'
            @contentTypeIds [ '0x0100C4AE7CEF4055486987E22766C23F7F35' ]
        */
        addContentTypesToList: function (listTitle, contentTypeIds) {
            var scopes = [];
            var contentTypeIds = CIB.utilities.ensureArray(contentTypeIds);

            var contentTypesAdded = new jQuery.Deferred();

            helper.message('Adding content types to list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);

            list.set_contentTypesEnabled(true);

            var contentTypes = web.get_contentTypes();
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
           Removed an existing content types from a list
           @listTitle 'Documents'
           @contentTypeNames [ 'Docuemnt' ]
       */
        removeContentTypesFromList: function (listTitle, contentTypeNames) {
            var scopes = [];
            var contentTypeNames = CIB.utilities.ensureArray(contentTypeNames);

            var contentTypesRemoved = new jQuery.Deferred();

            helper.message('Removing content types from list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);

            list.set_contentTypesEnabled(true);

            var listContentTypes = list.get_contentTypes();

            context.load(listContentTypes, 'Include(Id, Name)');

            context.executeQueryAsync(function () {

                $.each(contentTypeNames, function () {
                    var contentTypeName = this;
                    var found = false;
                    var contentTypeEnumerator = listContentTypes.getEnumerator();
                    while (contentTypeEnumerator.moveNext()) {
                        var contentType = contentTypeEnumerator.get_current();
                        if (contentType.get_name().toLowerCase() == contentTypeName.toLowerCase()) {
                            var scope = $.handleExceptionsScope(context, function () {
                                contentType.deleteObject();
                            });

                            scope.successMessage = 'Content type ' + contentTypeName + ' removed from list';
                            scopes.push(scope);

                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        helper.message('Could not find \'' + contentTypeName + '\' in list ' + listTitle + '.', 'info');
                    }
                });

                helper.executeQuery(scopes, contentTypesRemoved);

            }, function (sender, args) {
                var error = helper.handleError(sender, args);
                if (error.handled) { contentTypesRemoved.resolve(error.message); }
                else { contentTypesRemoved.reject(error.message); }
            });

            return contentTypesRemoved.promise();
        },

        /*
           Sets the default content type for a list
           @listTitle 'Documents'
           @contentTypeName 'Document'
        */
        setDefaultContentType: function (listTitle, contentTypeName) {
            var contentTypeSet = new jQuery.Deferred();

            helper.message('Setting default content type on list \'' + listTitle + '\' to ' + contentTypeName);

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);
            var listContentTypes = list.get_contentTypes();
            var rootFolder = list.get_rootFolder();

            context.load(rootFolder, 'ContentTypeOrder', 'UniqueContentTypeOrder');
            context.load(listContentTypes, 'Include(Id, Name)');

            context.executeQueryAsync(function () {
                var scopes = [];

                var scope = $.handleExceptionsScope(context, function () {
                    var newOrder = new Array();
                    var contentTypeEnumerator = listContentTypes.getEnumerator();
                    var order = rootFolder.get_contentTypeOrder();
                    while (contentTypeEnumerator.moveNext()) {
                        var contentType = contentTypeEnumerator.get_current();
                        if (contentType.get_name().toLowerCase() == "folder")
                            continue;
                        if (contentType.get_name().toLowerCase() == contentTypeName.toLowerCase()) {
                            newOrder.splice(0, 0, contentType.get_id());
                            continue;
                        }
                        for (var i = 0; i < order.length; i++) {
                            if (order[i].toString() == contentType.get_id()) {
                                newOrder.push(contentType.get_id());
                                break;
                            }
                        }
                    }
                    rootFolder.set_uniqueContentTypeOrder(newOrder);
                    rootFolder.update();
                });

                scope.successMessage = 'Default content type set on list \'' + listTitle + '\' to ' + contentTypeName;
                scopes.push(scope);

                helper.executeQuery(scopes, contentTypeSet);

            }, function (sender, args) {
                var error = helper.handleError(sender, args);
                if (error.handled) { contentTypeSet.resolve(error.message); }
                else { contentTypeSet.reject(error.message); }
            });

            return contentTypeSet.promise();
        },

        /*
           Creates an index for columns in a list
           @listTitle 'Documents'
           @indicies [ 'cmppMonth', 'cmppYear' ]
        */
        addIndiciesToList: function (listTitle, indicies) {
            var scopes = [];
            var indicies = CIB.utilities.ensureArray(indicies);

            var indiciesSet = new jQuery.Deferred();

            helper.message('Setting indicies on list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);
            var listFields = list.get_fields();

            $.each(indicies, function () {
                var fieldName = this;
                var scope = $.handleExceptionsScope(context, function () {
                    var field = listFields.getByInternalNameOrTitle(fieldName);
                    field.set_indexed(true);
                    field.update();
                    list.update();
                });
                scope.successMessage = 'Index created on column ' + fieldName + ' in list \'' + listTitle + '\'';
                scopes.push(scope);
            });

            helper.executeQuery(scopes, indiciesSet);

            return indiciesSet.promise();
        },

        /*
           Enforces unique values for columns on a list
           @listTitle 'Documents'
           @columns [ 'Title' ]
        */
        enforceUniqueValues: function (listTitle, columns) {
            var scopes = [];
            var columns = CIB.utilities.ensureArray(columns);

            var uniqueValuesEnforced = new jQuery.Deferred();

            helper.message('Enforcing unique values on list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);
            var listFields = list.get_fields();

            $.each(columns, function () {
                var fieldName = this;
                var scope = $.handleExceptionsScope(context, function () {
                    var field = listFields.getByInternalNameOrTitle(fieldName);
                    field.set_indexed(true);
                    field.set_enforceUniqueValues(true);
                    field.update();
                    list.update();
                });
                scope.successMessage = 'Enforced unique values on column ' + fieldName + ' in list \'' + listTitle + '\'';
                scopes.push(scope);
            });

            helper.executeQuery(scopes, uniqueValuesEnforced);

            return uniqueValuesEnforced.promise();
        },

        addListViewWebPartToPage: function (url, listName, viewName, title, zone, index) {

            var webpartAdded = new jQuery.Deferred();

            CIB.installer.addWebPartsToPage({
                url: url,
                title: title,
                assembly: 'Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c',
                type: 'Microsoft.SharePoint.WebPartPages.XsltListViewWebPart',
                zone: zone,
                index: index,
                properties: '<property name="ListUrl" type="string">' + listName + '</property>'
            }).done(function (messages, definition) {

                /*  Cannot chagne the ViewGuid of XsltListViewWebPart
                    This is a CSOM issue, CSOM only let's you update properties on the current class and not it's base type.
                    This means we cannot set what view to use as it's on the BaseXsltListViewWebPart
                    Instead, as a workaround we modify the view to match the one specified
                    This works for standard views, the type of view can't be changed so for calendars this approach will not work
                */

                var webpartId = definition.get_id();
                var list = hostContext.get_web().get_lists().getByTitle(listName.replace('Lists/', ''));
                var view = list.get_views().getById(webpartId);
                var modelView = list.get_views().getByTitle(viewName);
                var modelViewFields = modelView.get_viewFields();
                context.load(view);
                context.load(modelView);
                context.load(modelViewFields);
                context.executeQueryAsyncPromise()
                .done(function () {
                    view.set_viewData(modelView.get_viewData());
                    view.set_viewJoins(modelView.get_viewJoins());
                    view.set_viewProjectedFields(modelView.get_viewProjectedFields);
                    view.set_viewQuery(modelView.get_viewQuery());
                    view.get_viewFields().removeAll();
                    var viewFieldsEnumerator = modelViewFields.getEnumerator();
                    while (viewFieldsEnumerator.moveNext()) {
                        var fieldName = viewFieldsEnumerator.get_current();
                        view.get_viewFields().add(fieldName);
                    }
                    view.update();
                    context.executeQueryAsyncPromise()
                        .done(function () {
                            helper.message('Web part view updated to match ' + viewName, 'success');
                            webpartAdded.resolve();
                        })
                        .fail(function (message) {
                            helper.message(message, 'error');
                            webpartAdded.reject(message);
                        });
                })
                .fail(function (message) {
                    helper.message(message, 'error');
                    webpartAdded.reject(message);
                });
            });

            return webpartAdded.promise();

        },

        /*
           Add a webpart to a page
           @webparts {
                         url: 'Lists/Milestones/EditForm.aspx',
                         title: 'CMPP App View',
                         assembly: 'Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c',
                         type: 'Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart',
                         zone: 'Main',
                         index: 1,
                         properties: '<property name="Content" type="string"><![CDATA[<script>// example script</script>]]></property>'
                     }
        */
        addWebPartsToPage: function (webparts) {
            var scopes = [];
            var webparts = CIB.utilities.ensureArray(webparts);
            var web = hostContext.get_web();

            var webpartsAdded = new jQuery.Deferred();

            $.each(webparts, function () {
                var webpart = this;

                if (!webpart.url || !webpart.title || !webpart.assembly || !webpart.type || !webpart.zone || !webpart.index)
                    throw new Error('Web part object must have url, title, assembly, type, zone and index attributes');

                helper.message('Adding webpart \'' + webpart.title + '\' to file ' + webpart.url + '.');

                var file = web.getFileByServerRelativeUrl(($.getServerRealtiveHostWebUrl() + '/' + webpart.url).replace('//', '/'));
                var webPartManager = file.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                var webparts = webPartManager.get_webParts();
                context.load(webparts, 'Include(WebPart.Title)');
                context.executeQueryAsync(function () {
                    var existingTitles = [];
                    var webPartsEnumerator = webparts.getEnumerator();
                    while (webPartsEnumerator.moveNext()) {
                        var existingWebPart = webPartsEnumerator.get_current().get_webPart();
                        existingTitles.push(existingWebPart.get_title());
                    }
                    var newWebPart;
                    var definition;
                    if ($.inArray(webpart.title, existingTitles) < 0) {
                        var scope = $.handleExceptionsScope(context, function () {
                            var webPartXml = '<?xml version=\"1.0\" encoding=\"utf-8\"?>' +
                                '<webParts>' +
                                '<webPart xmlns="http://schemas.microsoft.com/WebPart/v3">' +
                                '<metaData>' +
                                    '<type name="' + webpart.type + ', ' + webpart.assembly + '" />' +
                                '<importErrorMessage>Cannot import this Web Part.</importErrorMessage>' +
                                '</metaData>' +
                                '<data>' +
                                '<properties>' +
                                    '<property name="Title" type="string">' + webpart.title + '</property>' +
                                    '<property name="ChromeType" type="chrometype">None</property>' +
                                    (webpart.properties ? webpart.properties : '') +
                                '</properties>' +
                                '</data>' +
                                '</webPart>' +
                                '</webParts>';

                            var webPartDefinition = webPartManager.importWebPart(webPartXml);
                            newWebPart = webPartDefinition.get_webPart();
                            definition = webPartManager.addWebPart(newWebPart, webpart.zone, webpart.index);
                            context.load(definition);
                        });

                        scope.successMessage = 'Webpart \'' + webpart.title + '\' added to file ' + webpart.url + '.';
                        scopes.push(scope);

                        helper.executeQuery(scopes, webpartsAdded, definition);
                    }
                    else {
                        helper.message('Webpart \'' + webpart.title + '\' already exists in file ' + webpart.url + '.', 'info');
                        webpartsAdded.resolve();
                        return;
                    }
                }, function (sender, args) {
                    var error = helper.handleError(sender, args);
                    if (!error.handled) { webpartsAdded.reject(error.message); }
                });
            });

            return webpartsAdded.promise();
        },

        /*
           Create a group on the site collection
           @groups { title: 'Project Managers', description: 'Members can manage project details in the system.' }
        */
        createGroup: function (groups) {
            var scopes = [];
            var groups = CIB.utilities.ensureArray(groups);

            var groupsCreated = new jQuery.Deferred();

            //helper.message('Setting indicies on list \'' + listTitle + '\'');

            var web = hostContext.get_web();
            var siteGropus = web.get_siteGroups();

            $.each(groups, function () {
                var group = this;

                if (!group.title || !group.description)
                    throw new Error('Group object must have title and description attributes');

                helper.message('Creating group \'' + group.title + '\'.');

                var scope = $.handleExceptionsScope(context, function () {
                    var newGroup = new SP.GroupCreationInformation();
                    newGroup.set_title(group.title);
                    newGroup.set_description(group.description);
                    siteGropus.add(newGroup);
                });
                scope.successMessage = 'Created group \'' + group.title + '\'.';
                scopes.push(scope);
            });

            helper.executeQuery(scopes, groupsCreated);

            return groupsCreated.promise();
        },

        /*
           Registers the remote event receivers which implement the reusable event receivers pattern
           @serviceUrl: https://application.apps.dev.echonet/Services/appeventreceiver.svc
        */
        registerRemoteEventReceivers: function (serviceUrl) {

            var eventReceiverRegistered = new jQuery.Deferred();

            helper.message('Registering event services at: ' + serviceUrl);

            var soapMessage = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"> \
                    <soap:Body> \
                        <Install xmlns="http://tempuri.org/"> \
                            <hostWebUrl>' + $.getHostWebUrl() + '</hostWebUrl> \
                            <serviceUrl>' + serviceUrl + '</serviceUrl> \
                        </Install> \
                    </soap:Body> \
                </soap:Envelope>';

            $.ajax({
                url: serviceUrl,
                type: "POST",
                dataType: "xml",
                data: soapMessage,
                contentType: "text/xml",
                beforeSend: function (xhr) {
                    xhr.setRequestHeader("SOAPAction", "http://tempuri.org/IInstallableEventService/Install");
                },
                success: function (response) {
                    helper.message('Event services at: ' + serviceUrl + ' registered.', 'success');
                    eventReceiverRegistered.resolve();
                },
                error: function (xhr, status, error) {
                    helper.message('Failed to register event receivers at: ' + serviceUrl + ' (' + error + ').', 'error');
                    eventReceiverRegistered.reject(error);
                }
            });

            return eventReceiverRegistered.promise();

        },

        /* 
            Installs a workflow using a .workflow file definition
            @workflowFileUrl: $.getAppWebUrl() + '/Workflows/A workflow file.workflow'
        */
        installWorkflowFromFile: function (workflowFileUrl) {

            var workflowCreated = new jQuery.Deferred();
            var workflowDependenciesLoaded = new jQuery.Deferred();

            if (!SP.WorkflowServices) {
                $.getScript($.getHostWebUrl() + '/_layouts/15/SP.WorkflowServices.js')
                    .fail(function (error) {
                        helper.message('Failed to load SP.Workflow.js or a depdency', 'error');
                        workflowDependenciesLoaded.reject(error);
                    })
                    .done(function () {
                        if (!SP.WorkflowServices) {
                            helper.message('Failed to load SP.Workflow.js or a depdency', 'error');
                            workflowDependenciesLoaded.reject(error);
                        }
                        else {
                            workflowDependenciesLoaded.resolve();
                        }
                    });
            }
            else {
                workflowDependenciesLoaded.resolve();
            }

            workflowDependenciesLoaded.promise().done(function () {
                $.get(workflowFileUrl)
                    .fail(function (error) {
                        helper.message('Failed to get workflow data from url: ' + workflowFileUrl, 'error');
                        workflowCreated.reject(error);
                    })
                    .done(function (workflow) {
                        if (typeof workflow === 'string') {
                            try {
                                workflow = JSON.parse(workflow);
                            }
                            catch (error) {
                                helper.message('Failed to parse workflow from url: ' + workflowFileUrl, 'error');
                                workflowCreated.reject(error);
                                return;
                            }
                        }

                        CIB.installer.installWorkflow(workflow)
                            .fail(function (error) {
                                helper.message('Failed to install workflow from url: ' + workflowFileUrl, 'error');
                                workflowCreated.reject(error);
                            })
                            .done(function () {
                                workflowCreated.resolve();
                            });
                    });
            });

            return workflowCreated.promise();

        },

        /*
           Creates a workflow definition on the host web
           [Internal use only]
        */
        installWorkflow: function (workflow) {

            var workflowDeifnitionCreated = new jQuery.Deferred();

            if (!$.isInternetExplorer()) {
                helper.message('The installWorkflow method is only supported in internet explorer, the method will run but errors may occur.', 'info');
            }

            if (!workflow.definition || !workflow.associations)
                throw new Error('Workflow data must have "definition" and "associations" properties set');

            if (!workflow.definition.displayName || !workflow.definition.xaml)
                throw new Error('Workflow definition must have at least "displayName" and "xaml" properties set');

            var workflowData = workflow.definition;
            var associations = CIB.utilities.ensureArray(workflow.associations);
            var collateral = CIB.utilities.ensureArray(workflow.collateral);

            helper.message('Creating workflow definition \'' + workflowData.displayName + '\'');

            var workflowWebContext = new SP.ClientContext(jQuery.getHostWebUrl());

            var web = workflowWebContext.get_web();
            var site = workflowWebContext.get_site();

            var workflowServicesManager;
            var workflowDeployment;
            var workflowSubscription;

            var workflowDefinitionId;
            var associationIds = {};

            var handleSharePointFail = function (message) {
                helper.message(message, 'error');
                workflowDeifnitionCreated.reject(message);
            };

            workflowServicesManager = new SP.WorkflowServices.WorkflowServicesManager.newObject(workflowWebContext, workflowWebContext.get_web());
            workflowWebContext.load(web, 'Id', 'Url', 'ServerRelativeUrl');
            workflowWebContext.load(site, 'Id');
            workflowWebContext.load(workflowServicesManager);
            workflowWebContext.executeQueryAsyncPromise()
                .fail(handleSharePointFail)
                .done(function () {
                    workflowDeployment = workflowServicesManager.getWorkflowDeploymentService();
                    workflowSubscription = workflowServicesManager.getWorkflowSubscriptionService();
                    workflowWebContext.load(workflowDeployment);
                    workflowWebContext.load(workflowSubscription);
                    $.when(CIB.installer.getListIds(), workflowWebContext.executeQueryAsyncPromise())
                        .fail(handleSharePointFail)
                        .done(checkExistingWorkflows);
                });

            var checkExistingWorkflows = function (listIds) {
                var workflowDefinitions = workflowDeployment.enumerateDefinitions(false);
                workflowWebContext.load(workflowDefinitions, 'Include(DisplayName, Id)');
                workflowWebContext.executeQueryAsyncPromise()
                    .fail(handleSharePointFail)
                    .done(function () {
                        var workflowDeifnitionIdLoaded = new jQuery.Deferred();
                        var workflowDefinition;
                        var workflowEnumerator = workflowDefinitions.getEnumerator();
                        while (workflowEnumerator.moveNext()) {
                            var workflow = workflowEnumerator.get_current();
                            if (workflow.get_displayName() === workflowData.displayName) {
                                helper.message('Workflow "' + workflowData.displayName + '" already exsists, it will be overwritten');
                                workflowDefinition = workflow;
                                workflowDeifnitionIdLoaded.resolve();
                                break;
                            }
                        }
                        if (!workflowDefinition) {
                            // Save a placeholder workflow to get a persistant id
                            workflowDefinition = new SP.WorkflowServices.WorkflowDefinition.newObject(workflowWebContext, workflowWebContext.get_web());
                            workflowDefinition.set_displayName(workflowData.displayName);
                            workflowDefinition.set_xaml("<Activity mc:Ignorable=\"mwaw\" x:Class=\"Workflow deployment in progress.MTW\" xmlns=\"http://schemas.microsoft.com/netfx/2009/xaml/activities\" xmlns:local=\"clr-namespace:Microsoft.SharePoint.WorkflowServices.Activities\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:mwaw=\"clr-namespace:Microsoft.Web.Authoring.Workflow;assembly=Microsoft.Web.Authoring\" xmlns:scg=\"clr-namespace:System.Collections.Generic;assembly=mscorlib\" xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"><Flowchart StartNode=\"{x:Reference __ReferenceID0}\"><FlowStep x:Name=\"__ReferenceID0\"><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments=\"x:String, x:String\"><x:String x:Key=\"Next\">4294967294</x:String></scg:Dictionary></mwaw:SPDesignerXamlWriter.CustomAttributes><Sequence><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments=\"x:String, x:String\"><x:String x:Key=\"StageAttribute\">StageContainer-8EDBFE6D-DA0D-42F6-A806-F5807380DA4D</x:String></scg:Dictionary></mwaw:SPDesignerXamlWriter.CustomAttributes><local:SetWorkflowStatus Disabled=\"False\" Status=\"Stage 1\"><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments=\"x:String, x:String\"><x:String x:Key=\"StageAttribute\">StageHeader-7FE15537-DFDB-4198-ABFA-8AF8B9D669AE</x:String></scg:Dictionary></mwaw:SPDesignerXamlWriter.CustomAttributes></local:SetWorkflowStatus><Sequence DisplayName=\"Stage 1\" /><Sequence><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments=\"x:String, x:String\"><x:String x:Key=\"StageAttribute\">StageFooter-3A59FA7C-C493-47A1-8F8B-1F481143EB08</x:String></scg:Dictionary></mwaw:SPDesignerXamlWriter.CustomAttributes></Sequence></Sequence></FlowStep></Flowchart></Activity>");
                            workflowDeployment.saveDefinition(workflowDefinition);
                            workflowWebContext.load(workflowDefinition, 'Id');
                            workflowWebContext.executeQueryAsyncPromise()
                                .fail(handleSharePointFail)
                                .done(function () {
                                    workflowDeifnitionIdLoaded.resolve();
                                });
                        }
                        workflowDeifnitionIdLoaded.promise().done(function () {
                            createWithDefinition(workflowDefinition, listIds);
                        });
                    });
            };

            var createWithDefinition = function (workflowDefinition, listIds) {

                var workflowAssociations = {};
                workflowDefinitionId = workflowDefinition.get_id();

                var hasErrors = updateWorkflowTokens(workflowData);

                if (hasErrors) {
                    handleSharePointFail('Failed to replace one or more tokens in the workflow. See installer logs for details');
                    return;
                }

                CIB.utilities.deserialiseSharePointObject(JSON.stringify(workflowData), workflowDefinition);

                if (workflowData['properties']) {
                    var properties = workflowData['properties'];
                    for (var property in properties) {
                        workflowDefinition.setProperty(property, properties[property]);
                    }
                }

                var createCollateral = function () {
                    var collateralCreated = new jQuery.Deferred();

                    if (collateral.length > 0) {
                        var collateralCounter = 0;
                        collateral.forEach(function (collateralFile, index) {

                            var hasErrors = updateWorkflowTokens(collateralFile);

                            if (hasErrors) {
                                handleSharePointFail('Failed to replace one or more tokens in a workflow form');
                                collateralCreated.reject();
                                return false;
                            }

                            helper.message('Uploading workflow file ' + collateralFile.url);

                            var url = ($.getServerRealtiveHostWebUrl() + '/' + collateralFile.url).replace('//', '/');
                            var urlParts = url.split('/');

                            var createInfo = new SP.FileCreationInformation();
                            createInfo.set_content(new SP.Base64EncodedByteArray());
                            for (var i = 0; i < collateralFile.contents.length; i++) {
                                createInfo.get_content().append(collateralFile.contents.charCodeAt(i));
                            }
                            createInfo.set_overwrite(true);

                            createInfo.set_url(urlParts[urlParts.length - 1]);
                            urlParts.splice(urlParts.length - 1, 1);
                            var files = workflowWebContext.get_web().getFolderByServerRelativeUrl(urlParts.join('/')).get_files();
                            var newFile = files.add(createInfo);
                            workflowWebContext.executeQueryAsyncPromise()
                                .fail(handleSharePointFail)
                                .done(function () {
                                    if (++collateralCounter == collateral.length) {
                                        collateralCreated.resolve();
                                    }
                                });
                        });
                    }
                    else {
                        collateralCreated.resolve();
                    }

                    return collateralCreated.promise();
                };

                workflowDefinition.set_draftVersion('');
                workflowDeployment.saveDefinition(workflowDefinition);
                workflowWebContext.load(workflowDefinition, 'Id');

                $.when(createCollateral(), workflowWebContext.executeQueryAsyncPromise())
                    .fail(handleSharePointFail)
                    .done(function () {

                        workflowDeployment.publishDefinition(workflowDefinition.get_id());

                        var existingAssociations = workflowSubscription.enumerateSubscriptionsByDefinition(workflowDefinitionId);

                        workflowWebContext.load(existingAssociations);

                        var validListIds = [];

                        for (var list in listIds)
                            validListIds.push(listIds[list].toString().toLowerCase());

                        workflowWebContext.executeQueryAsyncPromise()
                            .fail(handleSharePointFail)
                            .done(function () {

                                var workflowAssociations = {};
                                var duplicateAssociations = false;
                                var subscriptionEnumerator = existingAssociations.getEnumerator();
                                while (subscriptionEnumerator.moveNext()) {
                                    var subscription = subscriptionEnumerator.get_current();
                                    if (workflowDefinition.get_restrictToType() == 'List') {
                                        var associationListId = subscription.get_eventSourceId().toString();
                                        if (validListIds.indexOf(associationListId.toLowerCase()) < 0) {
                                            // Orphan association, this happens when the list has been deleted
                                            continue;
                                        }
                                    }
                                    if (workflowAssociations[subscription.get_name()]) {
                                        handleSharePointFail('The workflow definition ' + workflowData.displayName + ' has more than one associaiton named ' +
                                            subscription.get_name());
                                        duplicateAssociations = true;
                                        break;
                                    }
                                    workflowAssociations[subscription.get_name()] = subscription;
                                }

                                if (duplicateAssociations)
                                    return;

                                var scopes = [];
                                var hasErrors = false;

                                var publishedCounter = 0;
                                var statusColumnsCreated = new jQuery.Deferred();

                                $.each(associations, function (index, association) {

                                    hasErrors |= updateWorkflowTokens(association);

                                    if (hasErrors) {
                                        handleSharePointFail('Failed to replace one or more tokens in the association ' + association.name +
                                            '. See installer logs for details');
                                        return false;
                                    }

                                    var associationList = workflowWebContext.get_web().get_lists().getById(association['eventSourceId']);
                                    var listFields = associationList.get_fields();
                                    var statusFieldName = association['statusFieldName'];

                                    workflowWebContext.load(listFields, 'Include(InternalName)');
                                    workflowWebContext.executeQueryAsyncPromise()
                                       .fail(handleSharePointFail)
                                       .done(function () {

                                           var createField = true;
                                           var listFieldsEnumerator = listFields.getEnumerator();
                                           while (listFieldsEnumerator.moveNext()) {
                                               var listField = listFieldsEnumerator.get_current();
                                               if (listField.get_internalName() === statusFieldName) {
                                                   createField = false;
                                                   break;
                                               }
                                           }

                                           if (createField) {
                                               var fieldXml = "<Field Type='URL' DisplayName='" + statusFieldName + "' Name='" + statusFieldName + "' />";
                                               var field = listFields.addFieldAsXml(fieldXml, true, SP.AddFieldOptions.addToNoContentType);

                                               var statusFieldDisplayName = unescape(statusFieldName.replace(/_x/g, '%u').replace(/_/g, ''));

                                               field.set_title(statusFieldDisplayName);
                                               field.update();
                                           }

                                           if (++publishedCounter === associations.length) {
                                               statusColumnsCreated.resolve();
                                           }

                                       });
                                });

                                statusColumnsCreated.promise().done(function () {

                                    $.each(associations, function (index, association) {

                                        helper.message('Creating workflow association ' + association.name);

                                        var workflowAssociation = workflowAssociations[association.name];

                                        if (!workflowAssociation) {
                                            workflowAssociation = new SP.WorkflowServices.WorkflowSubscription.newObject(workflowWebContext);
                                        }

                                        CIB.utilities.deserialiseSharePointObject(JSON.stringify(association), workflowAssociation);

                                        if (association['properties']) {
                                            var properties = association['properties'];

                                            for (var property in properties) {
                                                workflowAssociation.setProperty(property, properties[property]);
                                            }
                                        }

                                        if (workflowData.restrictToType == 'List') {
                                            workflowAssociation.setProperty('StatusColumnCreated', '1');
                                            workflowSubscription.publishSubscriptionForList(workflowAssociation, association['eventSourceId']);
                                        }
                                        else if (workflowData.restrictToType == 'Site')
                                            workflowSubscription.publishSubscription(workflowAssociation);
                                        else {
                                            handleSharePointFail('Cannot create association as the restrictToType ' + workflowData.get_restrictToType() + ' was not recognised');
                                            return false;
                                        }
                                    });

                                    if (hasErrors)
                                        return;

                                    workflowWebContext.executeQueryAsyncPromise()
                                        .fail(handleSharePointFail)
                                        .done(function () {
                                            helper.message('Workflow definition ' + workflowData.displayName + ' created', 'success');
                                            workflowDeifnitionCreated.resolve();
                                        });
                                });
                            });
                    });
            };

            var updateWorkflowTokens = function (object) {
                var hasErrors = false;
                var tokens = new RegExp('\{\\$([^\:]*)\:([^\}]*)\}', 'gi');

                var updateTokens = function (object, depth) {
                    for (var property in object) {
                        if (typeof object[property] == 'string') {
                            if (object[property]) {
                                object[property] = object[property].replace(tokens, function (x, tokenType, tokenName) {
                                    if (tokenType === 'List') {
                                        if (listIds[tokenName])
                                            return listIds[tokenName];
                                    }
                                    else if (tokenType === 'Web') {
                                        if (tokenName === '$')
                                            return web.get_id().toString();
                                        else if (tokenName === '%')
                                            return web.get_url().toString();
                                        else if (tokenName === '^')
                                            return web.get_serverRelativeUrl();
                                    }
                                    else if (tokenType === 'Site') {
                                        if (tokenName === '$')
                                            return site.get_id().toString();
                                    }
                                    else if (tokenType === 'Definition') {
                                        if (tokenName === '$')
                                            return workflowDefinitionId.toString();
                                        else if (tokenName === '&')
                                            return workflowDefinitionId.toString().replace(/\-/gi, '');
                                    }
                                    hasErrors = true;
                                    var errorMessage = 'Failed to replace token in workflow, a ' + tokenType + ' cannot be found with the name: ' + tokenName;
                                    helper.message(errorMessage, 'error');
                                    return x;
                                });
                            }
                        } else if (depth < 3) {
                            updateTokens(object[property], (depth + 1));
                        }
                    }
                };
                updateTokens(object, 0);
                return hasErrors;
            };
            return workflowDeifnitionCreated.promise();
        },

        /*
           Adds an accordion group to a list
           listTitle: 'CIB List'
           groupTitle: 'Accordion Headeing'
           fields: ['Title']
        */
        ensureAccordianGroup: function (listTitle, groupTitle, fields) {

            var accordionGroupCreated = new jQuery.Deferred();
            var fields = CIB.utilities.ensureArray(fields);

            if (!listTitle || !groupTitle || !fields || fields.length == 0)
                throw new Error('List title, group title and fields must be set');

            helper.message('Creating accordion group \'' + groupTitle + '\' on list \'' + listTitle + '\'');

            var handleSharePointFail = function (message) {
                helper.message(message, 'error');
                accordionGroupCreated.reject(message);
            };

            var web = hostContext.get_web();
            var list = web.get_lists().getByTitle(listTitle);
            var rootFolder = list.get_rootFolder();
            var properties = rootFolder.get_properties();
            var listFields = {};
            $.each(fields, function (index, field) {
                listFields[field] = list.get_fields().getByInternalNameOrTitle(field);
                context.load(listFields[field], 'InternalName', 'Title', 'Required');
            });
            context.load(properties);
            context.executeQueryAsyncPromise()
                .fail(handleSharePointFail)
                .done(function () {
                    var accordionSettingsValue = properties.get_fieldValues()['CIBListFormAccordionSetting'];
                    var parser = new DOMParser();
                    if (!accordionSettingsValue) {
                        accordionSettingsValue = '<?xml version="1.0" encoding="utf-16"?>\
                            <AccordionSettings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
                                <Groups></Groups>\
                            </AccordionSettings>';
                    }
                    var accordionSettings = parser.parseFromString(accordionSettingsValue, "text/xml");
                    // Check if group already exists
                    var accordionGroup;
                    $.each(accordionSettings.getElementsByTagName('Group'), function (index, group) {
                        var name = group.getElementsByTagName('Name')[0];
                        if (name.textContent == groupTitle) {
                            accordionGroup = group;
                            return false;
                        }
                    });
                    if (!accordionGroup) {
                        var groups = accordionSettings.getElementsByTagName('Groups')[0];
                        var newGroup = accordionSettings.createElement("Group");
                        var fieldsElement = accordionSettings.createElement("Fields");
                        var groupName = accordionSettings.createElement("Name");
                        var groupOrder = accordionSettings.createElement("Order");
                        groupName.appendChild(accordionSettings.createTextNode(groupTitle));
                        groupOrder.appendChild(accordionSettings.createTextNode(groups.childElementCount + 1));
                        newGroup.appendChild(fieldsElement);
                        newGroup.appendChild(groupName);
                        newGroup.appendChild(groupOrder);
                        groups.appendChild(newGroup);
                        accordionGroup = newGroup;
                    }
                    var fieldsPresent = 0;
                    var fieldsUpdated = false;
                    $.each(fields, function (index, field) {
                        var listField = listFields[field];
                        var internalName = listField.get_internalName();
                        var displayName = listField.get_title();
                        var requiredValue = listField.get_required().toString().toLowerCase();

                        var groupFields = accordionGroup.getElementsByTagName('Fields')[0];
                        // Check field already exists
                        var exists = false;
                        $.each(groupFields.getElementsByTagName('Field'), function (index, field) {
                            if (field.getElementsByTagName('InteralName')[0].textContent == internalName) {
                                exists = true;
                                fieldsPresent++;
                                helper.message(displayName + ' is already present in accordion group ' + groupTitle, 'info');
                                return false;
                            }
                        });
                        var existsElsewhere = false;
                        if (!exists) {
                            // Check the field is not already in another group as this would cause errors
                            $.each(accordionSettings.getElementsByTagName('Field'), function (index, field) {
                                if (field.getElementsByTagName('InteralName')[0].textContent == internalName) {
                                    existsElsewhere = true;
                                    return false;
                                }
                            });
                        }
                        if (existsElsewhere) {
                            handleSharePointFail(displayName + ' is already present in a different accordion group');
                            fieldsUpdated = false;
                            return false;
                        }
                        else if (!exists) {
                            var fieldElement = accordionSettings.createElement('Field');
                            var displayNameElement = accordionSettings.createElement('DisplayName');
                            var internalNameElement = accordionSettings.createElement('InteralName');
                            var requiredElement = accordionSettings.createElement('Required');
                            displayNameElement.appendChild(accordionSettings.createTextNode(displayName));
                            internalNameElement.appendChild(accordionSettings.createTextNode(internalName));
                            requiredElement.appendChild(accordionSettings.createTextNode(requiredValue));
                            fieldElement.appendChild(displayNameElement);
                            fieldElement.appendChild(internalNameElement);
                            fieldElement.appendChild(requiredElement);
                            groupFields.appendChild(fieldElement);
                            fieldsUpdated = true;
                        }
                    });
                    if (fieldsUpdated) {
                        // Update xml in property bag
                        var serializer = new XMLSerializer();
                        var accordionXml = serializer.serializeToString(accordionSettings);
                        // Add namespaces in as text, the namespaces aren't used in the document so they can't be added using the api
                        accordionXml = accordionXml.replace('<AccordionSettings>',
                            '<AccordionSettings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');
                        accordionXml = '<?xml version="1.0" encoding="utf-16"?>' + accordionXml;
                        properties.set_item('CIBListFormAccordionSetting', accordionXml);
                        rootFolder.update();
                        context.executeQueryAsyncPromise()
                            .fail(handleSharePointFail)
                            .done(function () {
                                helper.message('Accordion group \'' + groupTitle + '\' added to list \'' + listTitle + '\'', 'success');
                                accordionGroupCreated.resolve();
                            });
                    }
                    else if (fieldsPresent == fields.length) {
                        accordionGroupCreated.resolve();
                    }
                });

            return accordionGroupCreated.promise();
        },

        /*
          Updates the CIB display settings for a list
          listTitle: 'CIB List'
          displaySettings: [{ field: 'Title', form: CIB.installer.displaySettingForm.editForm, display: CIB.installer.displaySettings.whereInGroup, group: 'Approvers' }]
       */

        displaySettings: {
            always: 'always', never: 'never', whereInGroup: 'whereInGroup', whereNotInGroup: 'whereNotInGroup'
        },

        displaySettingForm: {
            displayForm: 'Display', editForm: 'Edit', newForm: 'New'
        },

        displayMode:{
            write: 'writable', read: 'read-only'
        },

        updateDisplaySettings: function (listTitle, displaySettings) {

            var installer = CIB.installer;

            var displaySettingsUpdates = new jQuery.Deferred();
            var displaySettings = CIB.utilities.ensureArray(displaySettings);

            if (!listTitle || !displaySettings || displaySettings.length == 0)
                throw new Error('List title, and display settings must be set');

            $.each(displaySettings, function (index, setting) {
                if (!setting.field || !setting.form || !setting.display)
                    throw new Error('Field, Form and Display properties must be set on the field');
                if (setting.display == installer.displaySettings.whereInGroup || setting.display == installer.displaySettings.whereNotInGroup) {
                    if (!setting.group)
                        throw new Error('Group must be set on the field for where display settings');
                }
                else if (Array.isArray(setting.display)) {
                    $.each(setting.display, function (index, displaySetting) {
                        if (!displaySetting.condition || !displaySetting.groupName || !displaySetting.mode)
                            throw new Error('Condition,Group name and Mode must be set on field for where display settings');
                    })
                    if (!setting.logic)
                        throw new Error('logic condition must be set on multiple group where field');
                }
            });

            helper.message('Updating display settings on list \'' + listTitle + '\'');

            var handleSharePointFail = function (message) {
                helper.message(message, 'error');
                displaySettingsUpdates.reject(message);
            };

            var web = hostContext.get_web();
            var properties = web.get_allProperties();
            var list = web.get_lists().getByTitle(listTitle);
            context.load(list, 'Id');
            context.load(properties);
            context.executeQueryAsyncPromise()
                .fail(handleSharePointFail)
                .done(function () {
                    var listId = list.get_id().toString();
                    var propertyName = ('DisplaySetting' + listId).toLowerCase();
                    var displaySettingItems = properties.get_fieldValues()[propertyName];
                    if (displaySettingItems) {
                        displaySettingItems = displaySettingItems.split('#');
                    }
                    else {
                        displaySettingItems = [];
                    }

                    displaySettingItems = displaySettingItems.filter(function (setting) { return Boolean(setting); });
                    displaySettingItems = displaySettingItems.map(function (setting) { return setting.split('|'); });

                    $.each(displaySettings, function (index, displaySetting) {
                        // Check for existing field
                        var existingIndex = -1;
                        for (var i = 0; i < displaySettingItems.length; i++) {
                            if (displaySettingItems[i][0] === displaySetting.field) {
                                if (displaySettingItems[i][1] === displaySetting.form) {
                                    existingIndex = i;
                                    break;
                                }
                            }
                        }

                        var settingValue = displaySetting.display;
                        if (!Array.isArray(settingValue)) {
                            //Handle single group condition
                            if (settingValue === installer.displaySettings.whereInGroup || settingValue === installer.displaySettings.whereNotInGroup)
                                settingValue = 'where';
                            var settingConfig = settingValue + ';[Me];';
                            if (displaySetting.display === installer.displaySettings.whereNotInGroup) settingConfig += 'IsNotInGroup;'; else settingConfig += 'IsInGroup;';
                            if (displaySetting.group) {
                                    settingConfig += displaySetting.group + ';';
                                    settingConfig += ';writable;~AND';
                            }
                            else 
                                settingConfig += 'Approvers;writable;~AND';
                        }
                        else {
                            //Handle muliple group conditions
                            var settingConfig = 'where';
                            $.each(settingValue, function (idx, group) {
                                settingConfig += ';[Me];';
                                if (group.condition === installer.displaySettings.whereNotInGroup) settingConfig += 'IsNotInGroup;'; else settingConfig += 'IsInGroup;';
                                settingConfig += group.groupName + ';' + group.mode;
                                if(idx !== settingValue.length - 1)   settingConfig += ';$where';
                            })
                            if (displaySetting.logic)
                                settingConfig += ';~' + displaySetting.logic;
                        }
                        if (existingIndex >= 0) {
                            displaySettingItems[existingIndex][0] = displaySetting.field;
                            displaySettingItems[existingIndex][1] = displaySetting.form;
                            displaySettingItems[existingIndex][2] = settingConfig;
                        }
                        else {
                            displaySettingItems.push([displaySetting.field, displaySetting.form, settingConfig]);
                            // Ensure default values are present for all form display modes

                            var displayModeSettings = {};
                            for (var i = 0; i < displaySettingItems.length; i++) {
                                if (displaySettingItems[i][0] === displaySetting.field) {
                                    displayModeSettings[displaySettingItems[i][1]] = true;
                                }
                            }

                            for (var mode in installer.displaySettingForm) {
                                mode = installer.displaySettingForm[mode];
                                if (!displayModeSettings[mode]) {
                                    displaySettingItems.push([displaySetting.field, mode, installer.displaySettings.always + ';[Me];IsInGroup;Approvers;writable;~AND']);
                                }
                            }
                        }
                    });

                    for (var i = 0; i < displaySettingItems.length; i++)
                        displaySettingItems[i] = displaySettingItems[i].join('|');

                    var updatedValue = displaySettingItems.join('#') + '#';

                    properties.set_item(propertyName, updatedValue);
                    web.update();
                    context.executeQueryAsyncPromise()
                        .fail(handleSharePointFail)
                        .done(function () {
                            helper.message('Display settings updated on list \'' + listTitle + '\'', 'success');
                            displaySettingsUpdates.resolve();
                        });

                });

            return displaySettingsUpdates.promise();
        },

        /*
           Updates properties on a pre-existing webpart, such as updating the JSLink on a list form
           @webPartProperties { 
                file: '/teams/sitecolleciton/documents/forms/editForm.aspx' 
                title: 'WebPart title', 
                properties: { 'JSLink': '~/sitecollection/Style Library/customer/customization.js' }
           }
        */
        updateWebPartProperties: function (webPartProperties) {
            var scopes = [];
            var webPartProperties = CIB.utilities.ensureArray(webPartProperties);

            var webPartsUpdated = new jQuery.Deferred();

            var handleSharePointFail = function (message) {
                helper.message(message, 'error');
                webPartsUpdated.reject(message);
            };

            var web = hostContext.get_web();
            var updateCount = 0;

            $.each(webPartProperties, function () {
                var webPartDetails = this;

                if (!webPartDetails.title || !webPartDetails.file || !webPartDetails.properties)
                    throw new Error('web part must have title, file and properties attributes set');

                helper.message('Updating web part \'' + webPartDetails.title + '\'.');

                var file = web.getFileByServerRelativeUrl(webPartDetails.file);
                var webPartManager = file.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                var webparts = webPartManager.get_webParts();
                context.load(webparts, 'Include(WebPart.Title, WebPart.Properties)');
                context.executeQueryAsyncPromise()
                    .fail(handleSharePointFail)
                    .done(function () {
                        var webPartUpdated = false;
                        var webPartsEnumerator = webparts.getEnumerator();
                        while (webPartsEnumerator.moveNext()) {
                            var webpartDefinition = webPartsEnumerator.get_current();
                            var webpart = webpartDefinition.get_webPart();
                            if (webpart.get_title() === webPartDetails.title || webPartDetails.title === '*') {
                                webPartUpdated = true;
                                for (var propertyName in webPartDetails.properties) {
                                    var propertyValue = webPartDetails.properties[propertyName];
                                    webpart.get_properties().set_item(propertyName, propertyValue);
                                }
                                webpartDefinition.saveWebPartChanges();
                            }
                        }
                        if (webPartUpdated) {
                            context.executeQueryAsyncPromise()
                            .fail(handleSharePointFail)
                            .done(function () {
                                helper.message('Updated properties for web part \'' + webPartDetails.title + '\'.', 'success');
                                if (++updateCount == webPartProperties.length)
                                    webPartsUpdated.resolve();
                            });
                        }
                        else {
                            helper.message('No web part with a title \'' + webPartDetails.title + '\' was found in file: ' + webPartDetails.file, 'info');
                            if (++updateCount == webPartProperties.length)
                                webPartsUpdated.resolve();
                        }
                    });

            });

            return webPartsUpdated.promise();
        },

        getContentTypeIdByName: function(contentTypeName) {
            if (!contentTypeName)
                throw new Error('Content Type Name cannot be null');

            var contentTypeRetrieved = new jQuery.Deferred();
            var handleSharePointFail = function (message) {
                helper.message(message, 'error');
                contentTypeRetrieved.reject(message);
            };
            helper.message("Fetching content type id for name: " + contentTypeName);
            var contentTypeId = '';
            var web = hostContext.get_web();
            var contentTypeCollection = web.get_availableContentTypes();
            context.load(contentTypeCollection, 'Include(Id, Name)');
            context.executeQueryAsyncPromise()
                .fail(handleSharePointFail)
                .done(function() {
                    var contentTypeEnumerator = contentTypeCollection.getEnumerator();
                    while (contentTypeEnumerator.moveNext()) {
                        var contentType = contentTypeEnumerator.get_current();
                        if (contentType.get_name() === contentTypeName) {
                            contentTypeId = contentType.get_id();
                            break;
                        }
                    }
                    contentTypeRetrieved.resolve(contentTypeId);
                });
            return contentTypeRetrieved.promise();
        }

    };
}();

// Fix for SP.Requestexecutor for binary files
(function () {

    if (!SP.RequestExecutor)
        return;

    SP.RequestExecutorInternalSharedUtility.BinaryDecode = function SP_RequestExecutorInternalSharedUtility$BinaryDecode(data) {
        var ret = '';
        if (data) {
            var byteArray = new Uint8Array(data);
            for (var i = 0; i < data.byteLength; i++) {
                ret = ret + String.fromCharCode(byteArray[i]);
            }
        }
        ;
        return ret;
    };
    SP.RequestExecutorUtility.IsDefined = function SP_RequestExecutorUtility$$1(data) {
        var nullValue = null;
        return data === nullValue || typeof data === 'undefined' || !data.length;
    };
    SP.RequestExecutor.ParseHeaders = function SP_RequestExecutor$ParseHeaders(headers) {
        if (SP.RequestExecutorUtility.IsDefined(headers)) {
            return null;
        }
        var result = {};
        var reSplit = new RegExp('\r?\n');
        var headerArray = headers.split(reSplit);
        for (var i = 0; i < headerArray.length; i++) {
            var currentHeader = headerArray[i];
            if (!SP.RequestExecutorUtility.IsDefined(currentHeader)) {
                var splitPos = currentHeader.indexOf(':');
                if (splitPos > 0) {
                    var key = currentHeader.substr(0, splitPos);
                    var value = currentHeader.substr(splitPos + 1);
                    key = SP.RequestExecutorNative.trim(key);
                    value = SP.RequestExecutorNative.trim(value);
                    result[key.toUpperCase()] = value;
                }
            }
        }
        return result;
    };
    SP.RequestExecutor.internalProcessXMLHttpRequestOnreadystatechange = function SP_RequestExecutor$internalProcessXMLHttpRequestOnreadystatechange(xhr, requestInfo, timeoutId) {
        if (xhr.readyState === 4) {
            if (timeoutId) {
                window.clearTimeout(timeoutId);
            }
            xhr.onreadystatechange = SP.RequestExecutorNative.emptyCallback;
            var responseInfo = new SP.ResponseInfo();
            responseInfo.state = requestInfo.state;
            responseInfo.responseAvailable = true;
            if (requestInfo.binaryStringResponseBody) {
                responseInfo.body = SP.RequestExecutorInternalSharedUtility.BinaryDecode(xhr.response);
            }
            else {
                responseInfo.body = xhr.responseText;
            }
            responseInfo.statusCode = xhr.status;
            responseInfo.statusText = xhr.statusText;
            responseInfo.contentType = xhr.getResponseHeader('content-type');
            responseInfo.allResponseHeaders = xhr.getAllResponseHeaders();
            responseInfo.headers = SP.RequestExecutor.ParseHeaders(responseInfo.allResponseHeaders);
            if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 1223) {
                if (requestInfo.success) {
                    requestInfo.success(responseInfo);
                }
            }
            else {
                var error = SP.RequestExecutorErrors.httpError;
                var statusText = xhr.statusText;
                if (requestInfo.error) {
                    requestInfo.error(responseInfo, error, statusText);
                }
            }
        }
    };



})();