'use strict';

/*

    GlobalNavigation.js
    Provides a framework inheriting the global navigation of the host web

*/

var CIB = CIB || {};

CIB.navigation = function () {

    var globalContext = CIB.utilities.getContext();
    var context = globalContext.context;
    var hostContext = globalContext.hostContext;

    var isEditMode = false;
    var loaded = false;
    var navigationListExists = false;

    var navigationache = new CIB.utilities.CIBcache('CIB.App.NavigationLinks');

    $(document).ready(function () {
        CIB.navigation.load();
    });

    var getDeleteLink = function () {
        var deleteLink = $('<a class="nav-editmode-only ms-navedit-deletelink" title="Remove this link from navigation"><span class="ms-navedit-deletespan"><img class="ms-cancelImg" src="' + $.getHostWebUrl() + '/_layouts/15/defaultcss.ashx?ctag=2&resource=spcommon" width="271" height="268"></span></a>');

        deleteLink.click(function () {

            var id = $(this).prev().attr('data-item-id');
            var container = $(this).parent();
            var list = hostContext.get_web().get_lists().getByTitle('App Navigation Links');
            var navitem = list.getItemById(id);
            navitem.deleteObject();
            context.executeQueryAsyncPromise()
                   .fail(logAndReportError)
                   .done(function () {
                       container.remove();
                   })
                   .always(function() {
                       navigationache.invalidate();
                   });


        });

        return deleteLink;
    }

    var editNavigationLinks = function() {
        var addNewLink = $('<a class="ms-heroCommandLink ms-navedit-addNewLink" title="Add a link"><span class="ms-list-addnew-imgSpan16"><img class="ms-list-addnew-img16" src="' + $.getHostWebUrl() + '/_layouts/15/defaultcss.ashx?ctag=2&resource=spcommon"></span><span class="ms-navedit-addLinkText">link</span></a>');

        addNewLink.click(function () {
            $('#editLinkId').val('');
            $('#editLinkTitle').val('');
            $('#editLinkUrl').val('');
            $('#editLinkModel').modal('show');
        });

        var doneButton = $('<input class="ms-navedit-editButton" title="Done" value="Done" type="button"></button>')

        doneButton.click(finishEditing);

        $('.ms-core-listMenu-root').append($('<li class="nav-editmode-only add-new-link-item ui-state-disabled"></li>').append(addNewLink).append(doneButton));
        $('.ms-core-navigation .nav-item').append(getDeleteLink());
        $('.ms-listMenu-editLink').hide();
        $('.ms-core-navigation .nav-item').addClass('link-edit-mode');
        $('.sortable-link').removeClass('ui-state-disabled');
        isEditMode = true;
    };

    var finishEditing = function () {

        // Update sort order
        var order = 0;
        var list = hostContext.get_web().get_lists().getByTitle('App Navigation Links');
        $('.ms-core-listMenu-root .sortable-link').each(function () {
            order++;
            var node = $(this);
            var link = node.children('.ms-navedit-linkNode');
            var id = link.attr('data-item-id');
            var item = list.getItemById(id);
            item.set_item('Link_x0020_Order', order);
            item.update();
        });

        context.executeQueryAsyncPromise()
                   .fail(logAndReportError)
                   .done(function () {})
                   .always(function () {
                       $('.sortable-link').addClass('ui-state-disabled')
                       $('.ms-core-navigation .nav-item').removeClass('link-edit-mode');
                       $('.nav-editmode-only').remove();
                       $('.ms-listMenu-editLink').show();
                       isEditMode = false;
                       navigationache.invalidate();
                   });

    };

    var addNewLinkModalToPage = function () {

        var modal = $('<div id="editLinkModel" class="modal fade" aria-hidden="true" role="dialog">\
          <div class="modal-dialog">\
            <div class="modal-content">\
              <div class="modal-body">\
                <input id="editLinkId" type="hidden" />\
                <div class="link-input"><span>Title: </span><input id="editLinkTitle"  class="form-control" type="text" /></div>\
                <div class="link-input"><span>Url: </span><input id="editLinkUrl" class="form-control" type="text" /></div>\
              </div>\
              <div class="modal-footer">\
                <button type="button" id="saveLinkChanges" data-loading-text="Loading..."  class="btn btn-success">Save changes</button>\
                <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>\
              </div>\
            </div>\
          </div>\
        </div>');

        $('body').append(modal);
        
        modal.find('#saveLinkChanges').click(saveLinkChanges);

        modal.modal({ show: false });
    }

    var logAndReportError = function (message) {
        CIB.logging.logError('Global Navigation', message);
        $('.mp-breadcrumb-top .alert').remove();
        $('.mp-breadcrumb-top').append('<div class="alert alert-danger" role="alert" style="display:block; margin-top:20px">' +
            '<strong>Error updating navigation, please contact support: </strong><span>' + message + '</span></div>');

        $('#editLinkModel').modal('hide');
        $('#saveLinkChanges').bootstrapBtn('reset');
    }

    var saveLinkChanges = function () {

        $('#saveLinkChanges').bootstrapBtn('loading');

        if (navigationListExists) {
            var itemId = $('#editLinkId').val();
            var title = $('#editLinkTitle').val();
            var url = $('#editLinkUrl').val();

            var item;
            var list = hostContext.get_web().get_lists().getByTitle('App Navigation Links');

            if (!itemId) {
                var newItemInfo = new SP.ListItemCreationInformation();
                var item = list.addItem(newItemInfo);
            }
            else {
                item = list.getItemById(itemId);
            }

            item.set_item('Title', title);
            item.set_item('Link_x0020_Url', url);
            item.set_item('App_x0020_Url', window.location.host);
            item.set_item('Link_x0020_Order', $('.ms-core-listMenu-root .menu-item').length + 1);

            item.update();
            context.load(item);

            context.executeQueryAsyncPromise()
                   .fail(logAndReportError)
                   .done(function () {
                       if (!itemId) {
                           var menuItem = getMenuItemFromListItem(item);
                           menuItem.insertBefore('.ms-listMenu-editLink');
                           menuItem.append(getDeleteLink());
                           menuItem.removeClass('ui-state-disabled');
                           menuItem.addClass('link-edit-mode');
                           $('.mp-breadcrumb-top .alert').remove();
                       }
                       else {
                           var source = $('a[data-item-id="' + itemId + '"]');
                           source.attr('href', url);
                           source.children('span').text(title);
                       }

                       $('#editLinkModel').modal('hide');
                       $('#saveLinkChanges').bootstrapBtn('reset');
                   })
                   .always(function () {
                       navigationache.invalidate();
                   });
        }
        else {
            var lists = hostContext.get_web().get_lists();

            var newList = new SP.ListCreationInformation();
            newList.set_title('App Navigation Links');
            newList.set_templateType(100);

            var appNavigationList = lists.add(newList);
            var fields = appNavigationList.get_fields();

            fields.addFieldAsXml("<Field Type='URL' DisplayName='Link Url' />", false, SP.AddFieldOptions.AddToNoContentType);
            fields.addFieldAsXml("<Field Type='Number' DisplayName='Link Order' />", false, SP.AddFieldOptions.AddToNoContentType);
            fields.addFieldAsXml("<Field Type='Text' DisplayName='App Url' />", false, SP.AddFieldOptions.AddToNoContentType);

            context.executeQueryAsyncPromise()
                   .fail(logAndReportError)
                   .done(function () {
                       navigationListExists = true;
                       $('.mp-breadcrumb-top .alert').remove();
                       saveLinkChanges();
                   })
                  .always(function () {
                      navigationache.invalidate();
                  });
        }
    }

    var getMenuItemFromListItem = function (item) {
        return getMenuItem(item.get_item('Title'), item.get_item('Link_x0020_Url').get_url(), item.get_item('ID'));
    }

    var getMenuItem = function (title, url, id) {
        var anchor = $('<a class="static selected ms-navedit-dropNode menu-item ms-core-listMenu-item ms-displayInline ms-core-listMenu-selected ms-navedit-linkNode"><span class="menu-item-text">' + title + '</span></a>');
        anchor.attr('href', url);
        anchor.attr('data-item-id', id);

        anchor.click(function (e) {
            var link = $(this);
            if (isEditMode) {

                $('#editLinkId').val(link.attr('data-item-id'));
                $('#editLinkTitle').val(link.text());
                $('#editLinkUrl').val(link.attr('href'));
                $('#editLinkModel').modal('show');

                e.stopPropagation();
                return false;
            }
        });

        var linkItem = $('<li class="static sortable-link nav-item ms-navedit-dropNode ui-state-disabled"></li>')
        linkItem.append(anchor);
        return linkItem;
    }

    return {

        globalLinks: [],

        load: function () {
            
            if (loaded || $('.mp-breadcrumb-top').length == 0)
                return;

            loaded = true;

            var navigation = $('<div id="DeltaTopNavigation" class="ms-displayInline ms-core-navigation" role="navigation"></div>');
            var menu = $('<div class="noindex ms-core-listMenu-horizontalBox"></div>');
            var menuList = $('<ul class="root ms-core-listMenu-root static"></ul>');

            var editLinks = $('<a class="ms-navedit-editLinksText"><span class="ms-displayInlineBlock"><span class="ms-navedit-editLinksIconWrapper ms-verticalAlignMiddle"><img class="ms-navedit-editLinksIcon" src="' + $.getHostWebUrl() + '/_layouts/15/defaultcss.ashx?ctag=2&resource=spcommon"></span><span class="ms-metadata ms-verticalAlignMiddle">Edit Links</span></span></a></a>');
            var editLinksWrapper = $('<span class="ms-navedit-editSpan"></span>');
            var editLinkItem = $('<li class="static ms-verticalAlignTop ms-listMenu-editLink ms-navedit-editArea ui-state-disabled"></li>');

            if (CIB.navigation.globalLinks.length == 0 && ($.hasAppWeb() || $.isInternetExplorer())) {

                var hasEditPermissions = false;
                var navigationLinks = [];

                var loadNavigation = new jQuery.Deferred();

                if (navigationache.containsValue()) {
                    var data = navigationache.get();
                    hasEditPermissions = data.editPermissions;
                    navigationLinks = data.navigationLinks;
                    navigationListExists = data.navigationListExists;
                    loadNavigation.resolve();
                }
                else {
                    var web = hostContext.get_web();
                    
                    var requiredPermissions = new SP.BasePermissions();
                    requiredPermissions.set(SP.PermissionKind.manageWeb);
                    var editPermissions = web.doesUserHavePermissions(requiredPermissions);

                    var links;

                    var scope = $.handleExceptionsScope(context, function () {

                        var list = web.get_lists().getByTitle('App Navigation Links');

                        var query = new SP.CamlQuery();

                        query.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'App_x0020_Url\' />' +
                            '<Value Type=\'Text\'>' + window.location.host + '</Value></Eq></Where>' +
                            '<OrderBy><FieldRef Name="Link_x0020_Order"/></OrderBy>' +
                            '</Query><RowLimit>100</RowLimit></View>');

                        links = list.getItems(query);
                        context.load(links, 'Include(ID, Title, Link_x0020_Url)');
                    });

                    var failHandler = function (message) {
                        if (message.indexOf("List 'App Navigation Links' does not exist") != 0 && message.indexOf('Access denied.') != 0) {
                            logAndReportError(message);
                        }

                        hasEditPermissions = editPermissions.get_value();

                        loadNavigation.reject(message);
                    };

                    context.executeQueryAsyncPromise()
                       .fail(failHandler)
                       .done(function () {

                           if (scope.get_hasException()) {
                               failHandler(scope.get_errorMessage());
                               return;
                           }

                           navigationListExists = true;

                           var linksEnumerator = links.getEnumerator();
                           while (linksEnumerator.moveNext()) {
                               var link = linksEnumerator.get_current();
                               navigationLinks.push({ title: link.get_item('Title'), url: link.get_item('Link_x0020_Url').get_url(), id: link.get_item('ID') });
                           }

                           hasEditPermissions = editPermissions.get_value();

                           navigationache.set({ editPermissions: hasEditPermissions, navigationLinks: navigationLinks, navigationListExists: navigationListExists });

                           loadNavigation.resolve();

                       });

                }

                loadNavigation.always(function () {
                    for (var index in navigationLinks) {
                        var link = navigationLinks[index];
                        var linkItem = getMenuItem(link.title, link.url, link.id);
                        menuList.append(linkItem);
                    }

                    if (hasEditPermissions) {
                        if ($.isInternetExplorer()) {
                            addNewLinkModalToPage();
                            editLinks.click(editNavigationLinks);
                            editLinksWrapper.append(editLinks);
                            editLinkItem.append(editLinksWrapper);
                            menuList.append(editLinkItem);
                        }
                    }
                });
            }
            else {
                for (var index in CIB.navigation.globalLinks) {
                    var link = CIB.navigation.globalLinks[index];
                    menuList.append(getMenuItem(link.title, link.url, ''));
                }
            }

            menu.append(menuList);
            navigation.append(menu);
            $('.mp-breadcrumb-top').append(navigation);
            $('.ms-core-listMenu-root').sortable({
                cursor: "move",
                items: "li:not(.ui-state-disabled)",
                start: function (e, ui) {
                    $(e.target).data("ui-sortable").floating = true;
                },
            });
        }

    }

}();
