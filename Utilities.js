'use strict';

/*

    Utilities.js
    Provides a framework for common javascript or SharePoint functions

*/

var CIB = CIB || {};

if (!(window.console && console.log)) {
    console = {
        log: function () { },
        debug: function () { },
        info: function () { },
        warn: function () { },
        error: function () { }
    };
}

CIB.utilities = function () {

    if ($.fn && $.fn.button && $.fn.button.noConflict)
        $.fn.bootstrapBtn = $.fn.button.noConflict();

    SP.ClientContext.prototype.executeQueryAsyncPromise = function () {
        var sharePointQuery = new jQuery.Deferred();
        this.executeQueryAsync(function () {
            sharePointQuery.resolve();
        }, function (sender, args) {
            var message = args.get_message ? args.get_message() : args.get_errorMessage();
            sharePointQuery.reject(message, sender, args);
        });
        return sharePointQuery.promise();
    }

    var getPageHeight = function () {

        var elementHeights = $('*').filter(function () {
            var position = $(this).css('position');
            return position === 'absolute';
        }).map(function () { return $(this).offset().top + $(this).height() }).toArray();

        elementHeights.push($('#footerMarker').offset().top);

        return Math.max.apply(this, elementHeights);
    }

    $(document).ready(function () {

        // Setup responsive app part
        $('body').append('<div id="footerMarker" />');

        if ($.isAppPart()) {

            var counter = 0;
            var initialWidth = $(window).width();

            $(window).resize(function () {
                if ($(window).width() != initialWidth) {
                    initialWidth = $(window).width();
                    CIB.utilities.resizeAppPart('100%', getPageHeight());
                }
            });

            CIB.utilities.resizeAppPart('100%', getPageHeight());
        }

    });

    var CIBcache = function (key) {

        var cacheKey = window.location.host + "_" + key;

        return {

            get: function () {
                return JSON.parse(sessionStorage[cacheKey]).data;
            },

            set: function (data) {
                sessionStorage[cacheKey] = JSON.stringify({ data: data });
            },

            invalidate: function () {
                delete sessionStorage[cacheKey];
            },

            containsValue: function () {
                if (sessionStorage[cacheKey])
                    return true;
                return false;
            }

        }

    }

    return {

        CIBcache: CIBcache,

        isAppWeb: function () {
            if ($('meta[name="GENERATOR"]').attr('content') === "Microsoft SharePoint")
                return true;
            return (typeof _spPageContextInfo !== 'undefined') && _spPageContextInfo && jQuery.hasAppWeb();
        },

        hasAppWeb: function () {
            return jQuery.getAppWebUrl() && jQuery.getAppWebUrl() != '""' && jQuery.getAppWebUrl() != "undefined";
        },

        getContext: function () {
            var hasAppWeb = jQuery.isAppWeb() && jQuery.hasAppWeb();
            var context = new SP.ClientContext(hasAppWeb ? jQuery.getAppWebUrl() : jQuery.getHostWebUrl());
            if (hasAppWeb) {
                if (!SP.ProxyWebRequestExecutorFactory)
                    throw new Error('/_layouts/15/sp.requestexecutor.js is not refernced');
                var factory = new SP.ProxyWebRequestExecutorFactory(jQuery.getAppWebUrl());
                context.set_webRequestExecutorFactory(factory);
            }
            var hostContext = hasAppWeb ? new SP.AppContextSite(context, jQuery.getHostWebUrl()) : context;

            var contexts = {
                context: context,
                hostContext: hostContext
            };
            if (jQuery.hasAppWeb()) {
                contexts.appContext = new SP.ClientContext(jQuery.getAppWebUrl());
            }
            return contexts;
        },

        getUrlVars: function () {
            var vars = [], hash;
            var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
            for (var i = 0; i < hashes.length; i++) {
                hash = hashes[i].split('=');
                vars.push(hash[0]);
                vars[hash[0]] = decodeURIComponent(hash[1]);
            }
            return vars;
        },

        getUrlVar: function (name) {
            return jQuery.getUrlVars()[name];
        },

        getHostWebUrl: function () {
            var CIBAppFrameWorkSubWebUrl;
            var HostWebUrl;
            //Only SharePoint hosted app and Provided hosted app has query string "SPHostUrl", this is to find app web or host web.
            if (window.location.href.indexOf("SPHostUrl") > -1) {
                // get control, to fetch sub web url
                var isInput = $("[id$='CIBAppFrameWorkSubWebUrl']").is('input');
                if (isInput) {
                    CIBAppFrameWorkSubWebUrl = $("[id$='CIBAppFrameWorkSubWebUrl']").val();
                }
                else {
                    CIBAppFrameWorkSubWebUrl = $("[id$='CIBAppFrameWorkSubWebUrl']").text();
                }

                if (CIBAppFrameWorkSubWebUrl) {
                    HostWebUrl = CIBAppFrameWorkSubWebUrl;
                }
                else {
                    HostWebUrl = decodeURIComponent(jQuery.getUrlVar('SPHostUrl')).replace(/\#+$/, '');
                }
            }
            //If not app web, pass the current location url, it's used for logger.js.
            else {
                HostWebUrl = decodeURIComponent(window.location.href).replace(/\#+$/, '');
            }

            return HostWebUrl;
        },

        getServerRealtiveHostWebUrl: function () {
            var url = jQuery.getHostWebUrl();
            if (url.replace('http://', '').replace('https://', '').indexOf('/') < 0)
                url += '/';
            return '/' + url.replace(/^(?:\/\/|[^\/]+)*\//, "");
        },

        getServerHostUrl: function () {
            var m = jQuery.getHostWebUrl().match(/^http[s]?:\/\/[^/]+/);
            return m ? m[0] : null;

        },

        getServerAppUrl: function () {
            var m = jQuery.getAppWebUrl().match(/^http[s]?:\/\/[^/]+/);
            return m ? m[0] : null;
        },

        getAppWebUrl: function () {
            return decodeURIComponent(jQuery.getUrlVar('SPAppWebUrl')).replace(/\#+$/, '');
        },

        getServerRealtiveApptWebUrl: function () {
            var url = jQuery.getAppWebUrl();
            if (url.replace('http://', '').replace('https://', '').indexOf('/') < 0)
                url += '/';
            return '/' + url.replace(/^(?:\/\/|[^\/]+)*\//, "");
        },

        ensureArray: function (value) {
            if ($.isArray(value))
                return value;
            else
                return [value]
        },

        handleExceptionsScope: function (context, action) {
            var scope = new SP.ExceptionHandlingScope(context);
            var start = scope.startScope();
            var scopeTry = scope.startTry();
            action();
            scopeTry.dispose();
            var scopeCatch = scope.startCatch();
            scopeCatch.dispose();
            start.dispose();
            return scope;
        },

        isAppPart: function () {
            return jQuery.getUrlVar('SenderId');
        },

        resizeAppPart: function (width, height) {
            if (jQuery.isAppPart) {
                if (!width)
                    width = '100%';
                if (!height)
                    height = getPageHeight();

                window.parent.postMessage('<message senderId=' + jQuery.getUrlVar('SenderId') + '>resize(' + width + ', ' + height + ')</message>', jQuery.getHostWebUrl());
            }
        },

        isInternetExplorer: function () {
            return navigator.appVersion.indexOf('MSIE ') >= 0;
        },

        serialiseSharePointObject: function (object, includeId) {
            includeId = (typeof optionalArg === "undefined") ? false : includeId;
            var data = {};
            var keys = $.map(object, function (v, i) { return i; });
            for (var index in keys) {
                var key = keys[index];
                if (key === 'get_id' && !includeId)
                    continue;
                if (key.startsWith('get_')) {
                    var value = String(object[key].call(object));
                    if (value !== '[object Object]')
                        data[key.replace('get_', '')] = object[key].call(object);
                }
            }
            return JSON.stringify(data);
        },

        deserialiseSharePointObject: function (text, object, includeId) {
            includeId = (typeof optionalArg === "undefined") ? false : includeId;
            var data = JSON.parse(text);
            var keys = $.map(data, function (v, i) { return i; });
            for (var index in keys) {
                var key = keys[index];
                if (key === 'id' && !includeId)
                    continue;
                if (object['set_' + key]) {
                    object['set_' + key].call(object, data[key]);
                }
            }
        },

        changeAppPage: function (pageUrl) {
            var url = jQuery.getAppWebUrl().replace(/\/+$/, '');
            url += '/' + pageUrl.replace(/^\/+/, '');
            url += window.location.href.slice(window.location.href.indexOf('?'));
            window.location = url;
        },

        showHostWebDialog: function (options) {

            if (!options.url)
                throw new Error('Options must include a url from the host web');

            options.url = jQuery.getServerHostUrl() + '/Style Library/CIB/Pages/Dialog.aspx?target=' + encodeURIComponent(options.url) + '&parent=' + jQuery.getServerAppUrl();

            var openDialog = function () {

                var widthAdjusted = false;
                var windowProxy = new Porthole.WindowProxy(jQuery.getServerHostUrl() + '/Style Library/CIB/Pages/DialogProxy.aspx');

                windowProxy.addEventListener(function (result) {
                    if (result.data.hasOwnProperty('result')) {
                        SP.UI.ModalDialog.commonModalDialogClose(result.data.result, result.data.value);
                    }
                    else if (result.data['resize']) {

                        $('.modalLoadingMessage').remove();

                        if (SP.UI.ModalDialog.get_childDialog() != null && SP.UI.ModalDialog.get_childDialog().get_frameElement() != null)
                            SP.UI.ModalDialog.get_childDialog().get_frameElement().parentElement.style['padding-right'] = '0px'

                        var width = result.data.width;
                        var height = result.data.height;

                        var dialogElements = new Array();
                        var getDialogElement = function (elementArray, elementRef) {
                            elementArray[elementRef] = $('.ms-dlg' + elementRef, window.parent.document);
                        };
                        getDialogElement(dialogElements, "Border");
                        getDialogElement(dialogElements, "Title");
                        getDialogElement(dialogElements, "TitleText");
                        getDialogElement(dialogElements, "Content");
                        getDialogElement(dialogElements, "Frame");

                        if (($(window).height() - 20) < height) {
                            height = $(window).height() - 20;
                        }

                        deltaWidth = width - dialogElements["Border"].width();
                        deltaHeight = height - dialogElements["Border"].height();

                        for (var key in dialogElements) {

                            if (!widthAdjusted)
                                dialogElements[key].width(dialogElements[key].width() + deltaWidth);

                            if (key != "Title" && key != "TitleText") {
                                dialogElements[key].height(dialogElements[key].height() + deltaHeight);
                            }
                        }

                        widthAdjusted = true;
                        $('.ms-dlgTitle').css('width', '95%');

                        if (SP.UI.ModalDialog.get_childDialog() != null && SP.UI.ModalDialog.get_childDialog().get_dialogElement() != null) {
                            var dialogElement = $(SP.UI.ModalDialog.get_childDialog().get_dialogElement());
                            dialogElement.css("top", Math.max(0, (($(window).height() - dialogElement.outerHeight()) / 2) + $(window).scrollTop()) + "px");
                            dialogElement.css("left", Math.max(0, (($(window).width() - dialogElement.outerWidth()) / 2) + $(window).scrollLeft()) + "px");
                        }
                    }
                });

                OpenPopUpPageWithDialogOptions(options);

                var frameElement = SP.UI.ModalDialog.get_childDialog().get_frameElement();

                if (frameElement != null) {
                    frameElement = $(frameElement);
                    var loadingMessage = $('<div class="modalLoadingMessage">This shouldn\'t take long...</div>');
                    loadingMessage.css({
                        position: 'absolute',
                        top: frameElement.position().top + 'px',
                        left: frameElement.position().left + 'px',
                        width: frameElement.width(),
                        height: 40,
                        'background-color': 'white'
                    });
                    frameElement.parent().append(loadingMessage);
                }
            };

            if (typeof Porthole !== 'undefined')
                openDialog();
            else
                $.getScript(jQuery.getServerHostUrl() + '/Style Library/CIB/Scripts/lib/porthole.js', openDialog);
        }
    }

}();

jQuery.extend({
    CIBcache: CIB.utilities.CIBcache,
    isAppWeb: CIB.utilities.isAppWeb,
    hasAppWeb: CIB.utilities.hasAppWeb,
    getUrlVars: CIB.utilities.getUrlVars,
    getUrlVar: CIB.utilities.getUrlVar,
    getHostWebUrl: CIB.utilities.getHostWebUrl,
    getServerRealtiveHostWebUrl: CIB.utilities.getServerRealtiveHostWebUrl,
    getServerHostUrl: CIB.utilities.getServerHostUrl,
    getServerAppUrl: CIB.utilities.getServerAppUrl,
    getAppWebUrl: CIB.utilities.getAppWebUrl,
    getServerRealtiveApptWebUrl: CIB.utilities.getServerRealtiveApptWebUrl,
    ensureArray: CIB.utilities.ensureArray,
    handleExceptionsScope: CIB.utilities.handleExceptionsScope,
    resizeAppPart: CIB.utilities.resizeAppPart,
    isAppPart: CIB.utilities.isAppPart,
    isInternetExplorer: CIB.utilities.isInternetExplorer,
    changeAppPage: CIB.utilities.changeAppPage
});

(function ($) {

    $.whenSync = function () {

        var masterDeferred = $.Deferred();
        var masterResults = [];

        var callbacks = Array.prototype.slice.call(arguments);

        if (!callbacks.length) {
            masterDeferred.resolve()
            return (masterDeferred.promise());
        }

        var invokeCallback = function (callback) {

            try {
                callback.apply(this)
                    .then(
                        function () {
                            if (callbacks.length > 0) {
                                invokeCallback(callbacks.shift());
                            }
                            else {
                                masterDeferred.resolve();
                            }
                        },
                        function () {
                            masterDeferred.reject();
                        },
                        function () {
                        }
                    );
            } catch (syncError) {
                masterDeferred.reject(syncError);
                throw syncError;
            }
        }

        invokeCallback(callbacks.shift());

        return (masterDeferred.promise());
    }

})(jQuery);

