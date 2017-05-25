'use strict';

/*

    Logger.js
    Provides a logging framework for javascript errors

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

CIB.logging = function () {

    window.onerror = function (errorMsg, url, lineNumber) {
        CIB.logging.logError('Unhandled JavaScript Error', errorMsg, 'Line: ' + lineNumber + '\r\n' + url);
    };
    var loggingService = function () {
        var rootUrl = $.getServerHostUrl();
        var segments = rootUrl.replace('http://', '').replace('https://', '').split('.');
        
        var serviceUrl = '';

        for (var i = 0; i < segments.length; i++) {
            if (i == 0) {
                if (segments[1] === 'dev')
                    serviceUrl += 'logging.apps';
                else if (segments[0] === 'ihc')
                    serviceUrl += 'logging.sharepoint-ihc';
                else if (segments[0] === 'my-uat') 
                    serviceUrl += 'logging.sharepoint-uat';
                else if (segments[1] === 'staging')
                    serviceUrl += 'logging.sharepoint';
                else 
                    serviceUrl += 'logging.sharepoint';
            }
            else {
                serviceUrl += segments[i];
            }
            if (i < segments.length - 1) {
                serviceUrl += '.';
            }
        }

        serviceUrl += '/Services/Logger.svc/LogError';
        return rootUrl.split(':')[0] + '://' + serviceUrl;
    }();

    var analyticsService = function () {
        var rootUrl = $.getServerHostUrl();
        var segments = rootUrl.replace('http://', '').replace('https://', '').split('.');

        var serviceUrl = '';

        for (var i = 0; i < segments.length; i++) {
            if (i == 0) {
                if (segments[1] === 'dev')
                    serviceUrl += 'logging.apps';
                else if (segments[0] === 'ihc')
                    serviceUrl += 'logging.sharepoint-ihc';
                else if (segments[0] === 'my-uat')
                    serviceUrl += 'logging.sharepoint-uat';
                else if (segments[1] === 'staging')
                    serviceUrl += 'logging.sharepoint';
                else
                    serviceUrl += 'logging.sharepoint';
            }
            else {
                serviceUrl += segments[i];
            }
            if (i < segments.length - 1) {
                serviceUrl += '.';
            }
        }

        serviceUrl += '/Services/Logger.svc/LogEvent';
        return rootUrl.split(':')[0] + '://' + serviceUrl;
    }();

    return {

        applicationName: '',

        logError: function (category, message, stacktrace) {

            $.support.cors = true;

            if (!CIB.logging.applicationName) {
                if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo && _spPageContextInfo.webTitle) {
                    CIB.logging.applicationName = _spPageContextInfo.webTitle;
                }
                else {
                    CIB.logging.applicationName = $.getServerRealtiveHostWebUrl();
                }
            }

            if (console) {
                console.log('Logging error for: ' + CIB.logging.applicationName);
                console.log(category + ': ' + message);
            }

            $.ajax({
                type: 'POST',
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                },
                dataType: 'json',
                contentType: 'application/json',
                url: loggingService,
                data: JSON.stringify({
                    Application: CIB.logging.applicationName,
                    Category: category,
                    Message: message,
                    StackTrace: stacktrace ? stacktrace : window.location.href
                })
            })
            .fail(CIB.logging.logErrorFailed);
        },

        logEvent: function (application, eventType, xmlData) {

            $.support.cors = true;

            if (!CIB.logging.applicationName) {
                if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo && _spPageContextInfo.webTitle) {
                    CIB.logging.applicationName = _spPageContextInfo.webTitle;
                }
                else {
                    CIB.logging.applicationName = $.getServerRealtiveHostWebUrl();
                }
            }

            if (console) {
                console.log('Logging error for: ' + CIB.logging.applicationName);
                console.log('Event type: ' + ': ' + eventType);
                console.log('XML data: ' + ': ' + xmlData);
            }

            $.ajax({
                type: 'POST',
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                },
                dataType: 'json',
                contentType: 'application/json',
                url: analyticsService,
                data: JSON.stringify({
                    Application: application,
                    EventType: eventType,
                    XmlData: xmlData
                })
            })
            .fail(CIB.logging.logErrorFailed);
        },

        logEventAsyncPromise: function (application, eventType, xmlData) {

            var logEventAction = new jQuery.Deferred();

            $.support.cors = true;

            if (!CIB.logging.applicationName) {
                if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo && _spPageContextInfo.webTitle) {
                    CIB.logging.applicationName = _spPageContextInfo.webTitle;
                }
                else {
                    CIB.logging.applicationName = $.getServerRealtiveHostWebUrl();
                }
            }

            if (console) {
                console.log('Logging error for: ' + CIB.logging.applicationName);
                console.log('Event type: ' + ': ' + eventType);
                console.log('XML data: ' + ': ' + xmlData);
            }

            $.ajax({
                type: 'POST',
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                },
                dataType: 'json',
                contentType: 'application/json',
                url: analyticsService,
                data: JSON.stringify({
                    Application: application,
                    EventType: eventType,
                    XmlData: xmlData
                })
            })
            .done(function (content) {
                logEventAction.resolve(content);
            })
            .fail(function (sender, status) {
                CIB.logging.logErrorFailed(sender, status);
                logEventAction.reject(sender.statusText);
            });

            return logEventAction.promise();
        },

        logErrorFailed: function (sender, args) {
            if (console) {
                console.log('Failed to send error to logging service: ' + sender.statusText);
            }
        }
    };
}();
