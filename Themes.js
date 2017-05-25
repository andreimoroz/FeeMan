'use strict';

/*

    Themes.js
    Provides a framework getting the theme of the host web

*/

var CIB = CIB || {};

CIB.themes = function () {

    if (!CIB.utilities)
        throw new Error("Utilities.js must be referenced to use Themes");

    var globalContext = CIB.utilities.getContext();
    var context = globalContext.context;
    var hostContext = globalContext.hostContext;

    var themeLoading = false;
    var themeCached = new jQuery.Deferred();
    var themeCache = new CIB.utilities.CIBcache('CIB.App.NavigationLinks');

    $("<link />", {
        rel: "stylesheet",
        type: "text/css",
        href: $.getServerRealtiveHostWebUrl().replace(new RegExp("[\/]+$"), "") + "/_layouts/15/defaultcss.ashx"
    }).appendTo("head");

    if (themeCache.containsValue())
        themeCached.resolve();

    var failHandler = function (message) {
        if (console && console.log) console.log(message);
        if (CIB.logging) CIB.logging.logError('Themes Helper', message, window.location.href);
        themeLoading = false;
    };

    $(document).ready(function () {

        if (themeCache.containsValue())
            return;
        
        $styleContainer = $('<div />');
        $styleContainer.hide();
        $('body').append($styleContainer);

        var theme = { colours: {}, fonts: {} };

        var getStyle = function (tag, className) {
            var element = $('<' + tag + ' class="' + className + '" />');
            $styleContainer.append(element);
            return document.defaultView.getComputedStyle(element[0], null);
        };

        for (var i = 1; i <= 6; i++) {
            theme.colours['ContentAccent' + i] = getStyle('div', 'ms-ContentAccent' + i + '-fontColor').color;
        }
        
        theme.fonts['title'] = getStyle('div', 'ms-core-pageTitle').fontFamily;
        theme.fonts['navigation'] = getStyle('div', 'ms-core-navigation').fontFamily;
        theme.fonts['large-heading'] = getStyle('h1', 'ms-h1').fontFamily;
        theme.fonts['heading'] = getStyle('h2', 'ms-h2').fontFamily;
        theme.fonts['small-heading'] = getStyle('h4', 'ms-h4').fontFamily;
        theme.fonts['body'] = getStyle('div', 'ms-core-defaultFont').fontFamily;

        themeCache.set(theme);
        themeCached.resolve();

    });

    return {

        load: function () {
            
            var themeLoaded = new jQuery.Deferred();

            themeCached.done(function () {
                themeLoaded.resolve(themeCache.get());
            })
            .fail(function (message) {
                throw new Error(message);
            });

            return themeLoaded.promise();

        }

    }

}();
