"use strict";
var CIB = CIB || {};
window.console && console.log || (console = {
    log: function() {},
    debug: function() {},
    info: function() {},
    warn: function() {},
    error: function() {}
});
CIB.utilities = function() {
    var n, t;
    return $.fn && $.fn.button && $.fn.button.noConflict && ($.fn.bootstrapBtn = $.fn.button.noConflict()), SP.ClientContext.prototype.executeQueryAsyncPromise = function() {
        var n = new jQuery.Deferred;
        return this.executeQueryAsync(function() {
            n.resolve()
        }, function(t, i) {
            var r = i.get_message ? i.get_message() : i.get_errorMessage();
            n.reject(r, t, i)
        }), n.promise()
    }, n = function() {
        var n = $("*").filter(function() {
            var n = $(this).css("position");
            return n === "absolute"
        }).map(function() {
            return $(this).offset().top + $(this).height()
        }).toArray();
        return n.push($("#footerMarker").offset().top), Math.max.apply(this, n)
    }, $(document).ready(function() {
        if ($("body").append('<div id="footerMarker" />'), $.isAppPart()) {
            var t = $(window).width();
            $(window).resize(function() {
                $(window).width() != t && (t = $(window).width(), CIB.utilities.resizeAppPart("100%", n()))
            });
            CIB.utilities.resizeAppPart("100%", n())
        }
    }), t = function(n) {
        var t = window.location.host + "_" + n;
        return {
            get: function() {
                return JSON.parse(sessionStorage[t]).data
            },
            set: function(n) {
                sessionStorage[t] = JSON.stringify({
                    data: n
                })
            },
            invalidate: function() {
                delete sessionStorage[t]
            },
            containsValue: function() {
                return sessionStorage[t] ? !0 : !1
            }
        }
    }, {
        CIBcache: t,
        isAppWeb: function() {
            return $('meta[name="GENERATOR"]').attr("content") === "Microsoft SharePoint" ? !0 : typeof _spPageContextInfo != "undefined" && _spPageContextInfo && jQuery.hasAppWeb()
        },
        hasAppWeb: function() {
            return jQuery.getAppWebUrl() && jQuery.getAppWebUrl() != '""' && jQuery.getAppWebUrl() != "undefined"
        },
        getContext: function() {
            var t = jQuery.isAppWeb() && jQuery.hasAppWeb(),
                n = new SP.ClientContext(t ? jQuery.getAppWebUrl() : jQuery.getHostWebUrl()),
                r, u, i;
            if (t) {
                if (!SP.ProxyWebRequestExecutorFactory) throw new Error("/_layouts/15/sp.requestexecutor.js is not refernced");
                r = new SP.ProxyWebRequestExecutorFactory(jQuery.getAppWebUrl());
                n.set_webRequestExecutorFactory(r)
            }
            return u = t ? new SP.AppContextSite(n, jQuery.getHostWebUrl()) : n, i = {
                context: n,
                hostContext: u
            }, jQuery.hasAppWeb() && (i.appContext = new SP.ClientContext(jQuery.getAppWebUrl())), i
        },
        getUrlVars: function() {
            for (var t = [], n, r = window.location.href.slice(window.location.href.indexOf("?") + 1).split("&"), i = 0; i < r.length; i++) n = r[i].split("="), t.push(n[0]), t[n[0]] = decodeURIComponent(n[1]);
            return t
        },
        getUrlVar: function(n) {
            return jQuery.getUrlVars()[n]
        },
        getHostWebUrl: function() {
            var n, t, i;
            return window.location.href.indexOf("SPHostUrl") > -1 ? (i = $("[id$='CIBAppFrameWorkSubWebUrl']").is("input"), n = i ? $("[id$='CIBAppFrameWorkSubWebUrl']").val() : $("[id$='CIBAppFrameWorkSubWebUrl']").text(), t = n ? n : decodeURIComponent(jQuery.getUrlVar("SPHostUrl")).replace(/\#+$/, "")) : t = decodeURIComponent(window.location.href).replace(/\#+$/, ""), t
        },
        getServerRealtiveHostWebUrl: function() {
            var n = jQuery.getHostWebUrl();
            return n.replace("http://", "").replace("https://", "").indexOf("/") < 0 && (n += "/"), "/" + n.replace(/^(?:\/\/|[^\/]+)*\//, "")
        },
        getServerHostUrl: function() {
            var n = jQuery.getHostWebUrl().match(/^http[s]?:\/\/[^/]+/);
            return n ? n[0] : null
        },
        getServerAppUrl: function() {
            var n = jQuery.getAppWebUrl().match(/^http[s]?:\/\/[^/]+/);
            return n ? n[0] : null
        },
        getAppWebUrl: function() {
            return decodeURIComponent(jQuery.getUrlVar("SPAppWebUrl")).replace(/\#+$/, "")
        },
        getServerRealtiveApptWebUrl: function() {
            var n = jQuery.getAppWebUrl();
            return n.replace("http://", "").replace("https://", "").indexOf("/") < 0 && (n += "/"), "/" + n.replace(/^(?:\/\/|[^\/]+)*\//, "")
        },
        ensureArray: function(n) {
            return $.isArray(n) ? n : [n]
        },
        handleExceptionsScope: function(n, t) {
            var i = new SP.ExceptionHandlingScope(n),
                u = i.startScope(),
                f = i.startTry(),
                r;
            return t(), f.dispose(), r = i.startCatch(), r.dispose(), u.dispose(), i
        },
        isAppPart: function() {
            return jQuery.getUrlVar("SenderId")
        },
        resizeAppPart: function(t, i) {
            jQuery.isAppPart && (t || (t = "100%"), i || (i = n()), window.parent.postMessage("<message senderId=" + jQuery.getUrlVar("SenderId") + ">resize(" + t + ", " + i + ")<\/message>", jQuery.getHostWebUrl()))
        },
        isInternetExplorer: function() {
            return navigator.appVersion.indexOf("MSIE ") >= 0
        },
        serialiseSharePointObject: function(n, t) {
            var r, u, f, i, e;
            t = typeof optionalArg == "undefined" ? !1 : t;
            r = {};
            u = $.map(n, function(n, t) {
                return t
            });
            for (f in u)(i = u[f], i !== "get_id" || t) && i.startsWith("get_") && (e = String(n[i].call(n)), e !== "[object Object]" && (r[i.replace("get_", "")] = n[i].call(n)));
            return JSON.stringify(r)
        },
        deserialiseSharePointObject: function(n, t, i) {
            var u, f, e, r;
            i = typeof optionalArg == "undefined" ? !1 : i;
            u = JSON.parse(n);
            f = $.map(u, function(n, t) {
                return t
            });
            for (e in f)(r = f[e], r !== "id" || i) && t["set_" + r] && t["set_" + r].call(t, u[r])
        },
        changeAppPage: function(n) {
            var t = jQuery.getAppWebUrl().replace(/\/+$/, "");
            t += "/" + n.replace(/^\/+/, "");
            t += window.location.href.slice(window.location.href.indexOf("?"));
            window.location = t
        },
        showHostWebDialog: function(n) {
            if (!n.url) throw new Error("Options must include a url from the host web");
            n.url = jQuery.getServerHostUrl() + "/Style Library/CIB/Pages/Dialog.aspx?target=" + encodeURIComponent(n.url) + "&parent=" + jQuery.getServerAppUrl();
            var t = function() {
                var r = !1,
                    u = new Porthole.WindowProxy(jQuery.getServerHostUrl() + "/Style Library/CIB/Pages/DialogProxy.aspx"),
                    t, i;
                u.addEventListener(function(n) {
                    var i, f;
                    if (n.data.hasOwnProperty("result")) SP.UI.ModalDialog.commonModalDialogClose(n.data.result, n.data.value);
                    else if (n.data.resize) {
                        $(".modalLoadingMessage").remove();
                        SP.UI.ModalDialog.get_childDialog() != null && SP.UI.ModalDialog.get_childDialog().get_frameElement() != null && (SP.UI.ModalDialog.get_childDialog().get_frameElement().parentElement.style["padding-right"] = "0px");
                        var o = n.data.width,
                            e = n.data.height,
                            t = [],
                            u = function(n, t) {
                                n[t] = $(".ms-dlg" + t, window.parent.document)
                            };
                        u(t, "Border");
                        u(t, "Title");
                        u(t, "TitleText");
                        u(t, "Content");
                        u(t, "Frame");
                        $(window).height() - 20 < e && (e = $(window).height() - 20);
                        deltaWidth = o - t.Border.width();
                        deltaHeight = e - t.Border.height();
                        for (i in t) r || t[i].width(t[i].width() + deltaWidth), i != "Title" && i != "TitleText" && t[i].height(t[i].height() + deltaHeight);
                        r = !0;
                        $(".ms-dlgTitle").css("width", "95%");
                        SP.UI.ModalDialog.get_childDialog() != null && SP.UI.ModalDialog.get_childDialog().get_dialogElement() != null && (f = $(SP.UI.ModalDialog.get_childDialog().get_dialogElement()), f.css("top", Math.max(0, ($(window).height() - f.outerHeight()) / 2 + $(window).scrollTop()) + "px"), f.css("left", Math.max(0, ($(window).width() - f.outerWidth()) / 2 + $(window).scrollLeft()) + "px"))
                    }
                });
                OpenPopUpPageWithDialogOptions(n);
                t = SP.UI.ModalDialog.get_childDialog().get_frameElement();
                t != null && (t = $(t), i = $('<div class="modalLoadingMessage">This shouldn\'t take long...<\/div>'), i.css({
                    position: "absolute",
                    top: t.position().top + "px",
                    left: t.position().left + "px",
                    width: t.width(),
                    height: 40,
                    "background-color": "white"
                }), t.parent().append(i))
            };
            typeof Porthole != "undefined" ? t() : $.getScript(jQuery.getServerHostUrl() + "/Style Library/CIB/Scripts/lib/porthole.js", t)
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
    }),
    function(n) {
        n.whenSync = function() {
            var t = n.Deferred(),
                i = Array.prototype.slice.call(arguments),
                r;
            return i.length ? (r = function(n) {
                try {
                    n.apply(this).then(function() {
                        i.length > 0 ? r(i.shift()) : t.resolve()
                    }, function() {
                        t.reject()
                    }, function() {})
                } catch (u) {
                    t.reject(u);
                    throw u;
                }
            }, r(i.shift()), t.promise()) : (t.resolve(), t.promise())
        }
    }(jQuery)