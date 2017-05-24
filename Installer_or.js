"use strict";
var CIB = CIB || {};
CIB.installer = function() {
        var f, t, i, r = {},
            u = {},
            n;
        return $(document).ready(function() {
            if (f = CIB.utilities.getContext(), t = f.context, i = f.hostContext, !$.isInternetExplorer() && !$.hasAppWeb())
                if ($("#install-status").length > 0) $("#install-status").append($('<div class="alert alert-danger" role="alert" style="width:580px;"><strong>Unsupported browser<\/strong><span>The provisioning wizard will only work with internet explorer for provider hosted apps.<\/span><\/div>'));
                else throw new Error("Unsupported browser: The provisioning wizard will only work with internet explorer for provider hosted apps.");
        }), n = function() {
            return {
                executeQuery: function(i, r, u) {
                    t.executeQueryAsync(function() {
                        var t = [],
                            f = !0;
                        $.each(i, function() {
                            var i = this,
                                r;
                            i.get_hasException() ? (r = n.handleError(this, i), f &= r.handled, t.push(r.message)) : (n.message(i.successMessage, "success"), t.push(i.successMessage))
                        });
                        f ? u ? r.resolve(t, u) : r.resolve(t) : r.reject(t)
                    }, function(t, i) {
                        var u = n.handleError(t, i);
                        u.handled ? r.resolve(u.message) : r.reject(u.message)
                    })
                },
                handleError: function(t, i) {
                    for (var r = i.get_message ? i.get_message() : i.get_errorMessage(), f = [" is already activated at scope ", 'A duplicate field name "', 'A duplicate content type "', "A file or folder with the name ", "The specified name is already in use.", "A list, survey, discussion board, or document library with the specified title already exists in this Web site."], e = "error", u = 0; u < f.length; u++)
                        if (r.slice(0, f[u].length) == f[u] || r.indexOf(f[u]) > -1) {
                            e = "info";
                            r += " (expected if provisioned already)";
                            break
                        }
                    return n.message(r, e), {
                        handled: e == "info",
                        message: r
                    }
                },
                message: function(n, t) {
                    var r, i;
                    t || (t = "message");
                    console && console.log && console.log(n + " [" + t + "]");
                    t == "error" && CIB.logging.logError("Provisioning", n, window.location.href);
                    r = t == "success" ? "green" : t == "error" ? "red" : t == "info" ? "orange" : "gray";
                    $("#install-status").append('<span style="color:' + r + '">' + n + "<\/span>");
                    i = document.getElementById("install-status");
                    i && (i.scrollTop = i.scrollHeight)
                },
                updateListIds: function() {
                    var u = new jQuery.Deferred,
                        f = i.get_web().get_lists();
                    return t.load(f, "Include(Title, Id)"), t.executeQueryAsync(function() {
                        for (var t = f.getEnumerator(), n; t.moveNext();) n = t.get_current(), r[n.get_title()] = n.get_id();
                        u.resolve()
                    }, function(t, i) {
                        var r = n.handleError(t, i);
                        u.reject(r.message)
                    }), u.promise()
                },
                getViewsForList: function(r, f) {
                    var e = new $.Deferred,
                        s = i.get_web(),
                        h = s.get_lists().getByTitle(f),
                        o = h.get_views();
                    t.load(o, "Include(Id, Title)");
                    t.executeQueryAsync(function() {
                        for (var n = o.getEnumerator(); n.moveNext();) {
                            var t = n.get_current(),
                                i = t.get_title(),
                                r = t.get_id().toString().toLowerCase();
                            $.isEmptyObject(u[f]) && (u[f] = []);
                            u[f][i] = r
                        }
                        e.resolve()
                    }, function(t, i) {
                        n.handleError(t, i);
                        e.fail()
                    });
                    r.push(e)
                },
                updateViewIds: function() {
                    var t = [];
                    for (var i in r) n.getViewsForList(t, i);
                    return $.when.apply($, t).promise()
                }
            }
        }(), window.onerror = function(t, i, r) {
            CIB.logging.logError("Unhandled JavaScript Error", t, "Line: " + r + "\r\n" + i);
            n.message(t, "error")
        }, {
            message: function(t, i) {
                n.message(t, i)
            },
            getListIds: function() {
                var t = new jQuery.Deferred;
                return n.updateListIds().done(function() {
                    t.resolve(r)
                }).fail(function(n) {
                    t.reject(n)
                }), t.promise()
            },
            activateFeatures: function(r) {
                var u = [],
                    r = CIB.utilities.ensureArray(r),
                    f = new jQuery.Deferred;
                return $.each(r, function() {
                    var r = this,
                        f;
                    if (!r.id || !r.name || !r.scope) throw new Error("Feature object must had id, name and scope attributes");
                    if (r.scope != "site" && r.scope != "web") throw new Error("Feature scope must be either site or web");
                    f = $.handleExceptionsScope(t, function() {
                        n.message("Activateg feature '" + r.name + "'");
                        var t = r.scope == "site" ? i.get_site().get_features() : i.get_web().get_features(),
                            u = t.add(new SP.Guid(r.id), !1, SP.FeatureDefinitionScope.farm)
                    });
                    f.successMessage = "Feature '" + r.name + "' activated.";
                    u.push(f)
                }), n.executeQuery(u, f), f.promise()
            },
            createLists: function(r) {
                var u = [],
                    r = CIB.utilities.ensureArray(r),
                    f = new jQuery.Deferred;
                return $.each(r, function() {
                    var r = this,
                        f;
                    if (!r.name || !r.type) throw new Error("List object must have name and type attributes");
                    f = $.handleExceptionsScope(t, function() {
                        var o, u, t, e, f;
                        n.message("Creating list '" + r.name + "'");
                        o = i.get_web().get_lists();
                        t = new SP.ListCreationInformation;
                        t.set_title(r.name);
                        t.set_templateType(r.type);
                        r.feature && t.set_templateFeatureId(r.feature);
                        u = !1;
                        t = o.add(t);
                        (r.type == "10002" || r.type == "10000" || r.type == "10001") && (e = t.get_rootFolder(), f = e.get_properties(), r.type == "10002" ? f.set_item("InformationSecurityLevel", 0) : r.type == "10001" ? f.set_item("InformationSecurityLevel", 2) : r.type == "10000" && f.set_item("InformationSecurityLevel", 1), e.update());
                        r.hasOwnProperty("hidden") && (t.set_hidden(r.hidden), u = !0);
                        r.hasOwnProperty("onQuickLaunch") && (t.set_onQuickLaunch(r.onQuickLaunch), u = !0);
                        u && t.update()
                    });
                    f.successMessage = "List '" + r.name + "' created.";
                    u.push(f)
                }), n.executeQuery(u, f), f.promise()
            },
            createFolders: function(r) {
                var u = [],
                    r = CIB.utilities.ensureArray(r),
                    f = new jQuery.Deferred;
                return $.each(r, function() {
                    var r = this,
                        f;
                    if (!r.name || !r.list || !r.path) throw new Error("Folder object must have name, list and path attributes");
                    f = $.handleExceptionsScope(t, function() {
                        var f, t, u;
                        n.message("Creating folder " + r.name + " in list " + r.list);
                        f = i.get_web().get_lists().getByTitle(r.list);
                        t = new SP.ListItemCreationInformation;
                        t.set_underlyingObjectType(SP.FileSystemObjectType.folder);
                        t.set_leafName(r.name);
                        t.set_folderUrl($.getHostWebUrl() + "/" + r.path);
                        u = f.addItem(t);
                        u.set_item("Title", r.name);
                        u.update()
                    });
                    f.successMessage = "Folder " + r.name + " created in list " + r.list;
                    u.push(f)
                }), n.executeQuery(u, f), f.promise()
            },
            updateFileTokens: function(n) {
                var t = function(n, t) {
                        if (n[t]) return n[t]
                    },
                    i = function(n, t) {
                        if (u[n]) return u[n][t]
                    },
                    f = function(n) {
                        var u = new RegExp("{\\$([^:]*):([^}]*)}", "gi");
                        return n = n.replace(u, function(n, i, u) {
                            return i === "List" ? t(r, u) : n
                        }), u = new RegExp("{\\$([^:]*):([^}]*):([^}]*)}", "gi"), n.replace(u, function(n, t, r, u) {
                            return t === "ListView" ? i(r, u) : n
                        })
                    };
                return f(n)
            },
            copyFiles: function(r) {
                var r = CIB.utilities.ensureArray(r),
                    f = 0,
                    u = new jQuery.Deferred;
                return $.each(r, function() {
                    var o = [],
                        e = this;
                    if (!e.name || !e.url || !e.path) throw new Error("File object must have name, url and path attributes");
                    n.message("Copying file " + e.name);
                    var c = function(n) {
                            var t = n.binary ? !0 : !1,
                                i = "_api/web/GetFileByServerRelativeUrl('" + ($.getServerRealtiveApptWebUrl() + "/" + n.path).replace("//", "/") + "')/$value",
                                r = new SP.RequestExecutor($.getAppWebUrl()),
                                u = {
                                    url: i,
                                    method: "GET",
                                    binaryStringResponseBody: t,
                                    success: function(t) {
                                        h(n, t.body)
                                    },
                                    error: function() {
                                        s(n)
                                    }
                                };
                            r.executeAsync(u)
                        },
                        s = function(n) {
                            $.support.cors = !0;
                            var t = $.getAppWebUrl() + "/_api/SP.AppContextSite(@target)/web/GetFileByServerRelativeUrl('" + ($.getServerRealtiveApptWebUrl() + "/" + n.path).replace("//", "/") + "')/$value?@target='" + $.getHostWebUrl() + "'";
                            $.ajax({
                                url: t,
                                cache: !1
                            }).done(function(t) {
                                h(n, t)
                            }).fail(function(n) {
                                u.reject(n.statusText)
                            })
                        },
                        h = function(e, s) {
                            e.name != "Installer.js" && (s = s.replace(/{#HostWebURL#}/g, $.getHostWebUrl()), s = s.replace(/{#ServerRelativeHostWebURL#}/g, $.getServerRealtiveHostWebUrl()), s = CIB.installer.updateFileTokens(s));
                            var h = ($.getServerRealtiveHostWebUrl() + "/" + e.url).replace("//", "/"),
                                c = $.handleExceptionsScope(t, function() {
                                    var n, r, f, u;
                                    for (e.publish && $.handleExceptionsScope(t, function() {
                                            var n = i.get_web().getFileByServerRelativeUrl(h + "/" + e.name);
                                            n.checkOut()
                                        }), n = new SP.FileCreationInformation, n.set_content(new SP.Base64EncodedByteArray), r = 0; r < s.length; r++) n.get_content().append(s.charCodeAt(r));
                                    n.set_overwrite(!0);
                                    n.set_url(e.name);
                                    f = i.get_web().getFolderByServerRelativeUrl(h).get_files();
                                    u = f.add(n);
                                    e.publish && (u.checkIn("Checked in by provisioning framework.", SP.CheckinType.majorCheckIn), u.publish("Published by provisioning framework."))
                                });
                            c.successMessage = "File " + e.name + " created at " + e.url;
                            o.push(c);
                            ++f == r.length && n.executeQuery(o, u)
                        };
                    $.isAppWeb() ? c(e) : s(e)
                }), u.promise()
            },
            updateListIds: function() {
                return n.updateListIds()
            },
            updateViewIds: function() {
                return n.updateViewIds()
            },
            createSiteColumns: function(n) {
                var t = i.get_web().get_fields();
                return CIB.installer.createColumns(n, t)
            },
            createListColumns: function(n, t) {
                var r = i.get_web().get_lists().getByTitle(n),
                    u = r.get_fields();
                return CIB.installer.createColumns(t, u)
            },
            createColumns: function(i, u) {
                var o = [],
                    i = CIB.utilities.ensureArray(i),
                    f = new jQuery.Deferred,
                    e;
                if (!u) throw new Error("Field collection not provided, use createSiteColumns or createListColumns instead.");
                return e = function() {
                    $.each(i, function() {
                        var i = this,
                            e;
                        if (!i.id || !i.name || !i.type || !i.displayName || !i.group) throw new Error("Column object must have id, name, type, group and displayName attributes");
                        e = $.handleExceptionsScope(t, function() {
                            var b, e, s, o, a, l, h, v, y;
                            n.message("Creating column '" + i.displayName + "'");
                            var p = i.hidden ? "true" : "false",
                                w = i.required ? "true" : "false",
                                k = i.multi ? "true" : "false",
                                c = "<Field ID='" + i.id + "' Type='" + i.type + "' DisplayName='" + i.name + "' Name='" + i.name + "' Group='" + i.group + "' Required='" + w + "' />";
                            if (i.type.toLowerCase() == "calculated") {
                                if (!i.formula || !i.resultType) throw new Error("Calculated columns must have a formula and resultType set");
                                b = "<Formula>" + i.formula + "<\/Formula>";
                                c = c.replace(" />", ' ResultType="' + i.resultType + '">' + b + "<\/Field>")
                            }
                            if (k == "true" && (c = c.replace(" />", ' Mult="TRUE" />')), e = u.addFieldAsXml(c, !1, SP.AddFieldOptions.AddToNoContentType), p != null && e.set_hidden(p), e.set_title(i.displayName), e.set_required(w), i.defaultValue && e.set_defaultValue(i.defaultValue), t.load(e), i.type.toLowerCase() == "lookup") {
                                if (!r[i.lookupList]) {
                                    s = "The id for the list " + i.lookupList + " has not been loaded. updateListIds must be called before creating lookup fields";
                                    f.reject(s);
                                    throw new Error(s);
                                }
                                o = t.castTo(e, SP.FieldLookup);
                                o.set_lookupList(r[i.lookupList]);
                                o.set_lookupField(i.lookupField);
                                o.update();
                                i.additionalFields && $.each(i.additionalFields, function() {
                                    var n = this;
                                    u.addDependentLookup(n.displayName, e, n.target)
                                })
                            } else if (i.type.toLowerCase() == "lookupmulti") {
                                if (!r[i.lookupList]) {
                                    s = "The id for the list " + i.lookupList + " has not been loaded. updateListIds must be called before creating lookup fields";
                                    f.reject(s);
                                    throw new Error(s);
                                }
                                o = t.castTo(e, SP.FieldLookup);
                                o.set_lookupList(r[i.lookupList]);
                                o.set_lookupField(i.lookupField);
                                o.set_allowMultipleValues(!0);
                                o.update();
                                i.additionalFields && $.each(i.additionalFields, function() {
                                    var n = this;
                                    u.addDependentLookup(n.displayName, e, n.target)
                                })
                            } else i.type.toLowerCase() == "currency" && i.locale ? (a = t.castTo(e, SP.FieldCurrency), a.set_currencyLocaleId(i.locale), a.update()) : i.type.toLowerCase() == "number" ? (l = t.castTo(e, SP.FieldNumber), i.minimumValue && l.set_minimumValue(i.minimumValue), i.maximumValue && l.set_maximumValue(i.maximumValue), l.update()) : i.type.toLowerCase() == "choice" && i.choices ? (h = t.castTo(e, SP.FieldChoice), h.set_choices($.makeArray(i.choices)), h.update()) : i.type.toLowerCase() == "multichoice" && i.choices ? (h = t.castTo(e, SP.FieldMultiChoice), h.set_choices($.makeArray(i.choices)), h.update()) : i.type.toLowerCase() == "datetime" && i.dateOnly ? (v = t.castTo(e, SP.FieldDateTime), v.set_displayFormat(SP.DateTimeFieldFormatType.dateOnly), v.update()) : i.type.toLowerCase() == "taxonomyfieldtypemulti" ? (y = t.castTo(e, SP.Taxonomy.TaxonomyField), y.set_allowMultipleValues(!0), y.update()) : e.update()
                        });
                        e.successMessage = "Column " + i.displayName + " created";
                        o.push(e)
                    });
                    n.executeQuery(o, f)
                }, i.filter(function(n) {
                    return n.type == "lookup"
                }).length > 0 ? n.updateListIds().done(e) : e(), f.promise()
            },
            createContentTypes: function(r) {
                var u = [],
                    r = CIB.utilities.ensureArray(r),
                    f = new jQuery.Deferred;
                return $.each(r, function() {
                    var r = this,
                        f;
                    if (!r.name || !r.id || !r.group) throw new Error("Content Type object must have id, name and group attributes");
                    f = $.handleExceptionsScope(t, function() {
                        n.message("Creating content type '" + r.name + "'");
                        var u = i.get_web().get_contentTypes(),
                            t = new SP.ContentTypeCreationInformation;
                        t.set_id(r.id);
                        t.set_name(r.name);
                        t.set_group(r.group);
                        u.add(t)
                    });
                    f.successMessage = "Content type " + r.name + " created";
                    u.push(f)
                }), n.executeQuery(u, f), f.promise()
            },
            addColumnsToContentType: function(r, u) {
                var u = CIB.utilities.ensureArray(u),
                    f = new jQuery.Deferred,
                    s = i.get_web().get_fields(),
                    h = i.get_web().get_contentTypes(),
                    o = h.getById(r),
                    e = o.get_fieldLinks();
                return t.load(e), t.executeQueryAsync(function() {
                    for (var r = [], h = [], l = [], c = e.getEnumerator(), i; c.moveNext();) i = c.get_current(), l.push(i.get_id().toString().toLowerCase()), h.push(i.get_name());
                    $.each(u, function() {
                        var i = this.toString(),
                            u;
                        if ($.inArray(i, h) >= 0) {
                            n.message("Column already added to content type '" + i + "'. (expected if provisioned already)", "info");
                            return
                        }
                        u = $.handleExceptionsScope(t, function() {
                            var t = s.getByInternalNameOrTitle(i),
                                n = new SP.FieldLinkCreationInformation,
                                r = n.set_field(t);
                            e.add(n)
                        });
                        u.successMessage = "Added column " + i + " to content type";
                        r.push(u)
                    });
                    o.update(!0);
                    n.executeQuery(r, f)
                }, function(t, i) {
                    var r = n.handleError(t, i);
                    r.handled ? f.resolve(r.message) : f.reject(r.message)
                }), f.promise()
            },
            hideColumnsFromEditForm: function(r, u) {
                var e = [],
                    u = CIB.utilities.ensureArray(u),
                    f = new jQuery.Deferred;
                n.message("Hiding columns in list '" + r + "'");
                var s = i.get_web(),
                    h = s.get_lists().getByTitle(r),
                    o = h.get_fields();
                return t.load(o), t.executeQueryAsync(function() {
                    $.each(u, function() {
                        var n = this,
                            i = $.handleExceptionsScope(t, function() {
                                var t = o.getByInternalNameOrTitle(n);
                                t.setShowInEditForm(!1);
                                t.update()
                            });
                        i.successMessage = 'Column "' + n + '" hidden from edit view';
                        e.push(i)
                    })
                }, function(t, i) {
                    var r = n.handleError(t, i);
                    r.handled ? f.resolve(r.message) : f.reject(r.message)
                }), n.executeQuery(e, f), f.promise()
            },
            createView: function(r, u, f, e, o, s, h) {
                var a = [],
                    c = new jQuery.Deferred;
                n.message("Creating view " + u + " for list '" + r + "'");
                var v = i.get_web(),
                    y = v.get_lists().getByTitle(r),
                    l = y.get_views(),
                    f = $.ensureArray(f);
                return t.load(l, "Include(Title, ViewFields)"), t.executeQueryAsync(function() {
                    for (var i, p = l.getEnumerator(), v, y; p.moveNext();)
                        if (v = p.get_current(), u == v.get_title()) {
                            n.message("View '" + u + "' already exists for list " + r + ".", "info");
                            i = v;
                            break
                        }
                    y = $.handleExceptionsScope(t, function() {
                        var n;
                        if (i) {
                            for (var t = [], r = i.get_viewFields(), c = r.getEnumerator(); c.moveNext();) t.push(c.get_current());
                            f.forEach(function(n) {
                                t.indexOf(n) < 0 && r.add(n)
                            });
                            e && i.set_viewQuery(e);
                            s && i.set_rowLimit(s);
                            h != undefined && h != null && i.set_paged(h);
                            i.update()
                        } else n = new SP.ViewCreationInformation, n.set_title(u), n.set_viewFields(f), n.set_query(e), s && n.set_rowLimit(parseInt(s)), o && n.set_viewTypeKind(o), h != undefined && h != null && n.set_paged(h), l.add(n)
                    });
                    y.successMessage = u + (i ? " updated" : " created ") + "for list '" + r + "'";
                    a.push(y);
                    n.executeQuery(a, c)
                }, function(t, i) {
                    var r = n.handleError(t, i);
                    r.handled ? c.resolve(r.message) : c.reject(r.message)
                }), c.promise()
            },
            addContentTypesToList: function(r, u) {
                var o = [],
                    u = CIB.utilities.ensureArray(u),
                    s = new jQuery.Deferred,
                    f, e, h, c;
                return n.message("Adding content types to list '" + r + "'"), f = i.get_web(), e = f.get_lists().getByTitle(r), e.set_contentTypesEnabled(!0), h = f.get_contentTypes(), c = e.get_contentTypes(), $.each(u, function() {
                    var n = this,
                        i = $.handleExceptionsScope(t, function() {
                            var t = h.getById(n);
                            c.addExistingContentType(t)
                        });
                    i.successMessage = "Content type " + n + " added to list";
                    o.push(i)
                }), n.executeQuery(o, s), s.promise()
            },
            removeContentTypesFromList: function(r, u) {
                var s = [],
                    u = CIB.utilities.ensureArray(u),
                    f = new jQuery.Deferred,
                    h, e, o;
                return n.message("Removing content types from list '" + r + "'"), h = i.get_web(), e = h.get_lists().getByTitle(r), e.set_contentTypesEnabled(!0), o = e.get_contentTypes(), t.load(o, "Include(Id, Name)"), t.executeQueryAsync(function() {
                    $.each(u, function() {
                        for (var i = this, e = !1, h = o.getEnumerator(), u, f; h.moveNext();)
                            if (u = h.get_current(), u.get_name().toLowerCase() == i.toLowerCase()) {
                                f = $.handleExceptionsScope(t, function() {
                                    u.deleteObject()
                                });
                                f.successMessage = "Content type " + i + " removed from list";
                                s.push(f);
                                e = !0;
                                break
                            }
                        e || n.message("Could not find '" + i + "' in list " + r + ".", "info")
                    });
                    n.executeQuery(s, f)
                }, function(t, i) {
                    var r = n.handleError(t, i);
                    r.handled ? f.resolve(r.message) : f.reject(r.message)
                }), f.promise()
            },
            setDefaultContentType: function(r, u) {
                var f = new jQuery.Deferred;
                n.message("Setting default content type on list '" + r + "' to " + u);
                var h = i.get_web(),
                    o = h.get_lists().getByTitle(r),
                    s = o.get_contentTypes(),
                    e = o.get_rootFolder();
                return t.load(e, "ContentTypeOrder", "UniqueContentTypeOrder"), t.load(s, "Include(Id, Name)"), t.executeQueryAsync(function() {
                    var i = [],
                        o = $.handleExceptionsScope(t, function() {
                            for (var i = [], r = s.getEnumerator(), f = e.get_contentTypeOrder(), n, t; r.moveNext();)
                                if (n = r.get_current(), n.get_name().toLowerCase() != "folder") {
                                    if (n.get_name().toLowerCase() == u.toLowerCase()) {
                                        i.splice(0, 0, n.get_id());
                                        continue
                                    }
                                    for (t = 0; t < f.length; t++)
                                        if (f[t].toString() == n.get_id()) {
                                            i.push(n.get_id());
                                            break
                                        }
                                }
                            e.set_uniqueContentTypeOrder(i);
                            e.update()
                        });
                    o.successMessage = "Default content type set on list '" + r + "' to " + u;
                    i.push(o);
                    n.executeQuery(i, f)
                }, function(t, i) {
                    var r = n.handleError(t, i);
                    r.handled ? f.resolve(r.message) : f.reject(r.message)
                }), f.promise()
            },
            addIndiciesToList: function(r, u) {
                var f = [],
                    u = CIB.utilities.ensureArray(u),
                    e = new jQuery.Deferred;
                n.message("Setting indicies on list '" + r + "'");
                var s = i.get_web(),
                    o = s.get_lists().getByTitle(r),
                    h = o.get_fields();
                return $.each(u, function() {
                    var n = this,
                        i = $.handleExceptionsScope(t, function() {
                            var t = h.getByInternalNameOrTitle(n);
                            t.set_indexed(!0);
                            t.update();
                            o.update()
                        });
                    i.successMessage = "Index created on column " + n + " in list '" + r + "'";
                    f.push(i)
                }), n.executeQuery(f, e), e.promise()
            },
            enforceUniqueValues: function(r, u) {
                var f = [],
                    u = CIB.utilities.ensureArray(u),
                    e = new jQuery.Deferred;
                n.message("Enforcing unique values on list '" + r + "'");
                var s = i.get_web(),
                    o = s.get_lists().getByTitle(r),
                    h = o.get_fields();
                return $.each(u, function() {
                    var n = this,
                        i = $.handleExceptionsScope(t, function() {
                            var t = h.getByInternalNameOrTitle(n);
                            t.set_indexed(!0);
                            t.set_enforceUniqueValues(!0);
                            t.update();
                            o.update()
                        });
                    i.successMessage = "Enforced unique values on column " + n + " in list '" + r + "'";
                    f.push(i)
                }), n.executeQuery(f, e), e.promise()
            },
            addListViewWebPartToPage: function(r, u, f, e, o, s) {
                var h = new jQuery.Deferred;
                return CIB.installer.addWebPartsToPage({
                    url: r,
                    title: e,
                    assembly: "Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
                    type: "Microsoft.SharePoint.WebPartPages.XsltListViewWebPart",
                    zone: o,
                    index: s,
                    properties: '<property name="ListUrl" type="string">' + u + "<\/property>"
                }).done(function(r, e) {
                    var a = e.get_id(),
                        c = i.get_web().get_lists().getByTitle(u.replace("Lists/", "")),
                        o = c.get_views().getById(a),
                        s = c.get_views().getByTitle(f),
                        l = s.get_viewFields();
                    t.load(o);
                    t.load(s);
                    t.load(l);
                    t.executeQueryAsyncPromise().done(function() {
                        var i, r;
                        for (o.set_viewData(s.get_viewData()), o.set_viewJoins(s.get_viewJoins()), o.set_viewProjectedFields(s.get_viewProjectedFields), o.set_viewQuery(s.get_viewQuery()), o.get_viewFields().removeAll(), i = l.getEnumerator(); i.moveNext();) r = i.get_current(), o.get_viewFields().add(r);
                        o.update();
                        t.executeQueryAsyncPromise().done(function() {
                            n.message("Web part view updated to match " + f, "success");
                            h.resolve()
                        }).fail(function(t) {
                            n.message(t, "error");
                            h.reject(t)
                        })
                    }).fail(function(t) {
                        n.message(t, "error");
                        h.reject(t)
                    })
                }), h.promise()
            },
            addWebPartsToPage: function(r) {
                var f = [],
                    r = CIB.utilities.ensureArray(r),
                    e = i.get_web(),
                    u = new jQuery.Deferred;
                return $.each(r, function() {
                    var i = this;
                    if (!i.url || !i.title || !i.assembly || !i.type || !i.zone || !i.index) throw new Error("Web part object must have url, title, assembly, type, zone and index attributes");
                    n.message("Adding webpart '" + i.title + "' to file " + i.url + ".");
                    var s = e.getFileByServerRelativeUrl(($.getServerRealtiveHostWebUrl() + "/" + i.url).replace("//", "/")),
                        r = s.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared),
                        o = r.get_webParts();
                    t.load(o, "Include(WebPart.Title)");
                    t.executeQueryAsync(function() {
                        for (var h = [], c = o.getEnumerator(), l, a, e, s; c.moveNext();) l = c.get_current().get_webPart(), h.push(l.get_title());
                        if ($.inArray(i.title, h) < 0) s = $.handleExceptionsScope(t, function() {
                            var n = '<?xml version="1.0" encoding="utf-8"?><webParts><webPart xmlns="http://schemas.microsoft.com/WebPart/v3"><metaData><type name="' + i.type + ", " + i.assembly + '" /><importErrorMessage>Cannot import this Web Part.<\/importErrorMessage><\/metaData><data><properties><property name="Title" type="string">' + i.title + '<\/property><property name="ChromeType" type="chrometype">None<\/property>' + (i.properties ? i.properties : "") + "<\/properties><\/data><\/webPart><\/webParts>",
                                u = r.importWebPart(n);
                            a = u.get_webPart();
                            e = r.addWebPart(a, i.zone, i.index);
                            t.load(e)
                        }), s.successMessage = "Webpart '" + i.title + "' added to file " + i.url + ".", f.push(s), n.executeQuery(f, u, e);
                        else {
                            n.message("Webpart '" + i.title + "' already exists in file " + i.url + ".", "info");
                            u.resolve();
                            return
                        }
                    }, function(t, i) {
                        var r = n.handleError(t, i);
                        r.handled || u.reject(r.message)
                    })
                }), u.promise()
            },
            createGroup: function(r) {
                var u = [],
                    r = CIB.utilities.ensureArray(r),
                    f = new jQuery.Deferred,
                    e = i.get_web(),
                    o = e.get_siteGroups();
                return $.each(r, function() {
                    var i = this,
                        r;
                    if (!i.title || !i.description) throw new Error("Group object must have title and description attributes");
                    n.message("Creating group '" + i.title + "'.");
                    r = $.handleExceptionsScope(t, function() {
                        var n = new SP.GroupCreationInformation;
                        n.set_title(i.title);
                        n.set_description(i.description);
                        o.add(n)
                    });
                    r.successMessage = "Created group '" + i.title + "'.";
                    u.push(r)
                }), n.executeQuery(u, f), f.promise()
            },
            registerRemoteEventReceivers: function(t) {
                var i = new jQuery.Deferred,
                    r;
                return n.message("Registering event services at: " + t), r = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">                     <soap:Body>                         <Install xmlns="http://tempuri.org/">                             <hostWebUrl>' + $.getHostWebUrl() + "<\/hostWebUrl>                             <serviceUrl>" + t + "<\/serviceUrl>                         <\/Install>                     <\/soap:Body>                 <\/soap:Envelope>", $.ajax({
                    url: t,
                    type: "POST",
                    dataType: "xml",
                    data: r,
                    contentType: "text/xml",
                    beforeSend: function(n) {
                        n.setRequestHeader("SOAPAction", "http://tempuri.org/IInstallableEventService/Install")
                    },
                    success: function() {
                        n.message("Event services at: " + t + " registered.", "success");
                        i.resolve()
                    },
                    error: function(r, u, f) {
                        n.message("Failed to register event receivers at: " + t + " (" + f + ").", "error");
                        i.reject(f)
                    }
                }), i.promise()
            },
            installWorkflowFromFile: function(t) {
                var i = new jQuery.Deferred,
                    r = new jQuery.Deferred;
                return SP.WorkflowServices ? r.resolve() : $.getScript($.getHostWebUrl() + "/_layouts/15/SP.WorkflowServices.js").fail(function(t) {
                    n.message("Failed to load SP.Workflow.js or a depdency", "error");
                    r.reject(t)
                }).done(function() {
                    SP.WorkflowServices ? r.resolve() : (n.message("Failed to load SP.Workflow.js or a depdency", "error"), r.reject(error))
                }), r.promise().done(function() {
                    $.get(t).fail(function(r) {
                        n.message("Failed to get workflow data from url: " + t, "error");
                        i.reject(r)
                    }).done(function(r) {
                        if (typeof r == "string") try {
                            r = JSON.parse(r)
                        } catch (u) {
                            n.message("Failed to parse workflow from url: " + t, "error");
                            i.reject(u);
                            return
                        }
                        CIB.installer.installWorkflow(r).fail(function(r) {
                            n.message("Failed to install workflow from url: " + t, "error");
                            i.reject(r)
                        }).done(function() {
                            i.resolve()
                        })
                    })
                }), i.promise()
            },
            installWorkflow: function(t) {
                var l = new jQuery.Deferred;
                if ($.isInternetExplorer() || n.message("The installWorkflow method is only supported in internet explorer, the method will run but errors may occur.", "info"), !t.definition || !t.associations) throw new Error('Workflow data must have "definition" and "associations" properties set');
                if (!t.definition.displayName || !t.definition.xaml) throw new Error('Workflow definition must have at least "displayName" and "xaml" properties set');
                var f = t.definition,
                    a = CIB.utilities.ensureArray(t.associations),
                    v = CIB.utilities.ensureArray(t.collateral);
                n.message("Creating workflow definition '" + f.displayName + "'");
                var i = new SP.ClientContext(jQuery.getHostWebUrl()),
                    s = i.get_web(),
                    p = i.get_site(),
                    h, e, o, c, u = function(t) {
                        n.message(t, "error");
                        l.reject(t)
                    };
                h = new SP.WorkflowServices.WorkflowServicesManager.newObject(i, i.get_web());
                i.load(s, "Id", "Url", "ServerRelativeUrl");
                i.load(p, "Id");
                i.load(h);
                i.executeQueryAsyncPromise().fail(u).done(function() {
                    e = h.getWorkflowDeploymentService();
                    o = h.getWorkflowSubscriptionService();
                    i.load(e);
                    i.load(o);
                    $.when(CIB.installer.getListIds(), i.executeQueryAsyncPromise()).fail(u).done(w)
                });
                var w = function(t) {
                        var r = e.enumerateDefinitions(!1);
                        i.load(r, "Include(DisplayName, Id)");
                        i.executeQueryAsyncPromise().fail(u).done(function() {
                            for (var s = new jQuery.Deferred, o, c = r.getEnumerator(), h; c.moveNext();)
                                if (h = c.get_current(), h.get_displayName() === f.displayName) {
                                    n.message('Workflow "' + f.displayName + '" already exsists, it will be overwritten');
                                    o = h;
                                    s.resolve();
                                    break
                                }
                            o || (o = new SP.WorkflowServices.WorkflowDefinition.newObject(i, i.get_web()), o.set_displayName(f.displayName), o.set_xaml('<Activity mc:Ignorable="mwaw" x:Class="Workflow deployment in progress.MTW" xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities" xmlns:local="clr-namespace:Microsoft.SharePoint.WorkflowServices.Activities" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mwaw="clr-namespace:Microsoft.Web.Authoring.Workflow;assembly=Microsoft.Web.Authoring" xmlns:scg="clr-namespace:System.Collections.Generic;assembly=mscorlib" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Flowchart StartNode="{x:Reference __ReferenceID0}"><FlowStep x:Name="__ReferenceID0"><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments="x:String, x:String"><x:String x:Key="Next">4294967294<\/x:String><\/scg:Dictionary><\/mwaw:SPDesignerXamlWriter.CustomAttributes><Sequence><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments="x:String, x:String"><x:String x:Key="StageAttribute">StageContainer-8EDBFE6D-DA0D-42F6-A806-F5807380DA4D<\/x:String><\/scg:Dictionary><\/mwaw:SPDesignerXamlWriter.CustomAttributes><local:SetWorkflowStatus Disabled="False" Status="Stage 1"><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments="x:String, x:String"><x:String x:Key="StageAttribute">StageHeader-7FE15537-DFDB-4198-ABFA-8AF8B9D669AE<\/x:String><\/scg:Dictionary><\/mwaw:SPDesignerXamlWriter.CustomAttributes><\/local:SetWorkflowStatus><Sequence DisplayName="Stage 1" /><Sequence><mwaw:SPDesignerXamlWriter.CustomAttributes><scg:Dictionary x:TypeArguments="x:String, x:String"><x:String x:Key="StageAttribute">StageFooter-3A59FA7C-C493-47A1-8F8B-1F481143EB08<\/x:String><\/scg:Dictionary><\/mwaw:SPDesignerXamlWriter.CustomAttributes><\/Sequence><\/Sequence><\/FlowStep><\/Flowchart><\/Activity>'), e.saveDefinition(o), i.load(o, "Id"), i.executeQueryAsyncPromise().fail(u).done(function() {
                                s.resolve()
                            }));
                            s.promise().done(function() {
                                b(o, t)
                            })
                        })
                    },
                    b = function(t, r) {
                        var p, s, h, w;
                        if (c = t.get_id(), p = y(f), p) {
                            u("Failed to replace one or more tokens in the workflow. See installer logs for details");
                            return
                        }
                        if (CIB.utilities.deserialiseSharePointObject(JSON.stringify(f), t), f.properties) {
                            s = f.properties;
                            for (h in s) t.setProperty(h, s[h])
                        }
                        w = function() {
                            var t = new jQuery.Deferred,
                                r;
                            return v.length > 0 ? (r = 0, v.forEach(function(f) {
                                var c = y(f),
                                    s, h, a;
                                if (c) return u("Failed to replace one or more tokens in a workflow form"), t.reject(), !1;
                                n.message("Uploading workflow file " + f.url);
                                var l = ($.getServerRealtiveHostWebUrl() + "/" + f.url).replace("//", "/"),
                                    e = l.split("/"),
                                    o = new SP.FileCreationInformation;
                                for (o.set_content(new SP.Base64EncodedByteArray), s = 0; s < f.contents.length; s++) o.get_content().append(f.contents.charCodeAt(s));
                                o.set_overwrite(!0);
                                o.set_url(e[e.length - 1]);
                                e.splice(e.length - 1, 1);
                                h = i.get_web().getFolderByServerRelativeUrl(e.join("/")).get_files();
                                a = h.add(o);
                                i.executeQueryAsyncPromise().fail(u).done(function() {
                                    ++r == v.length && t.resolve()
                                })
                            })) : t.resolve(), t.promise()
                        };
                        t.set_draftVersion("");
                        e.saveDefinition(t);
                        i.load(t, "Id");
                        $.when(w(), i.executeQueryAsyncPromise()).fail(u).done(function() {
                            var s, h, v;
                            e.publishDefinition(t.get_id());
                            s = o.enumerateSubscriptionsByDefinition(c);
                            i.load(s);
                            h = [];
                            for (v in r) h.push(r[v].toString().toLowerCase());
                            i.executeQueryAsyncPromise().fail(u).done(function() {
                                for (var e = {}, v = !1, p = s.getEnumerator(), r, w; p.moveNext();)
                                    if (r = p.get_current(), t.get_restrictToType() != "List" || (w = r.get_eventSourceId().toString(), !(h.indexOf(w.toLowerCase()) < 0))) {
                                        if (e[r.get_name()]) {
                                            u("The workflow definition " + f.displayName + " has more than one associaiton named " + r.get_name());
                                            v = !0;
                                            break
                                        }
                                        e[r.get_name()] = r
                                    }
                                if (!v) {
                                    var c = !1,
                                        k = 0,
                                        b = new jQuery.Deferred;
                                    $.each(a, function(n, t) {
                                        if (c |= y(t), c) return u("Failed to replace one or more tokens in the association " + t.name + ". See installer logs for details"), !1;
                                        var e = i.get_web().get_lists().getById(t.eventSourceId),
                                            f = e.get_fields(),
                                            r = t.statusFieldName;
                                        i.load(f, "Include(InternalName)");
                                        i.executeQueryAsyncPromise().fail(u).done(function() {
                                            for (var n = !0, t = f.getEnumerator(), i; t.moveNext();)
                                                if (i = t.get_current(), i.get_internalName() === r) {
                                                    n = !1;
                                                    break
                                                }
                                            if (n) {
                                                var e = "<Field Type='URL' DisplayName='" + r + "' Name='" + r + "' />",
                                                    u = f.addFieldAsXml(e, !0, SP.AddFieldOptions.addToNoContentType),
                                                    o = unescape(r.replace(/_x/g, "%u").replace(/_/g, ""));
                                                u.set_title(o);
                                                u.update()
                                            }++k === a.length && b.resolve()
                                        })
                                    });
                                    b.promise().done(function() {
                                        ($.each(a, function(t, r) {
                                            var s, h, c;
                                            if (n.message("Creating workflow association " + r.name), s = e[r.name], s || (s = new SP.WorkflowServices.WorkflowSubscription.newObject(i)), CIB.utilities.deserialiseSharePointObject(JSON.stringify(r), s), r.properties) {
                                                h = r.properties;
                                                for (c in h) s.setProperty(c, h[c])
                                            }
                                            if (f.restrictToType == "List") s.setProperty("StatusColumnCreated", "1"), o.publishSubscriptionForList(s, r.eventSourceId);
                                            else if (f.restrictToType == "Site") o.publishSubscription(s);
                                            else return u("Cannot create association as the restrictToType " + f.get_restrictToType() + " was not recognised"), !1
                                        }), c) || i.executeQueryAsyncPromise().fail(u).done(function() {
                                            n.message("Workflow definition " + f.displayName + " created", "success");
                                            l.resolve()
                                        })
                                    })
                                }
                            })
                        })
                    },
                    y = function(t) {
                        var i = !1,
                            f = new RegExp("{\\$([^:]*):([^}]*)}", "gi"),
                            u = function(t, e) {
                                for (var o in t) typeof t[o] == "string" ? t[o] && (t[o] = t[o].replace(f, function(t, u, f) {
                                    if (u === "List") {
                                        if (r[f]) return r[f]
                                    } else if (u === "Web") {
                                        if (f === "$") return s.get_id().toString();
                                        if (f === "%") return s.get_url().toString();
                                        if (f === "^") return s.get_serverRelativeUrl()
                                    } else if (u === "Site") {
                                        if (f === "$") return p.get_id().toString()
                                    } else if (u === "Definition") {
                                        if (f === "$") return c.toString();
                                        if (f === "&") return c.toString().replace(/\-/gi, "")
                                    }
                                    i = !0;
                                    var e = "Failed to replace token in workflow, a " + u + " cannot be found with the name: " + f;
                                    return n.message(e, "error"), t
                                })) : e < 3 && u(t[o], e + 1)
                            };
                        return u(t, 0), i
                    };
                return l.promise()
            },
            ensureAccordianGroup: function(r, u, f) {
                var e = new jQuery.Deferred,
                    f = CIB.utilities.ensureArray(f);
                if (!r || !u || !f || f.length == 0) throw new Error("List title, group title and fields must be set");
                n.message("Creating accordion group '" + u + "' on list '" + r + "'");
                var o = function(t) {
                        n.message(t, "error");
                        e.reject(t)
                    },
                    a = i.get_web(),
                    c = a.get_lists().getByTitle(r),
                    l = c.get_rootFolder(),
                    s = l.get_properties(),
                    h = {};
                return $.each(f, function(n, i) {
                    h[i] = c.get_fields().getByInternalNameOrTitle(i);
                    t.load(h[i], "InternalName", "Title", "Required")
                }), t.load(s), t.executeQueryAsyncPromise().fail(o).done(function() {
                    var p = s.get_fieldValues().CIBListFormAccordionSetting,
                        nt = new DOMParser,
                        i, v, w, y, g, c;
                    if (p || (p = '<?xml version="1.0" encoding="utf-16"?>                            <AccordionSettings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">                                <Groups><\/Groups>                            <\/AccordionSettings>'), i = nt.parseFromString(p, "text/xml"), $.each(i.getElementsByTagName("Group"), function(n, t) {
                            var i = t.getElementsByTagName("Name")[0];
                            if (i.textContent == u) return v = t, !1
                        }), !v) {
                        var b = i.getElementsByTagName("Groups")[0],
                            a = i.createElement("Group"),
                            tt = i.createElement("Fields"),
                            k = i.createElement("Name"),
                            d = i.createElement("Order");
                        k.appendChild(i.createTextNode(u));
                        d.appendChild(i.createTextNode(b.childElementCount + 1));
                        a.appendChild(tt);
                        a.appendChild(k);
                        a.appendChild(d);
                        b.appendChild(a);
                        v = a
                    }
                    w = 0;
                    y = !1;
                    $.each(f, function(t, r) {
                        var e = h[r],
                            s = e.get_internalName(),
                            c = e.get_title(),
                            g = e.get_required().toString().toLowerCase(),
                            p = v.getElementsByTagName("Fields")[0],
                            l = !1,
                            a;
                        if ($.each(p.getElementsByTagName("Field"), function(t, i) {
                                if (i.getElementsByTagName("InteralName")[0].textContent == s) return l = !0, w++, n.message(c + " is already present in accordion group " + u, "info"), !1
                            }), a = !1, l || $.each(i.getElementsByTagName("Field"), function(n, t) {
                                if (t.getElementsByTagName("InteralName")[0].textContent == s) return a = !0, !1
                            }), a) return o(c + " is already present in a different accordion group"), y = !1, !1;
                        if (!l) {
                            var f = i.createElement("Field"),
                                b = i.createElement("DisplayName"),
                                k = i.createElement("InteralName"),
                                d = i.createElement("Required");
                            b.appendChild(i.createTextNode(c));
                            k.appendChild(i.createTextNode(s));
                            d.appendChild(i.createTextNode(g));
                            f.appendChild(b);
                            f.appendChild(k);
                            f.appendChild(d);
                            p.appendChild(f);
                            y = !0
                        }
                    });
                    y ? (g = new XMLSerializer, c = g.serializeToString(i), c = c.replace("<AccordionSettings>", '<AccordionSettings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'), c = '<?xml version="1.0" encoding="utf-16"?>' + c, s.set_item("CIBListFormAccordionSetting", c), l.update(), t.executeQueryAsyncPromise().fail(o).done(function() {
                        n.message("Accordion group '" + u + "' added to list '" + r + "'", "success");
                        e.resolve()
                    })) : w == f.length && e.resolve()
                }), e.promise()
            },
            displaySettings: {
                always: "always",
                never: "never",
                whereInGroup: "whereInGroup",
                whereNotInGroup: "whereNotInGroup"
            },
            displaySettingForm: {
                displayForm: "Display",
                editForm: "Edit",
                newForm: "New"
            },
            displayMode: {
                write: "writable",
                read: "read-only"
            },
            updateDisplaySettings: function(r, u) {
                var f = CIB.installer,
                    e = new jQuery.Deferred,
                    u = CIB.utilities.ensureArray(u);
                if (!r || !u || u.length == 0) throw new Error("List title, and display settings must be set");
                $.each(u, function(n, t) {
                    if (!t.field || !t.form || !t.display) throw new Error("Field, Form and Display properties must be set on the field");
                    if (t.display == f.displaySettings.whereInGroup || t.display == f.displaySettings.whereNotInGroup) {
                        if (!t.group) throw new Error("Group must be set on the field for where display settings");
                    } else if (Array.isArray(t.display) && ($.each(t.display, function(n, t) {
                            if (!t.condition || !t.groupName || !t.mode) throw new Error("Condition,Group name and Mode must be set on field for where display settings");
                        }), !t.logic)) throw new Error("logic condition must be set on multiple group where field");
                });
                n.message("Updating display settings on list '" + r + "'");
                var h = function(t) {
                        n.message(t, "error");
                        e.reject(t)
                    },
                    o = i.get_web(),
                    s = o.get_allProperties(),
                    c = o.get_lists().getByTitle(r);
                return t.load(c, "Id"), t.load(s), t.executeQueryAsyncPromise().fail(h).done(function() {
                    var y = c.get_id().toString(),
                        a = ("DisplaySetting" + y).toLowerCase(),
                        i = s.get_fieldValues()[a],
                        l, v;
                    for (i = i ? i.split("#") : [], i = i.filter(function(n) {
                            return Boolean(n)
                        }), i = i.map(function(n) {
                            return n.split("|")
                        }), $.each(u, function(n, t) {
                            for (var e, r, h, s, o = -1, u = 0; u < i.length; u++)
                                if (i[u][0] === t.field && i[u][1] === t.form) {
                                    o = u;
                                    break
                                }
                            if (e = t.display, Array.isArray(e) ? (r = "where", $.each(e, function(n, t) {
                                    r += ";[Me];";
                                    r += t.condition === f.displaySettings.whereNotInGroup ? "IsNotInGroup;" : "IsInGroup;";
                                    r += t.groupName + ";" + t.mode;
                                    n !== e.length - 1 && (r += ";$where")
                                }), t.logic && (r += ";~" + t.logic)) : ((e === f.displaySettings.whereInGroup || e === f.displaySettings.whereNotInGroup) && (e = "where"), r = e + ";[Me];", r += t.display === f.displaySettings.whereNotInGroup ? "IsNotInGroup;" : "IsInGroup;", t.group ? (r += t.group + ";", r += ";writable;~AND") : r += "Approvers;writable;~AND"), o >= 0) i[o][0] = t.field, i[o][1] = t.form, i[o][2] = r;
                            else {
                                for (i.push([t.field, t.form, r]), h = {}, u = 0; u < i.length; u++) i[u][0] === t.field && (h[i[u][1]] = !0);
                                for (s in f.displaySettingForm) s = f.displaySettingForm[s], h[s] || i.push([t.field, s, f.displaySettings.always + ";[Me];IsInGroup;Approvers;writable;~AND"])
                            }
                        }), l = 0; l < i.length; l++) i[l] = i[l].join("|");
                    v = i.join("#") + "#";
                    s.set_item(a, v);
                    o.update();
                    t.executeQueryAsyncPromise().fail(h).done(function() {
                        n.message("Display settings updated on list '" + r + "'", "success");
                        e.resolve()
                    })
                }), e.promise()
            },
            updateWebPartProperties: function(r) {
                var r = CIB.utilities.ensureArray(r),
                    u = new jQuery.Deferred,
                    f = function(t) {
                        n.message(t, "error");
                        u.reject(t)
                    },
                    o = i.get_web(),
                    e = 0;
                return $.each(r, function() {
                    var i = this;
                    if (!i.title || !i.file || !i.properties) throw new Error("web part must have title, file and properties attributes set");
                    n.message("Updating web part '" + i.title + "'.");
                    var h = o.getFileByServerRelativeUrl(i.file),
                        c = h.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared),
                        s = c.get_webParts();
                    t.load(s, "Include(WebPart.Title, WebPart.Properties)");
                    t.executeQueryAsyncPromise().fail(f).done(function() {
                        for (var l = !1, a = s.getEnumerator(), o, h, c, v; a.moveNext();)
                            if (o = a.get_current(), h = o.get_webPart(), h.get_title() === i.title || i.title === "*") {
                                l = !0;
                                for (c in i.properties) v = i.properties[c], h.get_properties().set_item(c, v);
                                o.saveWebPartChanges()
                            }
                        l ? t.executeQueryAsyncPromise().fail(f).done(function() {
                            n.message("Updated properties for web part '" + i.title + "'.", "success");
                            ++e == r.length && u.resolve()
                        }) : (n.message("No web part with a title '" + i.title + "' was found in file: " + i.file, "info"), ++e == r.length && u.resolve())
                    })
                }), u.promise()
            },
            getContentTypeIdByName: function(r) {
                var u, f;
                if (!r) throw new Error("Content Type Name cannot be null");
                u = new jQuery.Deferred;
                f = function(t) {
                    n.message(t, "error");
                    u.reject(t)
                };
                n.message("Fetching content type id for name: " + r);
                var e = "",
                    s = i.get_web(),
                    o = s.get_availableContentTypes();
                return t.load(o, "Include(Id, Name)"), t.executeQueryAsyncPromise().fail(f).done(function() {
                    for (var t = o.getEnumerator(), n; t.moveNext();)
                        if (n = t.get_current(), n.get_name() === r) {
                            e = n.get_id();
                            break
                        }
                    u.resolve(e)
                }), u.promise()
            }
        }
    }(),
    function() {
        SP.RequestExecutor && (SP.RequestExecutorInternalSharedUtility.BinaryDecode = function(n) {
            var i = "",
                r, t;
            if (n)
                for (r = new Uint8Array(n), t = 0; t < n.byteLength; t++) i = i + String.fromCharCode(r[t]);
            return i
        }, SP.RequestExecutorUtility.IsDefined = function(n) {
            return n === null || typeof n == "undefined" || !n.length
        }, SP.RequestExecutor.ParseHeaders = function(n) {
            var i, t, r, u, f;
            if (SP.RequestExecutorUtility.IsDefined(n)) return null;
            var e = {},
                s = new RegExp("\r?\n"),
                o = n.split(s);
            for (i = 0; i < o.length; i++) t = o[i], SP.RequestExecutorUtility.IsDefined(t) || (r = t.indexOf(":"), r > 0 && (u = t.substr(0, r), f = t.substr(r + 1), u = SP.RequestExecutorNative.trim(u), f = SP.RequestExecutorNative.trim(f), e[u.toUpperCase()] = f));
            return e
        }, SP.RequestExecutor.internalProcessXMLHttpRequestOnreadystatechange = function(n, t, i) {
            var r, u, f;
            n.readyState === 4 && (i && window.clearTimeout(i), n.onreadystatechange = SP.RequestExecutorNative.emptyCallback, r = new SP.ResponseInfo, r.state = t.state, r.responseAvailable = !0, r.body = t.binaryStringResponseBody ? SP.RequestExecutorInternalSharedUtility.BinaryDecode(n.response) : n.responseText, r.statusCode = n.status, r.statusText = n.statusText, r.contentType = n.getResponseHeader("content-type"), r.allResponseHeaders = n.getAllResponseHeaders(), r.headers = SP.RequestExecutor.ParseHeaders(r.allResponseHeaders), n.status >= 200 && n.status < 300 || n.status === 1223 ? t.success && t.success(r) : (u = SP.RequestExecutorErrors.httpError, f = n.statusText, t.error && t.error(r, u, f)))
        })
    }()