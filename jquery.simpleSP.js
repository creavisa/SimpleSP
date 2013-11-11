/* 
 * Simple Share Point Library
 * (c) 2013 Creavisa
 */

// Closure to avoid unnecesary globals
(function ($) {

var context = {},
    site = {
	  baseUrl: ""
    };

$.ssp = function(options) {
	site = new $.ssp.Site({
		baseUrl: options.url || 
		     window._spPageContextInfo.siteAbsoluteUrl ||
		     "",
		load: options.load || false
	});

	return site;
}

$.ssp.Site = function(opts, create) {
	var info = {},
	    desc = {},
	    // English
	    lang = 1033;
	
	if (opts.load) {
		info = request({site: opts.baseUrl});
		$.extend(this, info);
	} else if (create) {
		info = request({
			site: opts.baseUrl,
			path: opts.path || "/"
		});
	}
	
	if (create && info.error) {
		desc = {
		parameters: {
			__metadata:  {
				type: "SP.WebInfoCreationInformation" 
			}, 
			Url: opts.Url || opts.baseUrl, 
			Title: opts.Title, 
			Description: opts.Description,
			Language: lang,
			WebTemplate: opts.Template,
			UseUniquePermissions: true	
		}};
		
		info = request({
			path: "/Webs/Add",
			type: "POST",
			data: desc
		});
		
	}

	this.baseUrl = opts.baseUrl || opts.url;

	return this;
}

$.ssp.Site.prototype.setGlobalNavigation = function(style, terms) {
	var ctx = new SP.ClientContext(site.baseUrl),
	    web = ctx.get_web(),
		navSettings = new SP.Publishing.Navigation.WebNavigationSettings(ctx, web),
		nav = navSettings.get_globalNavigation();
	
	/*
	 * unknown: 0, 
	 * portalProvider: 1 (Structural Navigation), 
	 * taxonomyProvider: 2 (You will need the Term Store Id and Term Set Id too),
	 * inheritFromParentWeb: 3 
	 */
	if (style === "taxonomy" && terms) {
		nav.set_termStoreId(terms.storeId);
		nav.set_termSetId(terms.setId)
		nav.set_source(2);
	} else if (style === "provider") {
		nav.set_source(1);
	} else {
		//Inherit
		nav.set_source(3);
	}
	
	navSettings.update();
	nav.update();
	
	ctx.executeQueryAsync();
}

$.ssp.Site.prototype.setQuickNavigation = function(style, terms) {
	var ctx = new SP.ClientContext(site.baseUrl),
	    web = ctx.get_web(),
		navSettings = new SP.Publishing.Navigation.WebNavigationSettings(ctx, web),
		nav = navSettings.get_currentNavigation();
		
	if (style === "taxonomy" && terms) {
		nav.set_termStoreId(terms.storeId);
		nav.set_termSetId(terms.setId);
		nav.set_source(2);
	} else if (style === "provider") {
		nav.set_source(1);
	} else {
		//Inherit
		nav.set_source(3);
	}
	
	navSettings.update();
	nav.update();
	
	ctx.executeQueryAsync();
}

$.ssp.getLists = function() {
	var res;
	res = request({
		path: "/Lists"
		});
	
	if (res.error) {
		return [];
	}
	
	return res.results;
}

// List object definition
$.ssp.List = function (opts, create) {
	var list,
		c,
	    desc, //List description
	    req = "/Lists";
		
	this.site = opts.site || site.baseUrl;
		
	if (typeof opts === "string") {
		opts = {
			Title: opts
		};
	}

	if (opts.Id) {
		req += "('" + opts.Id + "')";
	} else if (opts.Title) {
		req += "/GetByTitle('" + opts.Title +"')";
	}

	list = request({
		path: req,
		site: this.site
	});
	
	if (list.error && create && opts.Title) {
		desc = {
			__metadata: {
				type: "SP.List"
			},
			Title: opts.Title,
			Description: opts.Description || "",
			AllowContentTypes: true,
			ContentTypesEnabled: true,
			BaseTemplate: opts.Template || 100 
			};
			
		list = request({
			type: "POST",
			path: "/Lists",
			data: desc,
			site: this.site
			});
			
		if (!list.error && opts.cols) {
			for (c in opts.cols) {
			if (opts.cols.hasOwnProperty(c)) {
				this.addColumn(opts.cols[c]);
			}
			}
		}
	} 
	
	// If it doesn't have a title of the error
	// persists:
	if (list.error) {
		// Returns a List object with an error
		$.extend(this, list);
		return this;
	}
	
	$.extend(this, list);
	this.updateItems();
	
	return this;
}

$.ssp.List.prototype.updateItems = function(async) {
	var tmp,
	    i = 0,
	    list = this;
		
	//Get list items
	tmp = request({
		path: "/Lists('" + list.Id +"')/Items",
		async: async,
		site: this.site
		});
	
	if (tmp && tmp.results) {
		list.Items = tmp.results;
	} else {
		list.Items = [];
	}
	
	context.list = {Id: list.Id };
	for (i in list.Items) {
	if (list.Items.hasOwnProperty(i)){
		list.Items[i].site = list.site;
		list.Items[i] = new $.ssp.List.Item(list.Items[i]);
	}
	}

	$.extend(this, list);

	return this;
}

$.ssp.List.prototype.addColumn = function(col) {
	var list = this,
	    colDesc,
	    type,
	    status;

	if (!list.Id) {
		return;
	}

	if (!col.type) {
		type = "Text";
	}
		
	// SharePoint expects the type value as a number.
	if (col.type === "Integer") {
		type = 1;
	} else if (col.type === "Text") {
		type = 2;
	} else if (col.type === "Boolean") {
		type = 8;
	} else if (col.type === "Guid") {
		type = 14;
	} else {
		//Invalid
		type = 0;
	}
		
	colDesc = {
		__metadata: {
			type: "SP.Field"
		},
		Title: col.Title,
		FieldTypeKind: type,
		// Use 'false' as default to avoid 
		// sending null or undef
		Required: col.Required || false,
		EnforceUniqueValues: col.Unique || col.EnforceUniqueValues || false
	};
		
	status = request({
		type: "POST",
		data: colDesc,
		path: "/Lists('" + list.Id + "')/Fields",
		site: this.site
		});	
}

$.ssp.List.prototype.add = function(item) {
	var list = this,
	    desc = {},
	    status;

	if (!list.Id || !item.Title) {
		return;
	}

	desc = {
		"__metadata": {
			type: list.ListItemEntityTypeFullName
		}
	};

	$.extend(desc, item);

	// Find a way to list the extra columns per item.type
	// and fill them, if the item has them declared
	status = request({
		type: "POST",
		path:  "/Lists('"+ list.Id +"')/Items",
		data: desc,
		site: this.site
	});

	this.updateItems();
	return status;
}

$.ssp.List.prototype.rm = function(item) {
	var status,
	    list = this;

	if (!item.Id || !list.Id) {
		return;
	}

	status = request({
		type: "DELETE",
		path: "/List('"+ list.Id +"')/Items('" + item.Id + "')",
		site: this.site
	});

	this.updateItems();
	return status;
}

$.ssp.List.prototype.grant = function(entity, role) {
	var ctx = new SP.ClientContext(this.site),
	    web = ctx.get_web(),
	    target = web.get_lists().getById(this.Id);
			
	// The 'false' means that it shouldn't copy it's role
	// Assignment from it's parent.
	target.breakRoleInheritance(false);

	var roleDef, grp, bindings, assignments;
	
	roleDef = ctx.get_site().get_rootWeb().get_roleDefinitions().getById(role.Id);
	
	bindings = SP.RoleDefinitionBindingCollection.newObject(ctx);
	assignments = target.get_roleAssignments();
	bindings.add(roleDef);
	
	if (entity.EntityType == "User") {
		grp = web.get_siteUsers().getByEmail(entity.Description);
	} else {
		grp = web.get_siteGroups().getById(entity.Id);
	}
	
	assignments.add(grp, bindings);
		
	ctx.executeQueryAsync(
		function() {console.log("Role assigments set")},
		function() {console.error("failed setting list assignments")}
	);
}

$.ssp.List.prototype.removeList = function() {
	var status,
	    list = this;

	if (!list.Id) {
		return;
	}

	status = request({
		type: "DELETE",
		path: "/Lists('"+ list.Id +"')",
		site: this.site
	});

	return status;
}

$.ssp.List.prototype.getBy = function (field, value) {
	var path,
	    res;
		
	path = "/List('"+ this.Id +"')/Items";
	path += "?$filter="+ field +" eq '"+ value +"'";
	
	res = request({
		path: path,
		site: this.site
	});
	
	if (res.error) {
		return res;
	}else if (res.results) {
		return res.results;
	} else {
		return;
	}
}

$.ssp.List.Item = function(desc) {
	var res;
	this.site = desc.site || site.baseUrl;
	
	res = request({
		path: "/Lists('"+ context.list.Id +"')/Items('"+ desc.Id +"')",
		site: this.site
	});
	
	if (!res.error) {
		$.extend(this, res);
	}

	return this;
}

$.ssp.List.Item.prototype.update = function(changes) {
	var res, info = {
		__metadata: this.__metadata
	};

	$.extend(changes, info);
	
	res = $.ssp.req({
		type: "MERGE",
		etag: this.__metadata.etag,
		uri: this.__metadata.uri,
		data: changes
	});

	return res;
}

$.ssp.Group = function(opts, create) {
	var group,
	    res,
	    path;
	
	this.site = opts.site || site.baseUrl;
	
	if (typeof opts === "string") {
		opts = {Title: opts};
	}
	
	if (opts.Title) {
		path = "/SiteGroups/GetByName('" + opts.Title +"')";
	} else if (opts.Id) {
		path = "/SiteGroups/GetById('" + opts.Id +"')";
	} else {
		return this;
	}
	
	res = request({
		site: this.site,
		path: path
	});
	
	if (create && res.error && opts.Title) {
		path = "/SiteGroups"
		group = {
			__metadata: {
				type: "SP.Group"
			},
			Title: opts.Title,
			Description: opts.Description || ""
		};
		
		res = request({
			type: "POST",
			path: path,
			site: this.site,
			data: group
		});
	}	

	$.extend(this, res);
	
	return this;	
}

$.ssp.Group.prototype.add = function(user) {
	var desc,
	    res = {};
		
	desc ={
		__metadata: {
			type: "SP.User"
		},
		Email: user.Description,
		Title: user.DisplayText,
		LoginName: user.Key	
	};

	res = request({
		type: "POST",
		path: "/SiteGroups("+ this.Id +")/Users",
		site: this.site
	});
	
	return res;
}

$.ssp.Group.prototype.rm = function(user) {
	return request({
		type: "DELETE",
		path: "/SiteGroups("+this.Id+")/Users/GetById("+ user.Id +")",
		site: this.site
		});
}

$.ssp.Role = function(opts, create) {
	var res,
	    path,
	    role = {}
	
	this.site = opts.site || site.baseUrl;
		
	if (typeof opts === "string") {
		opts = {Name: opts};
	}
	
	if (opts.Name) {
		path = "/RoleDefinitions/GetByName('" + opts.Name +"')";
	} else if (opts.uuid) {
		path = "/RoleDefinitions/GetById('" + opts.uuid +"')";
	} else {
		return this;
	}
	
	res = request({
		site: this.site,
		path: path
	});
	
	if (create && res.error && opts.Name) {
		role = {
			__metadata: {
				type: "SP.RoleDefinition"
			},
			Name: opts.Name,
			Description: opts.Description || "",
			// Use BasePermissions the SP namespace
			BasePermissions: opts.BasePermissions || new SP.BasePermissions(), 
			Order: 1
		};
		
		res = request({
			type: "POST",
			site: this.site,
			path: "/RoleDefinitions",
			data: role
		});
	}
		
	$.extend(this, res);
		
	return this;
}

$.ssp.getTemplates = function() {
	var templates,
	    lang = "1033";
	
	templates = request({
		path: "/GetAvailableWebTemplates("+ lang +")"
	});

	if (templates.error) {
		return [];
	}
	
	return templates.results;
}

$.ssp.getSubSites = function() {
	var sites;

	sites = request({
		path: "/Webs"
	});
	
	if (sites.error) {
		return [];
	}

	return sites.results;
}

// Render and initialize the client-side People Picker.
$.ssp.peoplePicker = function(ElementId, users) {
	// Create a schema to store picker properties, and set the properties.
	var schema = {
		PrincipalAccountType: "User",
		SearchPrincipalSource: 15,
		ResolvePrincipalSource: 15,
		AllowMultipleValues: true,
		MaximumEntitySuggestions: 50,
		width: "200px"
	};

	if (!users) {
		users = [];
	}

	// Render and initialize the picker. 
	// Pass the ID of the DOM element that contains the picker, an array of initial
	// PickerEntity objects to set the picker value, and a schema that defines
	// picker properties.
	window.SPClientPeoplePicker_InitStandaloneControlWrapper(ElementId, users, schema);
}

// Automate somethings about the request and make it 
// synchronous to avoid race conditions when creating sites
// or lists.
function request(desc) {
	var ret,
	    opts = {
		path: desc.path || "",
		site: desc.site || site.baseUrl,
		type: desc.type || "GET",
		data: desc.data || "",
		async: desc.async || false,
		headers: {
			"Accept": "application/json;odata=verbose",
			"Content-Type": "application/json;odata=verbose"
		}
			
	};
	
	opts.path = "/_api/Web" + opts.path;

	if (opts.type !== "GET") {
		opts.headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
		// This is forcing the write, must fix
		// Maybe have the eTag with the rest of the data
		opts.headers["If-Match"] = opts.etag || "*";
	}

	if (opts.data && opts.type !== "GET") {
		opts.data = JSON.stringify(opts.data);
	} 
	
	if (opts.type === "MERGE") {
		opts.headers["X-HTTP-Method"] = "MERGE";
		opts.type = "POST";
	}
	
	if (desc.uri) {
		opts.site = desc.uri;
		opts.path = "";
	}
 
	$.ajax({
		type: opts.type,
		url: opts.site + opts.path,
		headers: opts.headers, 
		async: opts.async,
		data: opts.data, 
		success: function(res) {
			ret = res.d;
		},
		error: function(req, status, error) {
			ret = {
				status: status,
				error: error
			};
		}
	});
	
	return ret;			
}

$.ssp.req = request;

})(jQuery);
