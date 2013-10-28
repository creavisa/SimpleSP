/* 
 * Simple Share Point Library
 * (c) 2013 Creavisa
 */

// Closure to avoid unnecesary globals
(function ($) {

var  site = {
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
			path: opts.path || '/'
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

$.ssp.getLists = function() {
	var res;
	res = request({path: '/Lists'});
	
	if (res.error) {
		return [];
	}
	
	return res.results;
}

// List object definition
$.ssp.List = function (opts, create) {
	var list,
	    desc, //List description
	    req = '/Lists';
		
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

	list = request({path: req});
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
			data: desc
			});
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
		async: async
		});
	
	if (tmp && tmp.results) {
		list.Items = tmp.results;
	} else {
		list.Items = [];
	}

	for (i in list.Items) {
	if (list.Items.hasOwnProperty(i)){
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
		type = 'Text';
	}
		
	// SharePoint expects the type value as a number.
	if (type === "Integer") {
		type = 1;
	} else if (type === "Text") {
		type = 2;
	} else if (type === "Boolean") {
		type = 8;
	} else if (type === "Guid") {
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
		path: "/List('" + list.Id + "')/Fields"
		});	
}

$.ssp.List.prototype.add = function(item) {
	var list = this,
	    desc = {},
	    status;

	if (!list.Id || !item.Title) {
		return;
	}

	if (!item.type) {
		// get default type from list
		item.type = list.ListItemEntityTypeFullName;
	}

	desc = {
		"__metadata": {
			type: item.type
		},
		Title: item.Title
	};

	// Find a way to list the extra columns per item.type
	// and fill them, if the item has them declared
	status = request({
		type: "POST",
		path:  "/Lists('"+ list.Id +"')/Items",
		data: desc
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
		path: "/List('"+ list.Id +"')/Items('" + item.Id + "')"
	});

	this.updateItems();
	return status;
}

$.ssp.List.prototype.grant = function(group, role) {
	var ctx = new SP.ClientContext(site.baseUrl),
	    web = ctx.get_web(),
	    roleDef = ctx.get_site().get_rootWeb().get_roleDefinitions().getById(role.Id),
	    grp = web.get_siteGroups().getById(group.Id),
	    newBindings = SP.RoleDefinitionBindingCollection.newObject(ctx),
	    target = web.get_lists().getById(this.Id),
	    roleAssignments;
			
	// Add the role to the collection.
	newBindings.add(roleDef);
	// Get the list to work with and break permissions 
	// so its permissions can be managed directly.
	target.breakRoleInheritance(true, false);
	// Get the RoleAssignmentCollection for the target list.
	roleAssignments = target.get_roleAssignments();
	// Add the user to the target list and assign the use 
	// to the new RoleDefinitionBindingCollection.
	roleAssignments.add(grp, newBindings);
	ctx.executeQueryAsync(
		function() {console.log("Succeed") },
		function() {console.log("Failed")}
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
		path: "/List('"+ list.Id +"')"
	});

	return status;
}

$.ssp.List.Item = function(desc) {
	$.extend(this, desc);

	return this;
}

$.ssp.List.Item.prototype.update = function(changes) {
	var res;

	$.extend(this, changes);

	res = request({
		type: "POST",
		path: "/Lists("+ this.ListId +")/Item/update("+ this.Id +")",
		data: this
	});

	return res;
}

$.ssp.Group = function(opts, create) {
	var group,
	    res,
	    path;
	
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
	
	res = request({path: path});
	
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
		path: "/SiteGroups("+ this.Id +")/Users"
	});
	
	return res;
}

$.ssp.Group.prototype.rm = function(user) {
	return request({
		type: "DELETE",
		path: "/SiteGroups("+this.Id+")/Users/GetById("+ user.Id +")",
		});
}

$.ssp.Role = function(opts, create) {
	var res,
	    path,
	    role = {}
		
		
	if (typeof opts === "string") {
		opts = {Name: opts};
	}
	
	if (opts.Name) {
		path = "/RoleDefinitions/GetByTitle('" + opts.Name +"')";
	} else if (opts.uuid) {
		path = "/RoleDefinitions/GetById('" + opts.uuid +"')";
	} else {
		return this;
	}
	
	res = request({path: path});
	
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

	if (opts.type !== "GET") {
		opts.headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
		// This is forcing the write, must fix
		// Maybe have the eTag with the rest of the data
		opts.headers["If-Match"] = "*";
	}

	if (opts.data && opts.type !== "GET") {
		desc.data = JSON.stringify(desc.data);
	}
 
	$.ajax({
		type: opts.type,
		url: opts.site + "/_api/Web" + opts.path,
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
