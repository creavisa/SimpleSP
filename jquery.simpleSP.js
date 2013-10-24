/* 
 * Simple Share Point Library
 * (c) 2013 Creavisa
 */

// Closure to avoid unnecesary globals
(function ($) {

var  site = {
	baseURL: ""
};

$.ssp = function(options) {
	site = new $.ssp.Site({
		url: options.url || window.location.host,
		load: options.load || false
	});

	return site;
}

$.ssp.Site = function(opts, create) {
	var info = {},
		desc = {},
		// English
		lang = 1033;
	
	if (opts.load || create) {
		info = request({site: opts.url});
		$.extend(this, info);
	} 
	
	if (create && info.error) {
		desc = {
		parameters: {
			__metadata:  {
				type: "SP.WebInfoCreationInformation" 
			}, 
			Url: opts.Url, 
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

	this.baseURL = opts.url;

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
	    tmp,
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
	
	//Get list items
	tmp = request({path: req + "/Items"});
	if (tmp && tmp.results) {
		list.Items = tmp.results;
	} else {
		list.Items = [];
	}

	$.extend(this, list);

	return this;
}

$.ssp.List.prototype.setTitle = function(title) {
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
		EnforceUniqueValues: col.Unique || Col.EnforceUniqueValues || false
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
		path:  "/Lists('"+ list.Id +"')/Items"
	});

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

	return status;
}

$.ssp.List.prototype.grant = function(group, role) {
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

$.ssp.Group = function(opts, create) {
	var group,
		res,
		path;
	
	if (typeof opts === "string") {
		opts = {Title: opts};
	}
	
	if (opts.Title) {
		path = "/SiteGroups/GetByTitle('" + opts.title +"')";
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
		
		$.extend(this, res);
	}	
	
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
	var res;
	res = request({
		type: "DELETE",
		path: "/SiteGroups("+this.Id+")/Users/GetById("+ user.Id +")",
		success: log
		});

	return res;
}

$.ssp.Role = function(opts, create) {
	
	var res,
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
		site: desc.site || site.baseURL,
		type: desc.type || "GET",
		data: desc.data || "",
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
		async: false,
		data: desc.data, 
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

})(jQuery);
