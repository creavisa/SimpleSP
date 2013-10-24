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

$.ssp.Site = function(opts) {
	var info;
	if (this.load) {
		info = request({site: opts.url});
		$.extend(this, info);
	} 

	this.baseURL = opts.url;

	return this;
}

$.ssp.getLists = function() {
	return request({path: '/Lists'});
}

// List object definition
$.ssp.List = function (title, uuid) {
	var obj,
	    req = '/Lists';

	if (uuid) {
		req += "('" + uuid + "')";
	} else if (title) {
		req += "/Lists/GetByTitle('" + title +"')";
	}

	obj = request({path: req});
	//Get list items
	var tmp = request({path: req + "/items"});
	if (tmp && tmp.results) {
		obj.items = tmp.results;
	} else {
		obj.itmes = [];
	}

	$.extend(this, obj);

	return this;
}

$.ssp.List.prototype.setTitle = function(title) {
}

$.ssp.List.prototype.addColumn = function(col) {
	var list = this,
	    colDesc,
	    status;

	if (!list.uuid) {
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
		Title: name,
		FieldTypeKind: type,
		// Use 'false' as default to avoid 
		// sending null or undef
		Required: required || false,
		EnforceUniqueValues: unique || false
	};
		
	status = request({
		type: "POST",
	    	path: "/List('" + list.uuid "')/Fields",
	    	data: colDesc
		});	
}

$.ssp.List.prototype.add = function(item) {
	var list = this,
	    desc = {},
	    status;

	if (!list.uuid || !item.title) {
		return;
	}

	if (!item.type) {
		// get default type from list
		// item.type = list.defaultType;
	}

	desc = {
		__metadata: {
			type: item.type
		},
		Title: item.title
	};

	// Find a way to list the extra columns per item.type
	// and fill them, if the item has them declared
	status = request({
		type: "POST",
		path:  "/Lists('"+ list.uuid +"')/Items"
	});

	return status;
}

$.ssp.List.prototype.del = function(item) {
	var status,
	    list = this;

	if (!item.uuid || !list.uuid) {
		return;
	}

	status = request({
		type: "DELETE",
		path: "/List('"+ list.uuid +"')/Items('" + item.uuid + "')"
	});

	return status;
}

$.ssp.List.prototype.removeList = function() {
	var status,
	    list = this;

	if (!list.uuid) {
		return;
	}

	status = request({
		type: "DELETE",
		path: "/List('"+ list.uuid +"')"
	});

	return status;
}

$.ssp.getTemplates = function() {
	var templates,
	    lang = "1033";
	
	templates = request({
		path: "/GetAvailableWebTemplates("+ lang +")"
	});

	return templates;
}

$.ssp.getSubSites = function() {
	var sites;

	sites = request({
		path: "/Webs"
	});

	return sites;
}

$.ssp.createSite = function(desc) {
	var status, 
	    lang = 1033,
	    payload = {
	parameters: {
		'__metadata':  {
			'type': 'SP.WebInfoCreationInformation' 
		}, 
		'Url': desc.path, 
		'Title': desc.name, 
		'Description': desc.desc,
		'Language': lang,
		'WebTemplate':desc.tpl,
		'UseUniquePermissions': true	
	}};
	
	status = request({
		type: "POST",
		path: "/Webinfos/Add",
		data: payload
		});

	return status;
}

// Automate somethings about the request and make it 
// sincronous, to avoid race conditions when creating sites
// or lists.
function request(desc) {
	var ret,
	    opts = {
		path: desc.path || "",
		site: desc.site || site.baseURL,
		type: desc.type || "GET",
		data: dest.data || "",
		headers: {
			"Accept": "application/json;odata=verbose",
			"Content-Type": "application/json;odata=verbose"
		}
			
	};

	if (desc.type !== "GET") {
		opts.headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
		// This is forcing the write, must fix
		// Maybe have the eTag with the rest of the data
		opts.headers["If-Match"] = "*";
	}

	if (dest.data && desc.type !== "GET") {
		dest.data = JSON.stringify(dest.data);
	}
 
	$.ajax({
		type: "GET",
		url: opts.site + "/_api/Web" + opts.path,
		headers: opts.headers, 
		async: false,
		data: dest.data, 
		success: function(res) {
			ret = res.d;
		}
	});
	
	return ret;			
}

})(jQuery);
