/* 
 * Simple Share Point Library
 * (c) 2013 Creavisa
 */

// Our Library Object.
// It's exported to the global namespace (usually window).
var SSP = {};

// Closure to avoid unnecesary globals
(function ($) {

// 'site' is only the URL to be used internally
var site,
    siteData = {};

$.ssp = function(options) {
	if (options.site) {
		site = options.site;
	} else {
		site = window.location.host;
	}

	if (options.load) {
		siteData = request();
	}
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
	
}

// Constructor for List item
// Sets default columns and the correct list type
$.ssp.List.prototype.Item = function(contentType) {
}

// Automate somethings about the request and make it 
// sincronous, to avoid race conditions when creating sites
// or lists.
function request(desc) {
	var ret,
	    opts = {
		path: desc.path || "",
		site: desc.site || site,
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
		url: site + "/_api/web" + opts.path,
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
