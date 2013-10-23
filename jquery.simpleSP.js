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

$.ssp.List.prototype.addColumn = function(colDesc) {
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

	if (desc.type && desc.type !== "GET") {
		opts.headers['X-RequestDigest'] = $("#__REQUESTDIGEST").val();
	}

	
 
	$.ajax({
		type: "GET",
		url: site + "/_api/web" + opts.path,
		headers: opts.headers, 
		async: false,
		data: opts.data,
		success: function(res) {
			ret = res.d;
		}
	});
	
	return ret;			

}

})(jQuery);
