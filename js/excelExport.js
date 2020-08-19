/*
 * ####################################################################################################
 * https://www.npmjs.com/package/xlsx-style
 * ####################################################################################################
 */
var excelExport = {
    config: {
        fileName: "report",
        extension: ".xlsx",
        sheetName: "Sheet1",
        fileFullName: "report.xlsx",
        header: true,
        createEmptyRow: true,
        maxCellWidth: 20
    },
    worksheetObj: {},
    rowCount: 0,
    wsColswidth: [],
    merges: [],
    worksheet: {},
    range: {},
    init: function (options) {
        this.reset();
        if (options) {
            for (var key in this.config) {
                if (options.hasOwnProperty(key)) {
                    this.config[key] = options[key];
                }
            }
        }
        this.config['fileFullName'] = this.config.fileName + this.config.extension;
    },
    reset: function () {
        this.range = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}};
        this.worksheetObj = {};
        this.rowCount = 0;
        this.wsColswidth = [];
        this.merges = [];
        this.worksheet = {};
    },
    parse2Int0: function (num) {
        num = parseInt(num);
        num = Number.isNaN(num) ? 0 : num;
        return num;
    },
    cellWidth: function (cellText, pos) {
        var max = (cellText && cellText.length * 1.3);
        if (this.wsColswidth[pos]) {
            if (max > this.wsColswidth[pos].wch) {
                this.wsColswidth[pos] = {wch: max};
            }
        } else {
            this.wsColswidth[pos] = {wch: max};
        }
    },
    cellWidthValidate: function () {
        for (var i in this.wsColswidth) {
            if (this.wsColswidth[i].wch > this.config.maxCellWidth) {
                this.wsColswidth[i].wch = this.config.maxCellWidth;
            }
        }
    },
    datenum: function (v, date1904) {
        if (date1904)
            v += 1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    },
    setCellDataType: function (cell) {
        if (typeof cell.v === 'number') {
            cell.t = 'n';
        } else if (typeof cell.v === 'boolean') {
            cell.t = 'b';
        } else if (cell.v instanceof Date) {
            cell.t = 'n';
            cell.z = XLSX.SSF._table[14];
            cell.v = this.datenum(cell.v);
        } else {
            cell.t = 's';
        }
    },
    jhAddHeader: function (rowObj) {
        if (rowObj.hasOwnProperty('labels')){
            var c= 0;
            for (var key in rowObj.labels)
            {
               
                var cellObj = rowObj.labels[key];
                if (this.range.s.r > this.rowCount)
                    this.range.s.r = this.rowCount;
                if (this.range.s.c > c)
                    this.range.s.c = c;
                if (this.range.e.r < this.rowCount)
                    this.range.e.r = this.rowCount;
                if (this.range.e.c < c)
                    this.range.e.c = c;
                   
                var cellText = null;
                cellText = cellObj;
                var cell = {v: cellText};
                var calColWidth = true;
          
                if (calColWidth) {
                    this.cellWidth(cell.v, c);
                }
                if (cell.v === null)
                continue;
               
               
                    
                    var cell_ref = XLSX.utils.encode_cell({c: c, r: this.rowCount});
                    c++;
                    this.setCellDataType(cell);
                    console.log(cell);
                    this.worksheet[cell_ref] = cell;
                
            }
            this.rowCount++;
        }
    },
    
    jhAddRow: function (rowObj) {
    
        if (rowObj.hasOwnProperty('data')){
            for (var r in rowObj.data)
            {
                
                var c= 0;
                for (var key in rowObj.data[r])
                {
              
                    var cellObj = rowObj.data[r][key];
                    if (this.range.s.r > this.rowCount)
                        this.range.s.r = this.rowCount;
                    if (this.range.s.c > c)
                        this.range.s.c = c;
                    if (this.range.e.r < this.rowCount)
                        this.range.e.r = this.rowCount;
                    if (this.range.e.c < c)
                        this.range.e.c = c;
                    
                    var cellText = null;
                    cellText = cellObj;
                    var cell = {v: cellText};
                    var calColWidth = true;
            
                    if (calColWidth) {
                        this.cellWidth(cell.v, c);
                    }
                    if (cell.v === null)
                    continue;
                
                
                        var cell_ref = XLSX.utils.encode_cell({c: c, r: this.rowCount});
                        c++;
                        this.setCellDataType(cell);
                        console.log(cell);
                        this.worksheet[cell_ref] = cell;
                    
                }
                this.rowCount++;
         }
           
        }
       
        this.rowCount++;
    },
    createWorkSheet: function () {
        
        
      
        this.jhAddHeader(this.worksheetObj);
        this.jhAddRow(this.worksheetObj);
        
        this.cellWidthValidate();
        //console.log(this.merges);
        //this.worksheet['!merges'] = [{s: {r: 0, c: 0}, e: {r: 0, c: 4}},{s: {r: 5, c: 0}, e: {r: 6, c: 3}}];//this.merges;
        this.worksheet['!merges'] = this.merges;
        this.worksheet['!cols'] = this.wsColswidth;
        if (this.range.s.c < 10000000)
            this.worksheet['!ref'] = XLSX.utils.encode_range(this.range);
        return this.worksheet;
    },
    s2ab: function (s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i)
            view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    },
    getBlob: function (workbookObj, options) {
        this.init(options);
        var workbook = new Workbook();
        this.worksheetObj=workbookObj;
        this.createWorkSheet();
        workbook.SheetNames.push("Sheet1");
        workbook.Sheets["Sheet1"] = this.worksheet;
        var wbout = XLSX.write(workbook, {bookType: 'xlsx', bookSST: true, type: 'binary'});
        var blobData = new Blob([this.s2ab(wbout)], {type: "application/octet-stream"});
        return blobData;
    },
    export: function (workbookObj, options) {
        saveAs(this.getBlob(workbookObj, options), this.config.fileFullName);
    },
}

function Workbook() {
    if (!(this instanceof Workbook))
        return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

var saveAs = saveAs
  // IE 10+ (native saveAs)
  || (typeof navigator !== "undefined" &&
      navigator.msSaveOrOpenBlob && navigator.msSaveOrOpenBlob.bind(navigator))
  // Everyone else
  || (function(view) {
	"use strict";
	// IE <10 is explicitly unsupported
	if (typeof navigator !== "undefined" &&
	    /MSIE [1-9]\./.test(navigator.userAgent)) {
		return;
	}
	var
		  doc = view.document
		  // only get URL when necessary in case BlobBuilder.js hasn't overridden it yet
		, get_URL = function() {
			return view.URL || view.webkitURL || view;
		}
		, URL = view.URL || view.webkitURL || view
		, save_link = doc.createElementNS("http://www.w3.org/1999/xhtml", "a")
		, can_use_save_link = !view.externalHost && "download" in save_link
		, click = function(node) {
			var event = doc.createEvent("MouseEvents");
			event.initMouseEvent(
				"click", true, false, view, 0, 0, 0, 0, 0
				, false, false, false, false, 0, null
			);
			node.dispatchEvent(event);
		}
		, webkit_req_fs = view.webkitRequestFileSystem
		, req_fs = view.requestFileSystem || webkit_req_fs || view.mozRequestFileSystem
		, throw_outside = function(ex) {
			(view.setImmediate || view.setTimeout)(function() {
				throw ex;
			}, 0);
		}
		, force_saveable_type = "application/octet-stream"
		, fs_min_size = 0
		, deletion_queue = []
		, process_deletion_queue = function() {
			var i = deletion_queue.length;
			while (i--) {
				var file = deletion_queue[i];
				if (typeof file === "string") { // file is an object URL
					URL.revokeObjectURL(file);
				} else { // file is a File
					file.remove();
				}
			}
			deletion_queue.length = 0; // clear queue
		}
		, dispatch = function(filesaver, event_types, event) {
			event_types = [].concat(event_types);
			var i = event_types.length;
			while (i--) {
				var listener = filesaver["on" + event_types[i]];
				if (typeof listener === "function") {
					try {
						listener.call(filesaver, event || filesaver);
					} catch (ex) {
						throw_outside(ex);
					}
				}
			}
		}
		, FileSaver = function(blob, name) {
			// First try a.download, then web filesystem, then object URLs
			var
				  filesaver = this
				, type = blob.type
				, blob_changed = false
				, object_url
				, target_view
				, get_object_url = function() {
					var object_url = get_URL().createObjectURL(blob);
					deletion_queue.push(object_url);
					return object_url;
				}
				, dispatch_all = function() {
					dispatch(filesaver, "writestart progress write writeend".split(" "));
				}
				// on any filesys errors revert to saving with object URLs
				, fs_error = function() {
					// don't create more object URLs than needed
					if (blob_changed || !object_url) {
						object_url = get_object_url(blob);
					}
					if (target_view) {
						target_view.location.href = object_url;
					} else {
						if(navigator.userAgent.match(/7\.[\d\s\.]+Safari/)	// is Safari 7.x
								&& typeof window.FileReader !== "undefined"			// can convert to base64
								&& blob.size <= 1024*1024*150										// file size max 150MB
								) {	
							var reader = new window.FileReader();
							reader.readAsDataURL(blob);
							reader.onloadend = function() {
								var frame = doc.createElement("iframe");
								frame.src = reader.result;
								frame.style.display = "none";
								doc.body.appendChild(frame);
								dispatch_all();
								return;
							}
							filesaver.readyState = filesaver.DONE;
							filesaver.savedAs = filesaver.SAVEDASUNKNOWN;
							return;
						}
						else {
							window.open(object_url, "_blank");
							filesaver.readyState = filesaver.DONE;
							filesaver.savedAs = filesaver.SAVEDASBLOB;
							dispatch_all();
							return;
						}
					}
				}
				, abortable = function(func) {
					return function() {
						if (filesaver.readyState !== filesaver.DONE) {
							return func.apply(this, arguments);
						}
					};
				}
				, create_if_not_found = {create: true, exclusive: false}
				, slice
			;
			filesaver.readyState = filesaver.INIT;
			if (!name) {
				name = "download";
			}
			if (can_use_save_link) {
				object_url = get_object_url(blob);
				// FF for Android has a nasty garbage collection mechanism
				// that turns all objects that are not pure javascript into 'deadObject'
				// this means `doc` and `save_link` are unusable and need to be recreated
				// `view` is usable though:
				doc = view.document;
				save_link = doc.createElementNS("http://www.w3.org/1999/xhtml", "a");
				save_link.href = object_url;
				save_link.download = name;
				var event = doc.createEvent("MouseEvents");
				event.initMouseEvent(
					"click", true, false, view, 0, 0, 0, 0, 0
					, false, false, false, false, 0, null
				);
				save_link.dispatchEvent(event);
				filesaver.readyState = filesaver.DONE;
				filesaver.savedAs = filesaver.SAVEDASBLOB;
				dispatch_all();
				return;
			}
			// Object and web filesystem URLs have a problem saving in Google Chrome when
			// viewed in a tab, so I force save with application/octet-stream
			// http://code.google.com/p/chromium/issues/detail?id=91158
			if (view.chrome && type && type !== force_saveable_type) {
				slice = blob.slice || blob.webkitSlice;
				blob = slice.call(blob, 0, blob.size, force_saveable_type);
				blob_changed = true;
			}
			// Since I can't be sure that the guessed media type will trigger a download
			// in WebKit, I append .download to the filename.
			// https://bugs.webkit.org/show_bug.cgi?id=65440
			if (webkit_req_fs && name !== "download") {
				name += ".download";
			}
			if (type === force_saveable_type || webkit_req_fs) {
				target_view = view;
			}
			if (!req_fs) {
				fs_error();
				return;
			}
			fs_min_size += blob.size;
			req_fs(view.TEMPORARY, fs_min_size, abortable(function(fs) {
				fs.root.getDirectory("saved", create_if_not_found, abortable(function(dir) {
					var save = function() {
						dir.getFile(name, create_if_not_found, abortable(function(file) {
							file.createWriter(abortable(function(writer) {
								writer.onwriteend = function(event) {
									target_view.location.href = file.toURL();
									deletion_queue.push(file);
									filesaver.readyState = filesaver.DONE;
									filesaver.savedAs = filesaver.SAVEDASBLOB;
									dispatch(filesaver, "writeend", event);
								};
								writer.onerror = function() {
									var error = writer.error;
									if (error.code !== error.ABORT_ERR) {
										fs_error();
									}
								};
								"writestart progress write abort".split(" ").forEach(function(event) {
									writer["on" + event] = filesaver["on" + event];
								});
								writer.write(blob);
								filesaver.abort = function() {
									writer.abort();
									filesaver.readyState = filesaver.DONE;
									filesaver.savedAs = filesaver.FAILED;
								};
								filesaver.readyState = filesaver.WRITING;
							}), fs_error);
						}), fs_error);
					};
					dir.getFile(name, {create: false}, abortable(function(file) {
						// delete file if it already exists
						file.remove();
						save();
					}), abortable(function(ex) {
						if (ex.code === ex.NOT_FOUND_ERR) {
							save();
						} else {
							fs_error();
						}
					}));
				}), fs_error);
			}), fs_error);
		}
		, FS_proto = FileSaver.prototype
		, saveAs = function(blob, name) {
			return new FileSaver(blob, name);
		}
	;
	FS_proto.abort = function() {
		var filesaver = this;
		filesaver.readyState = filesaver.DONE;
		filesaver.savedAs = filesaver.FAILED;
		dispatch(filesaver, "abort");
	};
	FS_proto.readyState = FS_proto.INIT = 0;
	FS_proto.WRITING = 1;
	FS_proto.DONE = 2;
	FS_proto.FAILED = -1;
	FS_proto.SAVEDASBLOB = 1;
	FS_proto.SAVEDASURI = 2;
	FS_proto.SAVEDASUNKNOWN = 3;

	FS_proto.error =
	FS_proto.onwritestart =
	FS_proto.onprogress =
	FS_proto.onwrite =
	FS_proto.onabort =
	FS_proto.onerror =
	FS_proto.onwriteend =
		null;

	view.addEventListener("unload", process_deletion_queue, false);
	saveAs.unload = function() {
		process_deletion_queue();
		view.removeEventListener("unload", process_deletion_queue, false);
	};
	return saveAs;
}(
	   typeof self !== "undefined" && self
	|| typeof window !== "undefined" && window
	|| this.content
));
// `self` is undefined in Firefox for Android content script context
// while `this` is nsIContentFrameMessageManager
// with an attribute `content` that corresponds to the window

if (typeof module !== "undefined" && module !== null) {
  module.exports = saveAs;
} else if ((typeof define !== "undefined" && define !== null) && (define.amd != null)) {
  define([], function() {
    return saveAs;
  });
} else if(typeof Meteor !== 'undefined') { // make it available for Meteor
  Meteor.saveAs = saveAs;
}
