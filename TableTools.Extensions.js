if (!!TableTools) {

    /*
    fnDeselectAll() was created as a work-around for a bug that exists in fnSelectNone().
    When fnCreatedCell is used the aoData.nTr is null for any rows not rendered,
    as the HTML hasn't been calculated.  Since fnGetSelected uses nTr to find the 
    affected cells' indexes it fails to find the non-rendered cells which causes 
    the workflow to break.  Given fnSelectAll() does not leverage nTr I think it is
    philosophically wrong to leverage this while attempting to select none / deselect all.
    Note - In non-legacy datatable / tabletools this is all handled much differently.    
    */
    TableTools.prototype.fnDeselectAll = function (a) {
        var selected = this.fnGetSelected(a);
        var thisDT = this;
        $(selected).each(function (i, o) {
            if (o) {
                thisDT.fnDeselect(o);
            }
        });
    }

    TableTools.prototype._fnGetJsonForExport = function (oConfig) {
        var json = {
            columns: [],
            data: []
        };

        var i, iLen, j, jLen;
        var aRow, aData = [], sLoopData = '', arr;
        var dt = this.s.dt, tr, child;
        var regex = new RegExp(oConfig.sFieldBoundary, "g"); /* Do it here for speed */
        var aColumnsInc = this._fnColumnTargets(oConfig.mColumns);
        var bSelectedOnly = (typeof oConfig.bSelectedOnly != 'undefined') ? oConfig.bSelectedOnly : false;

        /* Get columns to be exported */
        for (i = 0, iLen = dt.aoColumns.length ; i < iLen ; i++) {
            if (aColumnsInc[i]) {
                var exportTitle = dt.aoColumns[i].exportTitle;
                var cTitle = (typeof (exportTitle) == "function") ? exportTitle() : exportTitle;
                if (!cTitle) {
                    // If exportTitle isn't set default to sTitle
                    // Sanitize sTitle per DT/TT's standard approach
                    cTitle = this._fnHtmlDecode(dt.aoColumns[i].sTitle.replace(/\n/g, " ").replace(/<.*?>/g, "").replace(/^\s+|\s+$/g, ""));
                }

                var cxWidth = dt.aoColumns[i].exportWidth || 20;

                var column = {
                    title: { name: cTitle },
                    property: { name: cTitle },
                    width: cxWidth
                };

                json.columns.push(column);
            }
        }

        /* Get rows to be exported */
        // Initially get all rows;
        var aDataIndex = dt.aiDisplay;
        // Get selected rows if any
        var aSelected = this.fnGetSelected();
        // If there are selected rows, limit to only those selected
        if (this.s.select.type !== "none" && bSelectedOnly && aSelected.length !== 0) {
            aDataIndex = [];
            for (i = 0, iLen = aSelected.length ; i < iLen ; i++) {
                aDataIndex.push(dt.oInstance.fnGetPosition(aSelected[i]));
            }
        }

        // Iterate through determined rows
        for (j = 0, jLen = aDataIndex.length ; j < jLen ; j++) {
            tr = dt.aoData[aDataIndex[j]].nTr;
            var obj = {};
            var colCount = 0;

            /* Columns */
            for (i = 0, iLen = dt.aoColumns.length ; i < iLen ; i++) {
                if (aColumnsInc[i]) {
                    /* Convert to strings (with small optimisation) */
                    var mTypeData = dt.oApi._fnGetCellData(dt, aDataIndex[j], i, 'display');
                    if (oConfig.fnCellRender) {
                        sLoopData = oConfig.fnCellRender(mTypeData, i, tr, aDataIndex[j]) + "";
                    }
                    else if (typeof mTypeData == "string") {
                        /* Strip newlines, replace img tags with alt attr. and finally strip html... */
                        sLoopData = mTypeData.replace(/\n/g, " ");
                        sLoopData =
						 	sLoopData.replace(/<img.*?\s+alt\s*=\s*(?:"([^"]+)"|'([^']+)'|([^\s>]+)).*?>/gi,
						 		'$1$2$3');
                        sLoopData = sLoopData.replace(/<.*?>/g, "");
                    }
                    else {
                        sLoopData = mTypeData + "";
                    }

                    /* Trim and clean the data */
                    sLoopData = this._fnHtmlDecode(sLoopData.replace(/^\s+/, '').replace(/\s+$/, ''));

                    obj[json.columns[colCount++].property.name] = this._fnBoundData(sLoopData, oConfig.sFieldBoundary, regex)
                }
            }

            json.data.push(obj);
        }

        return json;
    }

    TableTools.BUTTONS.excel = $.extend({}, TableTools.buttonBase, {
        "sButtonText": "Excel",
        "mColumns": "visible",/* "all", "visible", "hidden" or array of column integers */
        "bSelectedOnly": false,
        "sFileName": "Export.xlsx",
        "sheetName": "Export",
        'fnClick': function (nButton, oConfig, flash) {
            var sheetName = (typeof (oConfig.sheetName) == "function") ? oConfig.sheetName() : oConfig.sheetName;

            // Define the sheet to be exported
            var sheet = $.extend({ name: sheetName }, this._fnGetJsonForExport(oConfig));

            var exp = {
                name: oConfig.sFileName,
                sheets: [sheet]
            };

            jk.exporters.excel.process(exp);
        }

    });

}
