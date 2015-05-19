function ExcelExporter(){
	// initialization
	if(!(jk.hasActiveX)) {
		var fileRefs = [
			"./jk-exporters/excel-exporter/support/xlsx.core.min.js"
		];

		includeSupportRef(fileRefs);
	}

	// Export
	this.process = function(data){
		if(jk.hasActiveX){
			createWorkbookIE(data);
		}
		else {
			createWorkbook(data);
		}
	}

	// IE Specific Functionality
	function createWorkbookIE(exp){
		// Start Excel
		var objExcel;
		objExcel = new ActiveXObject("Excel.Application");
		objExcel.Visible = false;

		// Create Workbook with initial worksheet
		var wb = objExcel.Workbooks.Add(1);

		// Add sheets to workbook
		addWorksheetsIE(wb, exp.sheets);

		// Delete initial worksheet and make the application instance visible
		wb.Worksheets("Sheet1").Delete();
		objExcel.Visible = true;

		// Save File
		if(exp.name) {
			var fileName = (exp.destination || "../Downloads") + "/" + exp.name + ".xlsx";
			wb.SaveAs(fileName);
		}
	}

	function addWorksheetsIE(wb, sheets){
		$(sheets).each(function(i, sheet){
			var ws = wb.Sheets.Add();
			ws.Name = sheet.name;

			populateWorksheetIE(ws, sheet.data, sheet.columns);
		});
	}

	function populateWorksheetIE(ws, data, cols){
		$(cols).each(function(iC, col){
			var title = col.title.name || col.property.name;
			setCellValueIE(0, iC, title, ws);

			$(data).each(function(iR, obj){
				var val = resolvePropertyValue(col.property, obj);

				// Add 1 to row to make up for title row
				setCellValueIE(iR+1, iC, val, ws);
			});
		});
	}

	function setCellValueIE(iR, iC, value, ws){
		if(value) {
			var range = ws.Cells(iR+1, iC+1);
			range.Value = value;
		}
	}



	// Non-IE Specific Functionality
	function createWorkbook(exp){
		var wb = new Workbook();

		addWorksheets(wb, exp.sheets);

		var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: true, type: 'binary'});
		var fileName = exp.name + ".xlsx";

		saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), fileName);
	}

	function addWorksheets(wb, sheets){
		$(sheets).each(function(i,sheet){
			var ws_name = sheet.name;
			wb.SheetNames.push(ws_name);

			var ws = {};
			populateWorksheet(ws, sheet.data, sheet.columns);

			wb.Sheets[ws_name] = ws;
		});
	}

	function populateWorksheet(ws, data, cols){
		var range = {s: {c:0, r:0}, e: {c:cols.length, r:data.length+1 }};

		$(cols).each(function(iC, col){
			var title = col.title.name || col.property.name;
			setCellValue(0, iC, title, ws);

			$(data).each(function(iR, obj){
				var val = resolvePropertyValue(col.property, obj);
				setCellValue(iR+1, iC, val, ws);
			});
		});

		ws['!ref'] = XLSX.utils.encode_range(range);
	}

	function setCellValue(iR, iC, value, ws){
		if (value) {
			var cell_ref = XLSX.utils.encode_cell({c: iC, r: iR});
			var cell = createCell(value);

			ws[cell_ref] = cell;
		}
	}

	function createCell(value){
		var cell = {v: value, t:'s'};

		// Set type
		switch(Object.prototype.toString.call(value)){
			case '[object Number]':
				cell.t = 'n';
				break;

			case '[object Boolean]':
				cell.t = 'b';
				break;

			case '[object String]':
				cell.t = 's';
				break;

			case '[object Date]':
				cell.t='n';
				cell.z = XLSX.SSF._table[14];
				cell.v = datenum(value);
				break;
		}

		return cell;
	}


	function resolvePropertyValue(property, object){
		var prop = property.name;
		var resolve = property.resolve;
		var val = null;

		if(prop){
			val = object[prop];
		}
		else if(typeof(resolve) == "function"){
			val = resolve(object);
		}

		return val;
	}

	function datenum(v, date1904) {
		if(date1904) v+=1462;
		var epoch = Date.parse(v);
		return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
	}

	function Workbook() {
		if(!(this instanceof Workbook)) return new Workbook();
		this.SheetNames = [];
		this.Sheets = {};
	}

	function s2ab(s) {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	}

	return this;
}

Exporter.prototype.excel = new ExcelExporter();
