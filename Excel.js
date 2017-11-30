/*
 * This is a JavaScript Scratchpad.
 *
 * Enter some JavaScript, then Right Click or choose from the Execute Menu:
 * 1. Run to evaluate the selected text (Ctrl+R),
 * 2. Inspect to bring up an Object Inspector on the result (Ctrl+I), or,
 * 3. Display to insert the result in a comment after the selection. (Ctrl+L)
 */

function getCellDate(value,colDataType,langCode) {
  var cell = {v: value};
	var formatMask = colDataType.formatMask;
	var dateStrict = !( formatMask.search("MMMM") >= 0 || formatMask.search("ddd") >= 0 ); //because of bug https://github.com/moment/moment/issues/4227	
	var langCode = langCode || 'en';	
	var parsedDate = moment(value,formatMask,langCode,dateStrict);		
	var epoch = new Date(1899,11,31);
	
	if(value) {
		if( parsedDate.isValid()) {
			//console.log(value);
			//console.log(formatMask);		  
			cell.t = 'n'; // excel recognizes date as number that have a date format string			
			cell.z = colDataType.formatMaskExcel;						
			cell.v = ((parsedDate.toDate() - epoch) / (24 * 60 * 60 * 1000)) + 1;			// + 1 because excel leap bug
		}	
	 else {
		 console.log("Can't parse date <" + value + "> with format <" + formatMask + "> strict:" + dateStrict);
		 cell.t = 's';	 
	 }
	}
 return cell;
}


function getCellNumber(value,colDataType,decimalSeparator) {	
	var num; 
  var re = new RegExp("[^0123456789" + decimalSeparator +"]","g");	
	var str = "" + value;
	var cell = {v: str};
	
	if(value) {
    str.replace(re, ""); //remove all symbols except digits and decimalSeparator		
		str.replace(decimalSeparator, "."); //change decimalSeparator to JS-decimal separator
		num = parseFloat(str);			
		if( !isNaN(num) ) {
			cell.t = 'n';	  
			cell.z = colDataType.formatMaskExcel;
			cell.v = num;	
			//cell.color = { name: 'accent5', rgb: '4BACC6' };
		} else {
			console.log("Can't parse number <" + value + ">");
			cell.t = 's';	 
			cell.v = value;	
		}
	}
 return cell;
}

function getCellChar(value) {
	return  {v: value,t: 's'};
}


function recalculateRangeAndGetCellAddr(range,colNo,rowNo) {
	// recalculate column range	
	if(range.s.c > colNo) range.s.c = colNo;				
	if(range.e.c < colNo) range.e.c = colNo;
	// recalculate row range		
	if(range.s.r > rowNo) range.s.r = rowNo;
	if(range.e.r < rowNo) range.e.r = rowNo;
	
	return XLSX.utils.encode_cell({c:colNo,r:rowNo});
	
}

function getWorksheet(data,properties) {	
	//console.log(colDataTypesArr);
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	var cellAddr = {}; 
	var cell = {};
	var R,C,I,A; // iterators
  var columnNum; 
	var rowNum = 0;
	var colDataTypesArr = properties.columnsProperties;
	var isControlBreak = properties.haveControlBreaks;
	var rowAdditionalnfo;
	var controlBreakArr = [];
	var startColumn = properties.hasAggregates ? 1 : 0; 
	
	// print headers
	for(I = 0; I < colDataTypesArr.length; I++) {
		columnNum = colDataTypesArr[I].displayOrder;
		if( columnNum < 1000000 ) {						
			cell = getCellChar(colDataTypesArr[I].heading);		  
			cellAddr = recalculateRangeAndGetCellAddr(range,columnNum  + startColumn,rowNum);
			ws[cellAddr] = cell;
		}	
	}
	rowNum++;
	
	//print data
	for(R = 0; R < data.length; R++) { // rows
		rowAdditionalnfo = data[R][data[R].length - 1] || {}; // last record is an object with additional proprties				
		//console.log(isControlBreak + " " + (rowAdditionalnfo.endControlBreak || false));
		
		// display control break		
		if( isControlBreak && (!rowAdditionalnfo.agg) )  { 
			for(C = 0; C < data[R].length; C++) {			//columns 
				columnNum = colDataTypesArr[C].displayOrder;
				if( columnNum > 1000000 ) {			//is control break
					controlBreakArr.push({ displayOrder : columnNum,
															   text : colDataTypesArr[C].heading + " : " + data[R][C]
															 });
				} // end column loop
			}
		  cellAddr = recalculateRangeAndGetCellAddr(range,startColumn,rowNum); 
			// sort contol break columns in display order and convert them to the simple array of strings
			controlBreakArr = controlBreakArr.sort(function(a,b){
				return a.displayOrder - b.displayOrder;
			}).map(function(a){
				return a.text;
			});
			//console.log(controlBreakArr);			
			cell = getCellChar(controlBreakArr.join(", "));				
			ws[cellAddr] = cell;
			rowNum++;							
			controlBreakArr = [];
		} 
		
		// display regular columns		
		for(C = 0; C < data[R].length - 1; C++) {			//columns; -1 because last record is an object with additional proprties
			columnNum = colDataTypesArr[C].displayOrder + startColumn;
			if( columnNum < 1000000 ) {			//display visible columns	
				cellAddr = recalculateRangeAndGetCellAddr(range,columnNum + startColumn,rowNum); 
				if(colDataTypesArr[C].dataType == 'NUMBER') {
					cell = getCellNumber(data[R][C],colDataTypesArr[C],properties.decimalSeparator)
				} else if(colDataTypesArr[C].dataType == 'DATE') {				
					cell = getCellDate(data[R][C],colDataTypesArr[C],properties.langCode);
				} else {
					// string
					cell = getCellChar(data[R][C]);
				}					
				ws[cellAddr] = cell;
			} 
		} // end column loop		
		// aggregations
		if(rowAdditionalnfo.agg) {
			// print name of aggregation in the first column
			if(rowAdditionalnfo.grandTotal) {
			  cell = getCellChar(properties.aggregateLabels[rowAdditionalnfo.agg].overallLabel);	
			} else {
				cell = getCellChar(properties.aggregateLabels[rowAdditionalnfo.agg].label);	
			}
			cellAddr = recalculateRangeAndGetCellAddr(range,0,rowNum); 
			ws[cellAddr] = cell;
		} else {
		  isControlBreak = rowAdditionalnfo.endControlBreak || false;			
		}
		rowNum++;				
	} // end row loop	
	
	
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range); // to do: clarify	
	return ws;
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

function getRows(iGrid) {
  var rows = [];
	console.log(iGrid.getDataElements());
	iGrid.model.forEach(function (row_in) { 		
		//console.log(row_in);		
		rows.push(row_in);  
  })	
  return rows;	
}

function getPreparedIGProperties(columns,propertiesFromPlugin) {
	// all regular columns have displayOrder < 1000000 
	// all hidden columns have displayOrder = 1000000
	// all control breakcolumns have displayOrder > 1000000
	
	var I; //iterator
	var currColumnNo = 0;
	var controlBreakColumnNo = 1000001;
	var colProp = propertiesFromPlugin.column_properties;
	var haveControlBreaks = false;
	
	// assign to the each data column a corresponding column number in excel 
	//
	// first sort columns in display order
	var displayInColumnArr = columns.map(function(a) { 		
		haveControlBreaks = (a.controlBreakIndex || haveControlBreaks) ? true : false;
		return  {index : a.index, 
						 displayOrder : a.hidden ? (1000000 + (a.controlBreakIndex || 0)): a.seq, //to place hidden and control break columns at the end after sorting
						 heading: a.heading,
						 headingAlignment: a.headingAlignment,
						 width: a.width,
						 dataType    : "VARCHAR2",
						 formatMask : "",
						 formatMaskExcel : "",
						 name : "",
						 id : a.id,
						 controlBreakIndex : a.controlBreakIndex,
						 hidden : a.hidden
						};		
	 }).sort(function(a, b) { 
    return a.displayOrder - b.displayOrder;
   });	
	
	// second renumerate display order - skip hidden columns
	for(I = 0; I < displayInColumnArr.length; I++) { // regular row
		if(displayInColumnArr[I].displayOrder < 1000000) {
		  displayInColumnArr[I].displayOrder = currColumnNo;
			currColumnNo++;
		}	else if (displayInColumnArr[I].controlBreakIndex) { // control break row
			displayInColumnArr[I].displayOrder = controlBreakColumnNo;
			controlBreakColumnNo++;
		}
	}		
  // third sort columns in the data order
	displayInColumnArr.sort(function(a, b) { 
    return a.index - b.index;
   });
	
	// add additional data from server to the colHeader (map by column id)
	displayInColumnArr.forEach(function(val,index) {
		var b; // iterator
		for(b = 0; b < colProp.length; b++) {			
			if(val.id == colProp[b].COLUMN_ID) {				
				val.dataType = colProp[b].DATA_TYPE;
				val.formatMask = colProp[b].DATE_FORMAT_MASK_JS;
				val.formatMaskExcel = colProp[b].DATE_FORMAT_MASK_EXCEL;		
				break;
			}				
		}		
	});
	
	return { columnsProperties : displayInColumnArr,
					 decimalSeparator: propertiesFromPlugin.decimal_seperator,
					 langCode : propertiesFromPlugin.lang_code,
					 haveControlBreaks : haveControlBreaks,
					 hasAggregates : false,
					 aggregateLabels : {}
				 };
}


function hasAggregates(rows) {
  // if aggregates exists last row always shows aggregates
	var lastRecord = rows[rows.length -1] || [];
	var rowAdditionalnfo = lastRecord[lastRecord.length -1] || {};
	return rowAdditionalnfo.agg ? true : false;
}

/* main code */

function main(propertiesFromPlugin) {
	var iGrid = apex.region("IG").widget();
	var currentIGView = iGrid.interactiveGrid("getCurrentView");
	var columnPropertiesFromIG = currentIGView.view$.grid("getColumns");
	var	ws_name = currentIGView.model.name;
	var wb = new Workbook(); 
	var ws;
	var wbout;
	var rows = getRows(currentIGView);	
	
	//console.log(iGrid);
	//console.log(rows);
	//console.log(columnPropertiesFromIG);
	//console.log(propertiesFromPlugin);	
	
	var properties = getPreparedIGProperties(columnPropertiesFromIG,propertiesFromPlugin);  
	properties.hasAggregates = hasAggregates(rows);
	properties.aggregateLabels = iGrid.interactiveGrid("getViews").grid.aggregateLabels;
	//console.log(properties);
	
	
	
	ws = getWorksheet(rows,properties); 
	// add worksheet to workbook 
	//return;
	wb.SheetNames.push(ws_name);
	wb.Sheets[ws_name] = ws;
	wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});

	saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "test.xlsx");

}

getColumnsProperties(main);


