var fs = require('fs');
var _ = require('lodash');
var LZString = require('lz-string');
var Q = require('q');
var appConfigData = require('../../../../appConfig/appConfigData').data;

var ExcelBuilder = {
    Drawings: require('../Excel/Drawings'),
    Drawing: require('../Excel/Drawing/index'),
    Pane: require('../Excel/Pane'),
    Paths: require('../Excel/Paths'),
    Positioning: require('../Excel/Positioning'),
	Picture: require('../Excel/Drawing/Picture'),
    RelationshipManager: require('../Excel/RelationshipManager'),
    SharedStrings: require('../Excel/SharedStrings'),
    SheetProtection: require('../Excel/SheetProtection'),
    SheetView: require('../Excel/SheetView'),
    StyleSheet: require('../Excel/StyleSheet'),
    Table: require('../Excel/Table'),
    util: require('../Excel/util'),
    Workbook: require('../Excel/Workbook'),
    Worksheet: require('../Excel/Worksheet'),
    WorksheetExportWorker: require('../Excel/WorksheetExportWorker'),
    XMLDOM: require('../Excel/XMLDOM'),
    ZipWorker: require('../Excel/ZipWorker'),
    Builder: require('../excel-builder'),
};


var ExcelReport = function(config,fileName,action) {
		
	try {
		this.fileName = fileName;
		this.path = appConfigData.excelReports.excelReportsDirLocation+this.fileName;
		this.config = LZString.decompressFromEncodedURIComponent(config);
		this.action = action;
		switch(this.action) {
			
			case 'start' : 
							var that = this;
							Q.nfcall(fs.writeFile,this.path, this.config, 'utf8').then(function(data) {
							}).fail(function(err){
								throw 'EB-excelReport: start '+err;
							});
							break;
			case 'process' : 
							var that = this;
							Q.nfcall(fs.appendFile, this.path,this.config,'utf8').then(function(data) {
							}).fail(function(err){
								throw 'EB-excelReport: process '+err;
							});
							break;	
			case 'finish' : 
							var that = this;
							Q.nfcall(fs.appendFile, this.path,this.config,'utf8').then(function(data) {
								var json = that.readConfig(that.path).then(function(json){
									that.initialize(json);
									that.validate();
									that.parse();
									that.generate();
								}).fail(function(err){
									throw 'EB-excelReport: finish-readConfig '+err;
								});
							}).fail(function(err){
								throw 'EB-excelReport: finish '+err;
							});
							break;
							
			case '': 
							break;	
		}
		
	} catch(e) {
		throw 'Error: EB-ExcelReport - '+e;
	}
};

_.extend(ExcelReport.prototype, {
	
	download: function(req,res) {
		var that = this;
		var appLogger = "";
		appLogger = req.app.locals.logging;
		var ip = req.header('x-iisnode-remote_addr') || req.connection.remoteAddress;
		appLogger.logMessage('Info', ip, 'download Request', 'Request received in EB for downloading the file '+this.fileName, 'EB-ExcelReport', ip, '/POST');
		
		Q.nfcall(fs.unlink, this.path.replace('.xlsx','.json') ).then(function(success){
			appLogger.logMessage('Info', ip, 'download Request', 'Successfully deleted the JSON file '+that.fileName.replace('.xlsx','.json'), 'EB-ExcelReport', ip, '/POST');
		} ).fail(function (err) {
			throw 'Error: EB-download '+e;
		});
		Q.nfcall(fs.readFile,this.path).then(function(data){
			res.set({
				'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				'Content-Disposition': 'attachment; filename='+that.fileName
			});
			Q.nfcall(fs.unlink, that.path).then(function(success){
				appLogger.logMessage('Info', ip, 'download Request', 'Successfully deleted the xlsx file '+that.fileName, 'EB-ExcelReport', ip, '/POST');
			} ).fail(function (err) {
				throw 'Error: EB-download '+e;
			});
			appLogger.logMessage('Info', ip, 'download Request', 'Successfully called the download from EB '+that.fileName, 'EB-ExcelReport', ip, '/POST');
			res.status(200);
			res.send(data);
		}).fail(function(err){
			res.status(500);
			appLogger.logMessage('Error', ip, 'download Request',that.fileName+' '+err,'EB-ExcelReport', ip, '/POST');
			res.render('error', {message: 'Looks something went wrong! While downloading, Please try after some time.', error:{status:''}} );
		});	
		
	},
	
	readConfig: function (path) {
		var json = {};
		return Q.nfcall(fs.readFile,path,'utf8').then(function(data){
			return JSON.parse(data);
		}).fail(function(err){
			throw 'Error: EB-readConfig '+err;
		});
	},
	
	getFile: function() {
		var file;
		Q.nfcall(fs.access,this.path).then(function(){
			file = this.path;
		}).fail(function(err){
			throw 'Error: EB-getFile '+err;
		});
		return Q.resolve(file);
	},
	
	initialize: function(config) {
	
		this.config = config || {};
		this.fileData = '';
		if(typeof(config.EXCEL_STARTPOSITION) != 'object') {
			this.config.EXCEL_STARTPOSITION = [1,1];  /* Excel cell position to begin  */
		}
		
		if(!config.EXCEL_TABLECOLS) {
			this.config.EXCEL_TABLECOLS = 2;	/* Number of table columns to be displayed in a worksheet */
		}
		
		if(!config.EXCEL_COLBUFFER) {
			this.config.EXCEL_COLBUFFER = 1; /* Number of row space b/t the table */
		}
		
		if(!config.EXCEL_ROWBUFFER) {
			this.config.EXCEL_ROWBUFFER = 1; /* Number of row space b/t the table */
		}
		
		if(!config.EXCEL_THEMESTYLE) {
			this.config.EXCEL_THEMESTYLE = 'TableStyleMedium1'; /* theme style of the table */
		}
		
		if(!config.EXCEL_IMGNAME) {
			this.config.EXCEL_IMGNAME = 'chart.jpg'; /* image name of the picture displayed */
		}
		
		if(!config.EXCEL_MEDIATYPE) {
			this.config.EXCEL_MEDIATYPE = 'image'; /* media type to be included in a worksheet */
		}
		if(!config.EXCEL_WORKSHEETNAME) {
			this.config.EXCEL_WORKSHEETNAME = ['Report'];
		}
		
		if(!config.EXCEL_WORKSHEETHEADER) {
			this.config.EXCEL_WORKSHEETHEADER = [[ '',{bold: true, text: "Report"}, '' ]];
		}
		
		if(!config.EXCEL_WORKSHEETFOOTER) {
			this.config.EXCEL_WORKSHEETFOOTER = [['', '', 'Page &P of &N']];
		}
		
		if(!config.EXCEL_ORIENTATION) {
			this.config.EXCEL_ORIENTATION = 'landscape';
		}
		
		if(!config.EXCEL_FILENAME) {
			this.config.EXCEL_FILENAME = 'dataExport.xlsx'; /* file name of the excel to be saved */
		}
		
		if(typeof(config.DATA) != 'object') {
			this.config.DATA = [[[[]],[[]]]]; /* array to hold the table/data */
		} 
		
		if(!config.EXCEL_IMAGEXCELL) {
			this.config.EXCEL_IMAGEXCELL = ExcelBuilder.Positioning.pixelsToXCellPos(600) || 0 ; /* number of cells the image to be drawn */
		} else {
			this.config.EXCEL_IMAGEXCELL = ExcelBuilder.Positioning.pixelsToXCellPos(+config.EXCEL_IMAGEXCELL);
		}
		
		if(!config.EXCEL_IMAGEYCELL) {
			this.config.EXCEL_IMAGEYCELL = ExcelBuilder.Positioning.pixelsToYCellPos(250) || 0; /* number of cells the image to be drawn */
		} else {
			this.config.EXCEL_IMAGEYCELL = ExcelBuilder.Positioning.pixelsToXCellPos(+config.EXCEL_IMAGEYCELL);
		}
		
		if(!config.EXCEL_APPLYLOCALSTYLE) {
			this.config.EXCEL_APPLYLOCALSTYLE = 0;
		}
		
		if(!config.EXCEL_DATAHEADERSTYLE) {
			this.config.EXCEL_DATAHEADERSTYLE = {
					font: { bold: true, underline: false, color: {theme: 1}, size: 11},
					fill:  { type: 'pattern',patternType: 'gray125',fgColor: 'dce6f1', bgColor: 'dce6f1'}, 
					alignment: { horizontal: 'left'},
					border: { top : {style: 'thin', color: '000000'}, bottom : {style: 'thin', color: '000000'}, left : {style: 'thin', color: '000000'}, right : {style: 'thin', color: '000000'} },
					format: '0.000'
			}
		}

		if(!config.EXCEL_DATACELLSTYLE) {
			this.config.EXCEL_DATACELLSTYLE = {
					font: { bold: false, underline: false, color: {theme: 1}, size: 8},
					alignment: { horizontal: 'left'},
					border: { top : {style: 'thin', color: '000000'}, bottom : {style: 'thin', color: '000000'}, left : {style: 'thin', color: '000000'}, right : {style: 'thin', color: '000000'} },
					format: '0.000'
			}
		
		}
		
		if(!config.EXCEL_DATANUMSTYLE) {
			this.config.EXCEL_DATANUMSTYLE = {
					font: { bold: false, underline: false, color: {theme: 1}, size: 8},
					alignment: { horizontal: 'right'},
					border: { top : {style: 'thin', color: '000000'}, bottom : {style: 'thin', color: '000000'}, left : {style: 'thin', color: '000000'}, right : {style: 'thin', color: '000000'} },
					format: '#,##0.000'
			}
		
		}
		
		if(!config.EXCEL_DATAHEADERNUMSTYLE) {
			this.config.EXCEL_DATAHEADERNUMSTYLE = {
					font: { bold: true, underline: false, color: {theme: 1}, size: 11},
					fill:  { type: 'pattern',patternType: 'gray125',fgColor: 'dce6f1', bgColor: 'dce6f1'}, 
					alignment: { horizontal: 'right'},
					border: { top : {style: 'thin', color: '000000'}, bottom : {style: 'thin', color: '000000'}, left : {style: 'thin', color: '000000'}, right : {style: 'thin', color: '000000'} },
					format: '#,##0.000'
			}
		}
		
		if(!config.EXCEL_DATATITLESTYLE) {
			this.config.EXCEL_DATATITLESTYLE = {
					font: { bold: true, underline: false, color: {theme: 1}, size: 16},
					alignment: { horizontal: 'center'}	
			}
		}
		
		this.config.EXCEL_MAXCOL = 0; /* number of columns in a table - default 0 */
		this.config.EXCEL_MAXROW = 0; /* number of rows in a table - default 0 */
		this.config.EXCEL_POS = []; /* array contains the excel cell position of the table */
		this.config.EXCEL_REFTABLE = []; /* array contains the excel table position */
		
		this.workbook = new ExcelBuilder.Workbook();
		var sheetView = new ExcelBuilder.SheetView;
		this.drawings = new ExcelBuilder.Drawings();
		this.stylesheet = this.workbook.getStyleSheet();
		this.predefinedStyles = {
			header: this.stylesheet.createFormat(this.config.EXCEL_DATAHEADERSTYLE),
			cell: this.stylesheet.createFormat(this.config.EXCEL_DATACELLSTYLE),
			number: this.stylesheet.createFormat(this.config.EXCEL_DATANUMSTYLE),
			headerNumber: this.stylesheet.createFormat(this.config.EXCEL_DATAHEADERNUMSTYLE),
			title: this.stylesheet.createFormat(this.config.EXCEL_DATATITLESTYLE)
		};
		this.dynamicStyles = {};
		this.worksheet = [];
		this.styles = {	header:this.predefinedStyles.header.id, 
						cell:this.predefinedStyles.cell.id, 
						number:this.predefinedStyles.number.id,
						headerNumber: this.predefinedStyles.headerNumber.id,
						title: this.predefinedStyles.title.id
					  };
		
	},
	validate: function() {
		
		if(this.config.EXCEL_STARTPOSITION) {
			if(_.isNaN(+this.config.EXCEL_STARTPOSITION[0]) || _.isNaN(+this.config.EXCEL_STARTPOSITION[1])) {
				throw 'Error: Please change the value of config EXCEL_STARTPOSITION, only numeric values are supported ';
			}
			
			if((this.config.EXCEL_STARTPOSITION[0] <=0) || (this.config.EXCEL_STARTPOSITION[1] <=0)) {
				throw 'Error: Please change the value of config EXCEL_STARTPOSITION, the position of excel cell should not be zero or below ';
			}
		}
		
		if(_.isNaN(+this.config.EXCEL_TABLECOLS)) {
			throw 'Error: Please change the value of config EXCEL_TABLECOLS, only numeric values are supported ';
		}
			
		if(this.config.EXCEL_TABLECOLS > 2) {
			throw 'Error: Please change the value of config EXCEL_TABLECOLS, the library only supports upto 2 columns';
		}
		
		if(this.config.EXCEL_TABLECOLS <= 0) {
			throw 'Error: Please change the value of config EXCEL_TABLECOLS, the library wont support the config value zero or below';
		}
		
		if(_.isNaN(+this.config.EXCEL_COLBUFFER)) {
			throw 'Error: Please change the value of config EXCEL_COLBUFFER, only numeric values are supported ';
		}
			
		if(_.isNaN(+this.config.EXCEL_ROWBUFFER)) {
			throw 'Error: Please change the value of config EXCEL_ROWBUFFER, only numeric values are supported ';
		}
		
		if(!_.isArray(this.config.DATA)) {
			throw 'Error: the value of config DATA, needs to be an array. format:[[[]]] ';
		}
		
		if(_.isArray(this.config.DATA)) {
			if(!_.isObject(this.config.DATA[0][0]))	{
				throw "Error: the value of config DATA, needs to be an array. format:[[{ 'htmlTable / table / image': [[{value: 'cusip', header: true},{value: 'date', header: true},['646136X22','10/26/2016 09:44:59 AM']]]}]] ";
			}	 
		}
		
		if(this.config.EXCEL_IMAGEXCELL) {
			if(_.isNaN(+this.config.EXCEL_IMAGEXCELL)) {
				throw 'Error: Please change the value of config EXCEL_IMAGEXCELL, only numeric values are supported ';
			}
		}
		
		if(this.config.EXCEL_IMAGEYCELL) {
			if(_.isNaN(+this.config.EXCEL_IMAGEYCELL)) {
				throw 'Error: Please change the value of config EXCEL_IMAGEYCELL, only numeric values are supported ';
			}
		}
		
		if(this.config.EXCEL_WORKSHEETNAME.length < this.config.DATA.length) {
			throw "Error: There are not enough number of sheets to write data to excel. Please assign more worksheets by assigning the values to config.EXCEL_WORKSHEETNAME";
		}
		
	},
	applyStyles: function() {
		
		if(!this.config.EXCEL_APPLYLOCALSTYLE) {
			return;
		}
		
		try {
			var n = 1, style = '';
			for(var i=0; i < this.config.DATA.length; i++ ) {
				for(var j=0; j < this.config.DATA[i].length; j++ ) {
					for (var property in this.config.DATA[i][j]){
						for ( var k=0; k < this.config.DATA[i][j][property].length;k++ ) {
							for(var l=0; l < this.config.DATA[i][j][property][k].length; l++ ) {
								if(typeof this.config.DATA[i][j][property][k][l] == 'object') {
									if(this.config.DATA[i][j][property][k][l] && typeof this.config.DATA[i][j][property][k][l]['metadata'] == 'object' ) {
										if(typeof this.config.DATA[i][j][property][k][l]['metadata']['style'] != 'undefined' ) {
											style = 'style'+n;
											this.dynamicStyles[style] = this.stylesheet.createFormat(this.config.DATA[i][j][property][k][l]['metadata']['style']);
											this.config.DATA[i][j][property][k][l]['metadata']['style'] = this.dynamicStyles[style].id;
											this.styles[style] = this.dynamicStyles[style].id;
											n++;
										}
									}
								}
							}	
						}
					}
				}
			}
		} catch(e) {
			throw "Error EB-applyStyles - "+e;
		}
	},
	addDatatoExcel: function(flag,wIndex,index) {
		var table,imageData,imageFileData,imageName,picRef,image;
		
		try {
			this.worksheet[wIndex].setPosXY(this.config.EXCEL_POS[index]);
			
			switch(flag) {
			
				case 'table':
							this.worksheet[wIndex].setData(this.config.DATA[wIndex][index]);
							if(this.config.DATA[wIndex][index][0].length > 1 && typeof this.config.EXCEL_REFTABLE[index] != 'undefined' ) {
								table = new ExcelBuilder.Table();
								table.styleInfo.themeStyle = this.config.EXCEL_THEMESTYLE;
								table.setReferenceRange(this.config.EXCEL_REFTABLE[index][0],this.config.EXCEL_REFTABLE[index][1]);
								table.setTableColumns(this.config.DATA[wIndex][index][1]);
								this.worksheet[wIndex].addTable(table);
								this.workbook.addTable(table);
							}
							break;
						
				case 'image':
							imageData = this.config.DATA[wIndex][index][0][0]['IMG'];
							imageFileData = this.config.EXCEL_IMGNAME.split('.');
							imageName = imageFileData[0] + index +'.'+ imageFileData[1];
							picRef = this.workbook.addMedia(this.config.EXCEL_MEDIATYPE,imageName,imageData);
							image = new ExcelBuilder.Picture();
							image.createAnchor('twoCellAnchor',{
								from:{
										x: +this.config.EXCEL_REFTABLE[index][0][0],
										y: +this.config.EXCEL_REFTABLE[index][0][1]	
									 },
			
								to: {
										x: +this.config.EXCEL_REFTABLE[index][1][0],
										y: +this.config.EXCEL_REFTABLE[index][1][1]
									}
							});
							image.setMedia(picRef);
							this.worksheet[wIndex].setData(this.config.DATA[wIndex][index][0][1]);
							this.drawings.addDrawing(image);
							break;
				
				case 'htmlTable' :
							this.worksheet[wIndex].setData(this.config.DATA[wIndex][index]);
							this.worksheet[wIndex].setColumns(this.config.DATA[wIndex][index][0]);
							break;
							
				case 'default' || '' :
							break;
				default:
						 break;
			
			}
		} catch(e) {
			throw 'Error: EB-addDatatoExcel - '+e;
		}
	},

	getExcelCellPostion : function(data) {
	
		var tRows = 1,tCols = 1,maxTableRows;
		var imageXCell = this.config.EXCEL_IMAGEXCELL;
		var imageYCell = this.config.EXCEL_IMAGEYCELL;
		var maxCols = this.config.EXCEL_MAXCOL;
		var rowBuffer = this.config.EXCEL_ROWBUFFER;
		var colBuffer = this.config.EXCEL_COLBUFFER;
		var excelStartPos = this.config.EXCEL_STARTPOSITION;
		var tableCols= this.config.EXCEL_TABLECOLS;
		var rtn = {};
		rtn.table = [];
		rtn.pos = [];
		rtn.mergedCells = [];
		
		try {
			for(var i=0; i< data.length; i++) { 
				switch(i%tableCols) {
						case 0: 
							/* Calculate the number of max rows of each table row */
							maxTableRows = 0;
							for(var j=1; j<=tableCols && i!=0; j++) {
								if(data[i-j][0][0] && data[i-j][0][0]['IMG'] && maxTableRows < imageYCell) {
									maxTableRows = imageYCell+1;
								} else if(maxTableRows < data[i-j].length) {
									maxTableRows = data[i-j].length;
								}
							}
							if (i > 1 && tableCols > 1 ) { 
								tRows = tRows+maxTableRows+rowBuffer;
							} else if(tableCols == 1){
								if(i==0) {
									tRows = maxTableRows+rowBuffer;
								} else {
									if(i > 1) { tRows = tRows+maxTableRows+rowBuffer } else { tRows = maxTableRows+rowBuffer+1; }
								}
							}
								
							/* row */
							if(data[i][0].length > 1) {
								if(data[i][0][0] && data[i][0][0]['IMG']) {
								/* image */
									if(i==0) {
										rtn.table.push([[excelStartPos[0]-1,excelStartPos[1]],[+excelStartPos[0]-1+imageXCell,+excelStartPos[1]-1+imageYCell]]);
										;
										rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0],excelStartPos[1]), ExcelBuilder.util.positionToLetterRef(+excelStartPos[0]-1+imageXCell,excelStartPos[0])]);
										
									} else {
										
										rtn.table.push([[excelStartPos[0]-1,excelStartPos[1]-1+tRows],[excelStartPos[0]-1+imageXCell,+imageYCell+excelStartPos[1]-1+tRows]]);
										
										rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0],excelStartPos[1]+tRows-1), ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+imageXCell,excelStartPos[1]+tRows-1)]);
									}
								} else {
								/* table */
									
									if(data[i] && data[i][0].length > 1) {
										if(i==0) {
											//rtn.table.push([[1,2],[6,4]])
											rtn.table.push([[excelStartPos[0],excelStartPos[1]+1],[excelStartPos[0]-1+data[i][0].length,excelStartPos[1]-1+data[i].length]]);
											
											data[i] && data[i][0].length > 1 ? rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0],excelStartPos[1]), ExcelBuilder.util.positionToLetterRef(excelStartPos[0]+data[i][0].length-1,excelStartPos[0])]) : '';
										} else {
											
											rtn.table.push([[excelStartPos[0],excelStartPos[1]+tRows],[+excelStartPos[0]-1+data[i][0].length,excelStartPos[1]-1+tRows+data[i].length-1]]);
											
											data[i] && data[i][0].length > 1 ? rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0],excelStartPos[1]+tRows-1), ExcelBuilder.util.positionToLetterRef(+excelStartPos[0]-1+data[i][0].length,excelStartPos[1]+tRows-1)]) : '';
											
										} 
									} else {
										rtn.table.push([[0,0],[0,0]]);
									}
								}
								/* Excel cell position */
								if(i == 0) {
									rtn.pos.push(excelStartPos);
								} else {
									rtn.pos.push([excelStartPos[0],+tRows+excelStartPos[1]-1]); 
								}
							} else {
								/* default values */
								rtn.pos.push('');
								rtn.table.push([[0,0],[0,0]]);
							}
							break;
					case 1: 
							/* column */
							
							tCols = +maxCols+colBuffer+1;
							if(data[i][0].length > 1) {
								if(data[i][0][0] && data[i][0][0]['IMG']) {
								/* image */
									if(i==1) {
										rtn.table.push([[excelStartPos[0]-1+tCols-1,excelStartPos[1]],[excelStartPos[0]-1+tCols+imageXCell,excelStartPos[1]-1+imageYCell]]);
										
										rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0]+tCols-1,excelStartPos[1]), ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+tCols+imageXCell,excelStartPos[1])]);
										
									} else {
										rtn.table.push([[excelStartPos[0]-1+tCols-1,excelStartPos[1]-1+tRows],[excelStartPos[0]-1+tCols+imageXCell,+excelStartPos[1]-1+imageYCell+tRows]]);
										
										rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0]+tCols-1,excelStartPos[1]+tRows-1), ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+tCols+imageXCell,excelStartPos[1]+tRows-1)]);
									}
						
								} else {
								/* table */
									if(data[i] && data[i][0].length > 1) {
										if(i==1) {
											
											rtn.table.push([[excelStartPos[0]-1+tCols,excelStartPos[1]+1],[excelStartPos[0]-1+tCols+data[i][0].length-1,data[i].length+excelStartPos[1]-1]]);
											
											data[i][0].length > 1 ? rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+tCols,excelStartPos[1]), ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+tCols+data[i][0].length-1,excelStartPos[1])]) : '';
											
										} else {
											
											rtn.table.push([[excelStartPos[0]-1+tCols,excelStartPos[1]+tRows],[excelStartPos[0]-1+tCols+data[i][0].length-1,+excelStartPos[1]-1+tRows+data[i].length-1]]);
											
											data[i] && 	data[i][0].length > 1 ? rtn.mergedCells.push([ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+tCols,excelStartPos[1]+tRows-1), ExcelBuilder.util.positionToLetterRef(excelStartPos[0]-1+tCols+data[i][0].length-1,excelStartPos[1]+tRows-1)]) : '';
										} 
									} else {
										rtn.table.push([[0,0],[0,0]]);
									}
								}
								if(i == 1) {
									rtn.pos.push([excelStartPos[0]-1+tCols,excelStartPos[1]]) ;
								} else {
									rtn.pos.push([excelStartPos[0]-1+tCols,excelStartPos[1]-1+tRows]);
									
								}
							} else {
								rtn.pos.push('');
								rtn.table.push([[0,0],[0,0]]);
							}
							break;
							
					default: 
							 break;
				}
			}
		} catch(e) {
			throw 'Error: EB-getExcelCellPostion - '+e;
		}
		
		return rtn;
	},

	getTblMaxColnRow : function(data) {
		var maxCols, maxRows;
		var imageXCell = this.config.EXCEL_IMAGEXCELL;
		var imageYCell = this.config.EXCEL_IMAGEYCELL;
		var tableCols= this.config.EXCEL_TABLECOLS	

		try {
			/* Get the initial values */
			if(data[0] && data[0][0]) {
				if(data[0][0][0] && data[0][0][0]['IMG']) {
					maxCols = imageXCell;
					maxRows = imageYCell;
				} else {
					maxCols = data[0][0].length;
					maxRows = data[0].length;
				}
			}
			/* Check and update the maximum columns and rows */
			for(var i=tableCols; i<data.length; i=i+tableCols) {
				if(data[i] && data[i][0]) {
					if(data[i][0][0] && data[i][0][0]['IMG'] && maxCols < imageXCell) {
						maxCols = imageXCell;
					} else if(maxCols < data[i][0].length) {
						maxCols = data[i][0].length;
					}
				}
			}
			/* not required most of case in row, since the size of the data varies in each row */
			for(var i=1; i<data.length; i=i+tableCols) {
				if(data[i] && data[i][0]) {
					if(data[i][0][0] && data[i][0][0]['IMG'] && maxRows < imageYCell) {
						maxRows = imageYCell;
					} else if(maxRows < data[i].length) {
						maxRows = data[i].length;
					}
				}
			}
		} catch(e) {
			throw 'Error: EB-getTblMaxColnRow - '+e;
		}
		return [maxCols,maxRows];
	},
	
	getExcelViewType: function(data) {
		var rtn = [];
		var type = '';
		try {
			Object.keys(data).forEach(function(key,value){
				if(data[key]['htmlTable']) {
					type = 'htmlTable';
				} else 
				if(data[key]['table']) {
					type = 'table';
				} else 
				if(data[key]['image']) {
					type = 'image';
				} else {
					type = '';
				}
				rtn.push(type);
				
				data[key] = type ? data[key][type]: [[]];
			});
		} catch(e) {
			throw 'Error: EB-getExcelViewType - '+e;
		}
		
		return rtn;
	},
	parse: function() {
		
		try {
			this.applyStyles();
			
			/* Excel needs the merged cells to be written before the header and footer, Hence write header and footer after merged cells have written */
			for(var i=0 ;i < this.config.EXCEL_WORKSHEETNAME.length; i++ ) {
				this.worksheet[i] = this.workbook.createWorksheet({name: this.config.EXCEL_WORKSHEETNAME[i]});
			}
			
			for(var i=0; i <this.config.DATA.length; i++ ){
				this.type = this.getExcelViewType(this.config.DATA[i]);
				this.maxRowCol = this.getTblMaxColnRow(this.config.DATA[i]);
				this.config.EXCEL_MAXCOL = this.maxRowCol[0];
				this.config.EXCEL_MAXROW = this.maxRowCol[1];
				this.excelPosition = this.getExcelCellPostion(this.config.DATA[i]);
				this.config.EXCEL_POS = this.excelPosition.pos;
				this.config.EXCEL_REFTABLE = this.excelPosition.table;
				for(var j=0; j< this.type.length; j++) {
					this.addDatatoExcel(this.type[j],i,j);
				}
			}
			/* Write the merged cells to sheet 1 */
			for(var j = 0; j< this.excelPosition.mergedCells.length; j++) {
				this.worksheet[0].mergeCells([this.excelPosition.mergedCells[j][0]],[this.excelPosition.mergedCells[j][1]]);
			}
			/* write the header and footer */
			for(var i=0 ;i < this.config.EXCEL_WORKSHEETNAME.length; i++ ) {
				this.worksheet[i].setHeader(this.config.EXCEL_WORKSHEETHEADER[i] ? this.config.EXCEL_WORKSHEETHEADER[i] : this.config.EXCEL_WORKSHEETHEADER[0]);
				this.worksheet[i].setFooter(this.config.EXCEL_WORKSHEETFOOTER[i] ? this.config.EXCEL_WORKSHEETFOOTER[i] : this.config.EXCEL_WORKSHEETFOOTER[0]);
				this.worksheet[i].setPageOrientation(this.config.EXCEL_ORIENTATION);
				this.worksheet[i].setStyle(this.styles);
				this.worksheet[i].setTblCol(this.config.EXCEL_TABLECOLS);
				this.worksheet[i].addDrawings(this.drawings);
			}
			
			this.workbook.addDrawings(this.drawings);
			for(var i=0 ;i < this.config.EXCEL_WORKSHEETNAME.length; i++ ) {
				this.workbook.addWorksheet(this.worksheet[i]);
			}
		} catch(e) {
			throw 'Error: EB-parse - '+e;
		}
	},
	generate: function() {
		try {
			var that = this;
			this.path = this.path.replace('.json','.xlsx');
			ExcelBuilder.Builder.createFile(this.workbook, {
				type: 'uint8array',
				path: that.path
			}).then(function (data) {
				data = new Buffer(data,'utf-8');
				Q.nfcall(fs.writeFile,that.path,data,'utf8').then(function(success){ 
					return success; 
				}).fail(function(err){ 
					return err; 
				});
			}).catch(function (e) {
				throw 'Error: EB-generate - '+e;
			});	
		} catch(e) {
			throw 'Error: EB-generate - '+e;
		}
	},
});

module.exports = ExcelReport;
