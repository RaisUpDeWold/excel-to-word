var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var multer = require('multer');
var Excel = require('exceljs');
var path = require('path');
var fs = require('fs');
var formidable = require('formidable');

app.use(bodyParser.json());
app.use(express.static(path.join(__dirname)));

var excelStorage = multer.diskStorage({ //multers disk storage settings
	destination: function (req, file, cb) {
		cb(null, './uploads/excel/')
	},
	filename: function (req, file, cb) {
		var datetimestamp = Date.now();
		cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
	}
});

var excelUpload = multer({ //multer settings
	storage: excelStorage,
	fileFilter : function(req, file, callback) { //file filter
		if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length-1]) === -1) {
			return callback(new Error('Wrong extension type'));
		}
		callback(null, true);
	}
}).single('file');

/** API path that will upload the files */
app.post('/uploadInputExcel', function(req, res) {
	var exceltojson;
	excelUpload(req,res,function(err){
		if(err){
			res.json({error_code:1,err_desc:err});
			return;
		}
		// Multer gives us file info in req.file object
		if(!req.file){
			res.json({error_code:1,err_desc:"No file passed"});
			return;
		}
		// Check the extension of the incoming file and use the appropriate module
		if(req.file.originalname.split('.')[req.file.originalname.split('.').length-1] === 'xlsx'){
			console.log(req.file.path);
			try {
				var resData = [];

				var workbook = new Excel.Workbook();
        workbook.xlsx.readFile(req.file.path)
          .then(function(data) {


            workbook.eachSheet(function (worksheet, sheetId) {
							worksheet.eachRow({includeEmpty: true}, function(row, rowNumber) {
								row.eachCell(function(cell, colNumber) {
									var cellObj = new Object();
									cellObj.name = cell.name;
									cellObj.value = cell.text;
									resData.push(cellObj);
								})
							});
						});
						res.json({ error_code: 0, err_desc: null, data: resData });
          });
			} catch (e){
				res.json({error_code:1,err_desc:"Corupted excel file"});
			}
		}
	});
});


var wordStorage = multer.diskStorage({ //multers disk storage settings
	destination: function (req, file, cb) {
		cb(null, './uploads/word/')
	},
	filename: function (req, file, cb) {
		var datetimestamp = Date.now();
		cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
	}
});

var wordUpload = multer({ //multer settings
	storage: wordStorage,
	fileFilter : function(req, file, callback) { //file filter
		if (['doc', 'docx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length-1]) === -1) {
			return callback(new Error('Wrong extension type'));
		}
		callback(null, true);
	}
}).single('file');

/** API path that will upload the files */
app.post('/uploadInputWord', function(req, res) {
	wordUpload(req,res,function(err){
		if(err){
			res.json({error_code:1,err_desc:err});
			return;
		}
		/** Multer gives us file info in req.file object */
		if(!req.file){
			res.json({error_code:1,err_desc:"No file passed"});
			return;
		}
		/** Check the extension of the incoming file and
		*  use the appropriate module
		*/
		if(req.file.originalname.split('.')[req.file.originalname.split('.').length-1] === 'docx'){
			console.log(req.file.path);
			res.json({ error_code: 0, err_desc: null, data: req.file.path });
		}
	})
});


app.get('/',function(req,res){
	res.sendFile(__dirname + "/index.html");
});

var port = process.env.PORT || '3000';

app.listen(port, function(){
	console.log('running on ' + port + '...');
});
