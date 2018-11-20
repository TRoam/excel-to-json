var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var multer = require('multer');
var readXlsxFile = require('read-excel-file/node');
const fs = require('fs');

app.use(bodyParser.json());

var storage = multer.diskStorage({
  //multers disk storage settings
  destination: function(req, file, cb) {
    cb(null, './uploads/');
  },
  filename: function(req, file, cb) {
    var datetimestamp = Date.now();
    cb(
      null,
      file.fieldname +
        '-' +
        datetimestamp +
        '.' +
        file.originalname.split('.')[file.originalname.split('.').length - 1]
    );
  }
});

var upload = multer({
  //multer settings
  storage: storage,
  fileFilter: function(req, file, callback) {
    //file filter
    if (
      ['xls', 'xlsx'].indexOf(
        file.originalname.split('.')[file.originalname.split('.').length - 1]
      ) === -1
    ) {
      return callback(new Error('Wrong extension type'));
    }
    callback(null, true);
  }
}).single('file');

var createFile = function(data, index) {
    var keys = data[0];
    var results = {};
    for(var i = 1; i < data.length; i++){
        for( var j = 0; j < data[i].length; j ++) {
            if(!results[keys[i]]) {
                results[keys[i]] = []; 
            }
            results[keys[i]].push(data[i][j]); 
        }
    }
    fs.writeFile(index, JSON.stringify(data), err => {
        if (err) throw err;
        console.log('complete');
    });

}

/** API path that will upload the files */
app.post('/upload', function(req, res) {
  upload(req, res, function(err) {
    if (err) {
      res.json({error_code: 1, err_desc: err});
      return;
    }
    /** Multer gives us file info in req.file object */
    if (!req.file) {
      res.json({error_code: 1, err_desc: 'No file passed'});
      return;
    }
    console.log(req.file.path);
    readXlsxFile(req.file.path, { sheet: 1 }).then((data) => {
        createFile(data, '1.json'); 
    })
    readXlsxFile(req.file.path, { sheet: 2 }).then((data) => {
        createFile(data, '.2.json'); 
    })
    readXlsxFile(req.file.path, { sheet: 3 }).then((data) => {
        createFile(data, '3.json'); 
    })
    readXlsxFile(req.file.path, { sheet: 4 }).then((data) => {
        createFile(data, '4.json'); 
    })
    readXlsxFile(req.file.path, { sheet: 5 }).then((data) => {
        createFile(data, '5.json'); 
    })
    readXlsxFile(req.file.path, { sheet: 6 }).then((data) => {
        createFile(data, '6.json'); 
    })
    res.json({result: 'ok'});
  });
});

app.get('/', function(req, res) {
  res.sendFile(__dirname + '/index.html');
});

app.listen('3000', function() {
  console.log('running on 3000...');
});
