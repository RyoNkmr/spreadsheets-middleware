var _ = require('lodash');
var GoogleSpreadsheet = require('google-spreadsheet');
var async = require('async');

/**
 * set limitation for loopback application as middleware
 *
 * @param {number} [defaultLimit]
 * @return {Function} middleware
 * @public
 */

module.exports = function(settings) {

  var conforms = _.conforms({
    sheetId: _.isString,
    privateKey: _.isString,
    clientEmail: _.isString
  });
  
  // var settings = settings;
  var doc = new GoogleSpreadsheet(settings.sheetId);

  //methods
  var setAuth = function(step) {
    doc.useServiceAccountAuth({
      client_email: settings.clientEmail,
      private_key: settings.privateKey
    }, step);
  };

  var getInfo = function(step) {
    doc.getInfo(function(err, info) {
      step();
    });
  };

  var testing = false;
  var testWithSheets = function(step) {
    testing = true;
    if(!doc) {
      var err = new Error('Bad Request');
      err.status = 400;
      step(err);
    }
    var result = {
      setTitle: false,
      setHeaderRow: false,
      getRows: false,
      resize: false,
      addRow: false,
      setTitle: false,
      del: false
    };
    doc.addWorksheet({title: 'test' + Math.floor(Math.random() * 100000)}, function(err, sheet) {
      async.series([
        function(next) {
          sheet.setTitle('changing title' + Math.floor(Math.random() * 100000), function(){
            console.log('set title');
            result.setTitle = true;
            next();
          });
        },
        function(next) {
          sheet.setHeaderRow(['hoge', 'fuga', 'piyopiyo', 'foobar'], function(){
            console.log('set headerRow');
            result.setHeaderRow = true;
            next();
          });
        },
        function(next) {
          sheet.getRows({offset: 1, limit: 10}, function(){
            console.log('got rows');
            result.getRows = true;
            next();
          });
        },
        function(next) {
          sheet.resize({rowCount: 20, colCount: 30}, function(){
            console.log('resized sheet');
            result.resize = true;
            next();
          });
        },
        function(next) {
          sheet.addRow({hoge: 'foo', fuga: 'bar', piyopiyo: 'haaa', foobar: 'foo'}, function(){
            console.log('added row');
            result.addRow = true;
            next();
          });
        },
        function(next) {
          sheet.setTitle('test completed' + Math.floor(Math.random() * 100000), function(){
            console.log('re-set title');
            result.setTitle = true;
            next();
          });
        },
        function(next){ 
          sheet.del(function(){
            console.log('deleted sheet');
            result.del = true;
            next();
          });
        }
      ],
      function(err) {
        step(err, result);
      });
    });
  };

  return function spreadsheetsMiddleware(req, res, next) {
    if (!conforms(settings)) {
      var err = new Error('Bad Request');
      err.status = 400;
      next(err);
    }

    var testComplete = function(err, results) {
      if(err) {
        var error = new Error('Internal Server Error');
        error.status = 500;
        next(error);
        return;
      } else {
        res.send(results[results.length-1]);
        testing = false;
      }
    };

    if(req.method === 'POST' && ((req.path === '/' && req.body._id) || (/\/.+/.test(req.path) && req.body))) {
      var _id, _data;
      if(req.path !== '/') {
        _id = /\/(.+)/.exec(req.path)[0];
        _data = req.body;
      } else {
        _id = req.body._id;
        _data = _.omit(req.body, ['_id']);
      }
      var workSheet;

      if(!doc) {
        var error = new Error('Already Running Connection Test');
        error.status = 400;
        next(error);
      } else {
        async.series([
          setAuth,
          function(step){
            //cheking sheet for model
            doc.getInfo(function(err, info){
              if(err){
                step(err);
              } else {
                workSheet = _.find(info.worksheets, ['title', _id]);
                console.log(workSheet);
                step();
              }
            });
          },
          function(step){
            // Add workSheet if needed
            if(!workSheet) {
              doc.addWorksheet({title: _id, rowCount: 1000, colCount: 100, headers: Object.keys(_data)}, function(err, sheet) {
                console.log(err, sheet);
                if(err) {
                  step(err);
                } else {
                  workSheet = sheet;
                  step();
                }
              });
            } else {
              step();
            }
          },
          function(step) {
            // Setting Headers
            if(workSheet) {
              workSheet.setHeaderRow(Object.keys(_data), function(){
                step();
              });
            } else {
              step();
            }
          },
          function(step) {
            // Add New Row
            if(workSheet) {
              workSheet.addRow(_data, function(){
                step();
              });
            }
          }
        ],
        function(err, results) {
          if(err) {
            err.status = 400;
            next(err);
          } else {
            var record = {};
            record[_id] = _data;
            res.send(record);
          }
        });
      }
    } else if(req.path === '/test' && req.method === 'GET') {
      console.log('testing GoogleSpreadsheet API');
      if(testing || !doc) {
        var error = new Error('Already Running Connection Test');
        error.status = 400;
        next(error);
      } else {
        async.series([setAuth, getInfo, testWithSheets], testComplete);
      }
    } else {
      var error = new Error('Already Running Connection Test');
      error.status = 400;
      next(error);
    }
  }

};
