const _ = require('lodash');
const GoogleSpreadsheet = require('google-spreadsheet');
const series = require('async/series');

const conforms = _.conforms({
  sheetId: _.isString,
  privateKey: _.isString,
  clientEmail: _.isString
});

const throwError = (next, code) => {
  let error;
  if(code === 400) {
    error = new Error('Bad Request');
    error.status = code;
  } else {
    error = new Error('Internal Server Error');
    error.status = 500;
  }
  next(error);
};

/**
 * spreadsheets Connection Test middleware
 *
 * @param {Object} [sheetId{string}, privateKey{string}, clientEmail{string}]
 * @return {Function} middleware
 * @public
 */

module.exports = settings => {

  let doc;
  let isTesting = false;

  // properly parse escaped multi line string
  // const _settings = _.cloneDeepWith(settings, opts => _.mapValues(opts, str => str.replace(/\\+n/g, '\n')));
  const _settings = settings;

  const setAuth = step => {
    doc.useServiceAccountAuth({
      client_email: _settings.clientEmail,
      private_key: _settings.privateKey
    }, step);
  };

  const getInfo = step => {
    doc.getInfo((err, info) => {
      step(err);
    });
  };

  const executeTest = step => {
    let result = {
      addWorksheet: false,
      setTitle: false,
      setHeaderRow: false,
      getRows: false,
      resize: false,
      addRow: false,
      setTitle: false,
      del: false
    };
    doc.addWorksheet({title: 'test' + Math.floor(Math.random() * 100000)}, function(err, sheet) {
      result.addWorksheet = true;
      series([
        next => {
          sheet.setTitle('changing title' + Math.floor(Math.random() * 100000), function(){
            result.setTitle = true;
            next();
          });
        },
        next => {
          sheet.setHeaderRow(['hoge', 'fuga', 'piyopiyo', 'foobar'], function(){
            result.setHeaderRow = true;
            next();
          });
        },
        next => {
          sheet.getRows({offset: 1, limit: 10}, function(){
            result.getRows = true;
            next();
          });
        },
        next => {
          sheet.resize({rowCount: 20, colCount: 30}, function(){
            result.resize = true;
            next();
          });
        },
        next => {
          sheet.addRow({hoge: 'foo', fuga: 'bar', piyopiyo: 'haaa', foobar: 'foo'}, () => {
            result.addRow = true;
            next();
          });
        },
        next => {
          sheet.setTitle('test completed' + Math.floor(Math.random() * 100000), () => {
            result.setTitle = true;
            next();
          });
        },
        next => {
          sheet.del(() => {
            result.del = true;
            next();
          });
        }
      ],
      err => {
        step(err, result);
      });
    });
  }

  const testWithSheets = step => {
    isTesting = true;
    if(!doc) {
      throwError(step, 400)
    } else {
      executeTest(step);
    }
  };

  return function spreadsheetsConnectionTestMiddleware(req, res, next) {
    if (!conforms(_settings)) {
      throwError(next, 400);
    }
    doc = new GoogleSpreadsheet(_settings.sheetId);
    if(req.path === '/test' && req.method === 'POST') {
      if(isTesting || !doc) {
        throwError(next, 400);
      } else {
        series([setAuth, getInfo, testWithSheets], (err, results) => {
          isTesting = false;
          if(err) {
            throwError(next, 500);
          } else {
            res.send(results[results.length-1]);
          }
        });
      }
    } else {
        throwError(next, 400);
    }
  }

};
