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
 * spreadsheets API middleware
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

  const getInfo = cb => {
    return step => {
      doc.getInfo((err, info) => {
        if(cb) {
          cb(err, info, step);
        }
        step(err);
      });
    } 
  };

  const executeTest = step => {
    let result = {
      setTitle: false,
      setHeaderRow: false,
      getRows: false,
      resize: false,
      addRow: false,
      setTitle: false,
      del: false
    };
    doc.addWorksheet({title: 'test' + Math.floor(Math.random() * 100000)}, function(err, sheet) {
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

  return function spreadsheetsMiddleware(req, res, next) {
    if (!conforms(_settings)) {
      throwError(next, 400);
    }
    doc = new GoogleSpreadsheet(_settings.sheetId);
    const testComplete = (err, results) => {
      isTesting = false;
      if(err) {
        throwError(next, 500);
      } else {
        res.send(results[results.length-1]);
      }
    };

    if(req.method === 'POST' && ((req.path === '/' && req.body._id) || (/\/.+/.test(req.path) && req.body))) {
      const _id = req.path !== '/' ? /\/(.[^\/]+)(\/.*)*/.exec(req.path)[1] : req.body._id;
      const _data = _.omit(req.body, ['_id']);
      let workSheet;

      if(!doc || !_id || !_data) {
        throwError(next, 400);
      } else {
        series([
          setAuth,
          getInfo((err, info, step) => {
            workSheet = _.find(info.worksheets, ['title', _id]);
          }),
          step => {
            // Add workSheet if needed
            if(!workSheet) {
              doc.addWorksheet({title: _id, rowCount: 2, colCount: 100, headers: Object.keys(_data)}, (err, sheet) => {
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
          step => {
            // Setting Headers
            if(workSheet) {
              let headerRow;
              series([
                next => {
                  workSheet.getCells({'min-row': 1, 'max-row': 1, 'return-empty': false}, (err, cells) => {
                    headerRow = _.cloneDeepWith(cells, targets => {
                      return _.uniq(_.concat(_.map(targets, 'value'), Object.keys(_data)));
                    });
                    next(err);
                  });
                },
                next => {
                  workSheet.setHeaderRow(headerRow, () => {
                    next();
                  });
                }],
              err => {
                step(err);
              });
            } else {
              step();
            }
          },
          step => {
            // Add New Row
            if(workSheet) {
              workSheet.addRow(_data, ()=> {
                step();
              });
            }
          }
        ],
        (err, results) => {
          if(err) {
            throwError(next, 500);
          } else {
            let record = {};
            record[_id] = _data;
            res.send(record);
          }
        });
      }
    } else if(req.path === '/test' && req.method === 'GET') {
      if(isTesting || !doc) {
        throwError(next, 400);
      } else {
        series([setAuth, getInfo(), testWithSheets], testComplete);
      }
    } else {
        throwError(next, 400);
    }
  }

};
