'use strict';

const GoogleSpreadsheet = require('google-spreadsheet');
const spreadsheetsMiddleware = require('../index.js');
const { assert } = require('chai');
const sinon = require('sinon');

suite('spreadsheets-middleware', ()=> {
  const sandbox = sinon.sandbox.create();
  let req;
  let res;

  setup(()=> {
    req = sandbox.stub();
    req.headers = sandbox.stub();
    res = sandbox.stub();
    res.status = sandbox.spy();
    res.send = sandbox.spy();
    res.end = sandbox.spy();
  })

  teardown(()=> {
    sandbox.restore();
  })

  suite('setup', ()=> {
    suite('with invalid params', ()=> {
      const middleware = spreadsheetsMiddleware({sheetId: 'me', clientEmail: 'hoge'});

      test('should call function `next` with error', () => {
        middleware(req, res, err => {
          assert.instanceOf(err, Error);
          assert.property(err, 'message', 'Bad Request');
        });
      });
    });
  });

});
