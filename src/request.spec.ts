import * as sinon from 'sinon';
import * as assert from 'assert';
import _request from './request';
import * as requestPromise from 'request-promise-native';
import Utils from './Utils';
import * as https from 'https';
import request = require('request');
import { WriteStream } from 'fs';

describe('Request', () => {
  const cmdInstance = {
    commandWrapper: {
      command: 'command'
    },
    log: (msg: any) => {},
    prompt: () => {},
    action: () => {}
  };

  let _options: requestPromise.OptionsWithUrl;

  beforeEach(() => {
    _request.cmd = cmdInstance;
  });

  afterEach(() => {
    Utils.restore([
      global.setTimeout,
      https.request,
      (_request as any).req
    ]);
  });

  it('fails when no command instance set', (cb) => {
    _request.cmd = undefined as any;
    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb('Error expected');
      }, (err: any) => {
        try {
          assert.equal(err, 'Command reference not set on the request object');
          cb();
        }
        catch (err) {
          cb(err);
        }
      });
  });

  it('sets user agent on all requests', (cb) => {
    sinon.stub(https, 'request').callsFake((options) => {
      _options = options;
      return new WriteStream({
        final: (cb) => {
          cb()
        }
      });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb('Error expected');
      }, () => {
        try {
          assert((_options.headers as request.Headers)['user-agent'].indexOf('NONISV|SharePointPnP|Office365CLI') > -1);
          cb();
        }
        catch (err) {
          cb(err);
        }
      });
  });

  it('uses gzip compression on all requests', (cb) => {
    sinon.stub(https, 'request').callsFake((options) => {
      _options = options;
      return new WriteStream({
        final: (cb) => {
          cb()
        }
      });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb('Error expected');
      }, () => {
        try {
          assert((_options.headers as request.Headers)['accept-encoding'].indexOf('gzip') > -1);
          cb();
        }
        catch (err) {
          cb(err);
        }
      });
  });

  it('sets method to GET for a GET request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(_options.method, 'GET');
          cb();
        }
        catch (err) {
          cb(err);
        }
      }, (err) => {
        cb(err);
      });
  });

  it('sets method to POST for a POST request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .post({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(_options.method, 'POST');
          cb();
        }
        catch (err) {
          cb(err);
        }
      }, (err) => {
        cb(err);
      });
  });

  it('sets method to PATCH for a PATCH request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .patch({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(_options.method, 'PATCH');
          cb();
        }
        catch (err) {
          cb(err);
        }
      }, (err) => {
        cb(err);
      });
  });

  it('sets method to PUT for a PUT request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .put({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(_options.method, 'PUT');
          cb();
        }
        catch (err) {
          cb(err);
        }
      }, (err) => {
        cb(err);
      });
  });

  it('sets method to DELETE for a DELETE request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .delete({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(_options.method, 'DELETE');
          cb();
        }
        catch (err) {
          cb(err);
        }
      }, (err) => {
        cb(err);
      });
  });

  it('returns response of a successful GET request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb();
      }, (err) => {
        cb(err);
      });
  });

  it('correctly handles failed GET request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.reject('Error');
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb('Error expected');
      }, (err) => {
        try {
          assert.equal(err, 'Error');
          cb();
        }
        catch (e) {
          cb(e);
        }
      });
  });

  it('repeats 429-throttled request after the designated retry value', (cb) => {
    let i: number = 0;
    let timeout: number = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {
              'retry-after': 60
            }
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      timeout = to;
      fn();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(timeout, 60000);
          cb();
        }
        catch (err) {
          cb(err)
        }
      }, (err) => {
        cb(err);
      });
  });

  it('repeats 429-throttled request after 10s if no value specified', (cb) => {
    let i: number = 0;
    let timeout: number = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {}
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      timeout = to;
      fn();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(timeout, 10000);
          cb();
        }
        catch (err) {
          cb(err)
        }
      }, (err) => {
        cb(err);
      });
  });

  it('repeats 429-throttled request after 10s if the specified value is not a number', (cb) => {
    let i: number = 0;
    let timeout: number = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {
              'retry-after': 'a'
            }
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      timeout = to;
      fn();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(timeout, 10000);
          cb();
        }
        catch (err) {
          cb(err)
        }
      }, (err) => {
        cb(err);
      });
  });

  it('repeats 429-throttled request until it succeeds', (cb) => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ < 3) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {}
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(i, 4);
          cb();
        }
        catch (err) {
          cb(err)
        }
      }, (err: any) => {
        cb(err);
      });
  });

  it('repeats 503-throttled request until it succeeds', (cb) => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ < 3) {
        return Promise.reject({
          response: {
            statusCode: 503,
            headers: {}
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.equal(i, 4);
          cb();
        }
        catch (err) {
          cb(err)
        }
      }, (err: any) => {
        cb(err);
      });
  });

  it('correctly handles request that was first 429-throttled and then failed', (cb) => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {}
          }
        })
      }
      else {
        return Promise.reject('Error');
      }
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb('Expected error')
      }, (err) => {
        try {
          assert.equal(err, 'Error');
          cb();
        }
        catch (e) {
          cb(e);
        }
      });
  });
})