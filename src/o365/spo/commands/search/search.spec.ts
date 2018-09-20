import commands from '../../commands';
import Command, {CommandValidate,CommandOption,CommandError} from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./search');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SEARCH, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let stubAuth: any = () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SEARCH), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.SEARCH);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/search') > -1) {
        return Promise.resolve(
          {
            "ElapsedTime": 83,
            "PrimaryQueryResult": {
              "CustomResults": [],
              "QueryId": "00000000-0000-0000-0000-000000000000",
              "QueryRuleId": "00000000-0000-0000-0000-000000000000",
              "RefinementResults": null,
              "RelevantResults": {
                "GroupTemplateId": null,
                "ItemTemplateId": null,
                "Properties": [
                  {
                    "Key": "GenerationId",
                    "Value": "9223372036854775806",
                    "ValueType": "Edm.Int64"
                  }
                ],
                "ResultTitle": null,
                "ResultTitleUrl": null,
                "RowCount": 0,
                "Table": {
                  "Rows": []
                },
                "TotalRows": 0,
                "TotalRowsIncludingDuplicates": 0
              },
              "SpecialTermResults": null
            },
            "Properties": [
              {
                "Key": "RowLimit",
                "Value": "10",
                "ValueType": "Edm.Int32"
              }
            ],
            "SecondaryQueryResults": [],
            "SpellingSuggestion": "",
            "TriggeredRules": []
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        query: '*'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ElapsedTime: 83,
          PrimaryQueryResult: {
            CustomResults: [],
            QueryId: "00000000-0000-0000-0000-000000000000",
            QueryRuleId: "00000000-0000-0000-0000-000000000000",
            RefinementResults: null,
            RelevantResults: {
              GroupTemplateId: null,
              ItemTemplateId: null,
              Properties: [
                {
                  Key: "GenerationId",
                  Value: "9223372036854775806",
                  ValueType: "Edm.Int64"
                }
              ],
              ResultTitle: null,
              ResultTitleUrl: null,
              RowCount: 0,
              Table: {
                Rows: []
              },
              TotalRows: 0,
              TotalRowsIncludingDuplicates: 0
            },
            SpecialTermResults: null
          },
          Properties: [
            {
              Key: "RowLimit",
              Value: "10",
              ValueType: "Edm.Int32"
            }
          ],
          SecondaryQueryResults: [],
          SpellingSuggestion: "",
          TriggeredRules: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with output option text', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/search') > -1) {
        return Promise.resolve(
          {
            "ElapsedTime": 83,
            "PrimaryQueryResult": {
              "CustomResults": [],
              "QueryId": "00000000-0000-0000-0000-000000000000",
              "QueryRuleId": "00000000-0000-0000-0000-000000000000",
              "RefinementResults": null,
              "RelevantResults": {
                "GroupTemplateId": null,
                "ItemTemplateId": null,
                "Properties": [
                  {
                    "Key": "GenerationId",
                    "Value": "9223372036854775806",
                    "ValueType": "Edm.Int64"
                  }
                ],
                "ResultTitle": null,
                "ResultTitleUrl": null,
                "RowCount": 0,
                "Table": {
                  "Rows": []
                },
                "TotalRows": 0,
                "TotalRowsIncludingDuplicates": 0
              },
              "SpecialTermResults": null
            },
            "Properties": [
              {
                "Key": "RowLimit",
                "Value": "10",
                "ValueType": "Edm.Int32"
              }
            ],
            "SecondaryQueryResults": [],
            "SpellingSuggestion": "",
            "TriggeredRules": []
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ElapsedTime: 83,
          PrimaryQueryResult: {
            CustomResults: [],
            QueryId: "00000000-0000-0000-0000-000000000000",
            QueryRuleId: "00000000-0000-0000-0000-000000000000",
            RefinementResults: null,
            RelevantResults: {
              GroupTemplateId: null,
              ItemTemplateId: null,
              Properties: [
                {
                  Key: "GenerationId",
                  Value: "9223372036854775806",
                  ValueType: "Edm.Int64"
                }
              ],
              ResultTitle: null,
              ResultTitleUrl: null,
              RowCount: 0,
              Table: {
                Rows: []
              },
              TotalRows: 0,
              TotalRowsIncludingDuplicates: 0
            },
            SpecialTermResults: null
          },
          Properties: [
            {
              Key: "RowLimit",
              Value: "10",
              ValueType: "Edm.Int32"
            }
          ],
          SecondaryQueryResults: [],
          SpellingSuggestion: "",
          TriggeredRules: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('command correctly handles reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/webs') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying query', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<query>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the query option is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation if all options are provided', () => {
    const actual = (command.validate() as CommandValidate)({ options: { query:'*' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.SEARCH));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 