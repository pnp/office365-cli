import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./theme-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.THEME_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.getAccessToken,
      auth.restoreAuth,
      request.get,
      request.post
    ]);  
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.THEME_LIST), true);
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
        assert.equal(telemetry.name, commands.THEME_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: false,
      }
    }, () => {

      try {
        assert(true)
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.get
        ]);
      }
    });
  });

  it('uses correct API url (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        return Promise.resolve('Correct Url')
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.get
        ]);
      }
    });
  });
  
  it('retrieves available themes from the tenant store - output json', (done) => {
    let expected:any = {
      "odata.metadata": "https://m365x642699.sharepoint.com/sites/HBI-testsite/_api/$metadata#SP.Utilities.ThemingOptions",
      "hideDefaultThemes": false,
      "themePreviews": [
          {
              "name": "Contoso01",
              "themeJson": "{\"palette\":{\"themePrimary\":\"#284b68\",\"themeLighterAlt\":\"#ecf3f8\",\"themeLighter\":\"#cfe0ed\",\"themeLight\":\"#8bb3d3\",\"themeTertiary\":\"#417bab\",\"themeSecondary\":\"#2c5474\",\"themeDarkAlt\":\"#24445e\",\"themeDark\":\"#193043\",\"themeDarker\":\"#16293a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}"
          },
          {
              "name": "Contoso02",
              "themeJson": "{\"palette\":{\"themePrimary\":\"#284b68\",\"themeLighterAlt\":\"#ecf3f8\",\"themeLighter\":\"#cfe0ed\",\"themeLight\":\"#8bb3d3\",\"themeTertiary\":\"#417bab\",\"themeSecondary\":\"#2c5474\",\"themeDarkAlt\":\"#24445e\",\"themeDark\":\"#193043\",\"themeDarker\":\"#16293a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}"
          }
      ]
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(expected);
        }
      }
      return Promise.reject('Invalid request');
      });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({options:{ debug: true, verbose: true, output: "json"}}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(expected),'Invalid request');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('retrieves available themes from the tenant store - output text', (done) => {
    let expected:any = {
      "odata.metadata": "https://m365x642699.sharepoint.com/sites/HBI-testsite/_api/$metadata#SP.Utilities.ThemingOptions",
      "hideDefaultThemes": false,
      "themePreviews": [
          {
              "name": "Contoso01",
              "themeJson": "{\"palette\":{\"themePrimary\":\"#284b68\",\"themeLighterAlt\":\"#ecf3f8\",\"themeLighter\":\"#cfe0ed\",\"themeLight\":\"#8bb3d3\",\"themeTertiary\":\"#417bab\",\"themeSecondary\":\"#2c5474\",\"themeDarkAlt\":\"#24445e\",\"themeDark\":\"#193043\",\"themeDarker\":\"#16293a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}"
          },
          {
              "name": "Contoso02",
              "themeJson": "{\"palette\":{\"themePrimary\":\"#284b68\",\"themeLighterAlt\":\"#ecf3f8\",\"themeLighter\":\"#cfe0ed\",\"themeLight\":\"#8bb3d3\",\"themeTertiary\":\"#417bab\",\"themeSecondary\":\"#2c5474\",\"themeDarkAlt\":\"#24445e\",\"themeDark\":\"#193043\",\"themeDarker\":\"#16293a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}"
          }
      ]
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(expected);
        }
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({options:{ debug: true, verbose: true}}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(expected),'Invalid request');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('retrieves available themes - no custom themes available', (done) => {
    let expected:any = {
      "odata.metadata": "https://m365x642699.sharepoint.com/sites/HBI-testsite/_api/$metadata#SP.Utilities.ThemingOptions",
      "hideDefaultThemes": false,
      "themePreviews": []
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(expected);
        }
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({options:{ debug: true, verbose: true}}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(expected),'Invalid request');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('retrieves available themes - handle error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({options:{ debug: true, verbose: true}}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.THEME_LIST));
  });
});