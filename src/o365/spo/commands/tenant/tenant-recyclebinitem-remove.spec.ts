import commands from '../../commands';
import Command, { CommandOption, CommandCancel, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./tenant-recyclebinitem-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';
import config from '../../../../config';

describe(commands.TENANT_RECYCLEBINITEM_REMOVE, () => {
  let vorpal: Vorpal;
  let log: any[];
  let requests: any[];
  let cmdInstance: any;
  
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate.toISOString() }); });

    vorpal = require('../../../../vorpal-init');
    log = [];
    requests = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post,
      global.setTimeout,
      (command as any).ensureFormDigest
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_RECYCLEBINITEM_REMOVE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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

  it('can be cancelled', () => {
    assert(command.cancel());
  });

  it('clears pending connection on cancel', () => {
    (command as any).timeout = {};
    const clearTimeoutSpy = sinon.spy(global, 'clearTimeout');
    (command.cancel() as CommandCancel)();
    Utils.restore(global.clearTimeout);
    assert(clearTimeoutSpy.called);
  });

  it('doesn\'t fail on cancel if no connection pending', () => {
    (command as any).timeout = undefined;
    (command.cancel() as CommandCancel)();
    assert(true);
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
    assert(find.calledWith(commands.TENANT_RECYCLEBINITEM_REMOVE));
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

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com' } });
    assert(actual);
  });

  it('aborts removing deleting site when prompt not confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', debug: true, verbose: true } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the deleted site collection from the tenant recycle bin when prompt confirmed, doesn\'t wait for completion', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","PollingInterval":15000,"IsComplete":true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('removes the deleted site collection from the tenant recycle bin, doesn\'t wait for completion (debug)', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","PollingInterval":15000,"IsComplete":true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('removes the deleted site collection from the tenant recycle bin, doesn\'t wait for completion (verbose)', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","PollingInterval":15000,"IsComplete":true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion. Operation immediately completed', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","PollingInterval":15000,"IsComplete":true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });
  
  it('removes the deleted site collection from the tenant recycle bin, wait for completion', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","IsComplete":false,"PollingInterval":15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;637392077403920220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;67edbe85-2b95-4c7b-a34a-1abbcd68dbe4" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20509.12004","ErrorInfo":null,"TraceCorrelationId":"47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
              },39,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000","IsComplete":true,"PollingInterval":5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion (debug)', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","IsComplete":false,"PollingInterval":15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;637392077403920220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;67edbe85-2b95-4c7b-a34a-1abbcd68dbe4" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20509.12004","ErrorInfo":null,"TraceCorrelationId":"47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
              },39,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000","IsComplete":true,"PollingInterval":5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, wait: true, debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion (verbose)', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":null,"TraceCorrelationId":"5e0d879f-207a-2000-5eb4-5be71488a82a"
              },16,{
              "IsNull":false
              },17,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4","IsComplete":false,"PollingInterval":15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;637392077403920220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;67edbe85-2b95-4c7b-a34a-1abbcd68dbe4" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20509.12004","ErrorInfo":null,"TraceCorrelationId":"47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
              },39,{
              "_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SpoOperation","_ObjectIdentity_":"47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000","IsComplete":true,"PollingInterval":5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, wait: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('did not remove the deleted site collection from the tenant recycle bin', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.20614.12002","ErrorInfo":{
                "ErrorMessage":"Unable to find the deleted site: https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fhr.","ErrorValue":null,"TraceCorrelationId":"b319879f-4090-2000-6ca2-90bc9381b277","ErrorCode":-2147024809,"ErrorTypeName":"System.ArgumentException"
                },"TraceCorrelationId":"b319879f-4090-2000-6ca2-90bc9381b277"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', confirm: true, wait: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("Unable to find the deleted site: https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fhr.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});