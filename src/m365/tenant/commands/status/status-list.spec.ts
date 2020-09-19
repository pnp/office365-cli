import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./status-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TENANT_STATUS_LIST, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  const textOutputForms = {
    Name: "Microsoft Forms",
    Status: "Normal service"
  };

  const textOutput = [
    {
      Name: "Microsoft Forms",
      Status: "Normal service"
    },
    {
      Name: "Planner",
      Status: "Normal service"
    },
    {
      Name: "Microsoft Stream",
      Status: "Extended recovery"
    },
    {
      Name: "SharePoint Online",
      Status: "Normal service"
    }
  ];

  const jsonOutput = {
    "value": [
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Service",
            "FeatureName": "service",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Form functionality",
            "FeatureName": "functionality",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Integration",
            "FeatureName": "integration",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Forms",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-18T13:15:36.6847769Z",
        "Workload": "Forms",
        "WorkloadDisplayName": "Microsoft Forms"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Planner",
            "FeatureName": "Planner",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Planner",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-18T13:15:36.6847769Z",
        "Workload": "Planner",
        "WorkloadDisplayName": "Planner"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Playback",
            "FeatureName": "Playback",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Live Events",
            "FeatureName": "Live Events",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Stream website",
            "FeatureName": "Stream website",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Stream",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-18T13:15:36.6847769Z",
        "Workload": "Stream",
        "WorkloadDisplayName": "Microsoft Stream"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Provisioning",
            "FeatureName": "provisioning",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "SharePoint Features",
            "FeatureName": "spofeatures",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Tenant Admin",
            "FeatureName": "tenantadmin",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Search and Delve",
            "FeatureName": "search",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Custom Solutions and Workflows",
            "FeatureName": "customsolutionsworkflows",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Project Online",
            "FeatureName": "projectonline",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Office Web Apps",
            "FeatureName": "officewebapps",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "SP Designer",
            "FeatureName": "spdesigner",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Access Services",
            "FeatureName": "accessservices",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "InfoPath Online",
            "FeatureName": "infopathonline",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "SharePoint",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-18T13:15:36.6847769Z",
        "Workload": "SharePoint",
        "WorkloadDisplayName": "SharePoint Online"
      }
    ]
  };
  
  const jsonOutputForms = {
    "value": [
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Service",
            "FeatureName": "service",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Form functionality",
            "FeatureName": "functionality",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Integration",
            "FeatureName": "integration",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Forms",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-18T14:29:02.7203865Z",
        "Workload": "Forms",
        "WorkloadDisplayName": "Microsoft Forms"
      }
    ]
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_STATUS_LIST), true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TENANT_STATUS_LIST));
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

  it('handles promise error while getting status of Microsoft 365 services', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services as text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services as text (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services - JSON Output With Workload', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Forms',
        output: 'json',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutputForms));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services - JSON Output With Workload (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Forms',
        output: 'json',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutputForms));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services - text Output With Workload', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Forms',
        output: 'text',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutputForms));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services - text Output With Workload (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        done();
        return Promise.resolve(jsonOutput);
      }
      done();
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Forms',
        output: 'text',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutputForms));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

});