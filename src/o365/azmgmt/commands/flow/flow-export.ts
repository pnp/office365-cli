import auth from '../../AzmgmtAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import AzmgmtCommand from '../../AzmgmtCommand';
import Utils from '../../../../Utils';
import * as os from 'os';
import * as path from 'path';
import * as fs from 'fs';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  id: string;
  packageDisplayName: string;
  packageDescription: string;
  packageCreatedBy: string;
  packageSourceEnvironment: string;
  format: string;
  path: string;
}

class AzmgmtFlowExportCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_EXPORT;
  }

  public get description(): string {
    return 'Exports the specified Microsoft Flow as a file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.format = args.options.format;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    
    let accessToken: string = '';
    let filenameFromApi = '';
    const formatArgument = args.options.format ? args.options.format.toLowerCase() : '';

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((retrievedAccessToken: string): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${retrievedAccessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving package resources for Microsoft Flow ${args.options.id}...`);
        }

        accessToken = retrievedAccessToken

        if (formatArgument === 'json') {

          if (this.debug) {
            cmd.log('format = json, skipping listing package resources step');
          }

          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${auth.service.resource}providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/listPackageResources?api-version=2016-11-01`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json'
          }),
          body: {
            "baseResourceIds": [
              `/providers/Microsoft.Flow/flows/${args.options.id}`
            ]
          },
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);

      })
      .then((res: any): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Initiating package export for Microsoft Flow ${args.options.id}...`);
        }

        let requestOptions: any = {
          
          url: `${auth.service.resource}providers/${formatArgument === 'json' ? 
            `Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.id)}?api-version=2016-11-01`
            : `Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/exportPackage?api-version=2016-11-01`}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json'
          }),
          json: true
        };

        if (formatArgument !== 'json') {
          requestOptions['body'] = {
            "includedResourceIds":[
              `/providers/Microsoft.Flow/flows/${args.options.id}`
            ],
            "details": {
              "displayName": args.options.packageDisplayName,
              "description": args.options.packageDescription,
              "creator": args.options.packageCreatedBy,
              "sourceEnvironment": args.options.packageSourceEnvironment
            },
            "resources": res.resources
          }
        }

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return formatArgument === 'json' ? request.get(requestOptions) : request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Getting file for Microsoft Flow ${args.options.id}...`);
        }

        if (res.errors && res.errors.length && res.errors.length > 0) {
          return Promise.reject(res.errors[0].message)
        }

        const downloadFileUrl: string = formatArgument === 'json' ? '' : res.packageLink.value;
        const filenameRegEx = /([^\/]+\.zip)/i;
        filenameFromApi = formatArgument === 'json' ? `${res.properties.displayName}.json` : (filenameRegEx.exec(downloadFileUrl) || ['output.zip'])[0];

        if (this.debug) {
          cmd.log(`Filename from PowerApps API: ${filenameFromApi}`);
          cmd.log('');
        }

        const requestOptions: any = {
          url: formatArgument === 'json' ? 
            `${auth.service.resource}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.id)}/exportToARMTemplate?api-version=2016-11-01`
            : downloadFileUrl,
          encoding: null, // Set encoding to null, otherwise binary data will be encoded to utf8 and binary data is corrupt 
          headers: formatArgument === 'json' ? Utils.getRequestHeaders({
              authorization: `Bearer ${accessToken}`,
              accept: 'application/json'
            })
            : Utils.getRequestHeaders({}),
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return formatArgument === 'json' ? 
          request.post(requestOptions)
          : request.get(requestOptions);
      })
      .then((file: string): void => {

        if (this.debug) {
          cmd.log('Response received');
          cmd.log(file);
          cmd.log('');
        }
        const path = args.options.path ? args.options.path : `./${filenameFromApi}`
        
        fs.writeFileSync(path, file, 'binary');
        if (this.verbose || !args.options.path) {
          cmd.log(`File saved to path '${path}'`);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The id of the Microsoft Flow to export'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment from which to export the Flow from'
      },
      {
        option: '-n, --packageDisplayName <name>',
        description: 'The display name to use in the exported package'
      },
      {
        option: '-d, --packageDescription <description>',
        description: 'The description to use in the exported package'
      },
      {
        option: '-c, --packageCreatedBy <name of creator>',
        description: 'The name of the person to be used as the creator of the exported package'
      },
      {
        option: '-s, --packageSourceEnvironment <name of source environment>',
        description: 'The name of the source environment from which the exported package was taken'
      },
      {
        option: '-f, --format <format type>',
        description: 'The format to export the Flow to json|zip. Default json'
      },
      {
        option: '-p, --path <path>',
        description: 'The path to save the exported package to'
      },

    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      const lowerCaseFormat = args.options.format ? args.options.format.toLowerCase() : '';

      if (!args.options.id) {
        return 'Required option id missing';
      }

      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.format && (lowerCaseFormat !== 'json' && lowerCaseFormat !== 'zip')) {
        return 'Option format must be json or zip. Default is zip';
      }

      if (lowerCaseFormat === 'json' && args.options.packageCreatedBy) {
        return 'packageCreatedBy cannot be specified with output of json';
      }

      if (lowerCaseFormat === 'json' && args.options.packageDescription) {
        return 'packageDescription cannot be specified with output of json';
      }

      if (lowerCaseFormat === 'json' && args.options.packageDisplayName) {
        return 'packageDisplayName cannot be specified with output of json';
      }

      if (lowerCaseFormat === 'json' && args.options.packageSourceEnvironment) {
        return 'packageSourceEnvironment cannot be specified with output of json';
      }

      if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
        return 'Specified path where to save the file does not exits';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_EXPORT).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Azure Management Service,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    To export the specified Microsoft Flow, you have to first connect to the Azure 
    Management Service using the ${chalk.blue(commands.CONNECT)} command.

    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    If the Microsoft Flow with the id you specified doesn't exist, you will
    get the ${chalk.grey(`The caller with object id \'abc\' does not have permission${os.EOL}` +
    '    for connection \'xyz\' under Api \'shared_logicflows\'.')} error.
   
  Examples:
  
    Export the specified Microsoft Flow as a ZIP file
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d

    Export the specified Microsoft Flow as a JSON file
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --format json

    Export the specified Microsoft Flow as a ZIP file with a Package Display Name of 'My flow name'
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --packageDisplayName 'My flow name'

    Export the specified Microsoft Flow as a ZIP file with the filename 'MyFlow.zip' saved to the current directory
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --path './MyFlow.zip'

`);
  }
}

module.exports = new AzmgmtFlowExportCommand();