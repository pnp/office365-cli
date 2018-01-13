import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  principals: string;
  confirm?: boolean;
}

class SpoSiteDesignRightsRevokeCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_RIGHTS_REVOKE}`;
  }

  public get description(): string {
    return 'Revokes access from a site design for one or more principals';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const revokePermissions: () => void = (): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((accessToken: string): Promise<ContextInfo> => {
          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
          }

          if (this.verbose) {
            cmd.log(`Retrieving request digest...`);
          }

          return this.getRequestDigest(cmd, this.debug);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.debug) {
            cmd.log('Response:')
            cmd.log(res);
            cmd.log('');
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              'X-RequestDigest': res.FormDigestValue,
              'content-type': 'application/json;charset=utf-8',
              accept: 'application/json;odata=nometadata'
            }),
            body: {
              id: args.options.id,
              principalNames: args.options.principals.split(',').map(p => p.trim())
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
        .then((res: any): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      revokePermissions();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to revoke access to site design ${args.options.id} from the specified users?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          revokePermissions();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the site design to revoke rights from'
      },
      {
        option: '-p, --principals <principals>',
        description: 'Comma-separated list of principals to revoke view rights from. Principals can be users or mail-enabled security groups in the form of "alias" or "alias@<domain name>.com"'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the site design'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.principals) {
        return 'Required parameter principals missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site using the
      ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To revoke access to a site design, you have to first connect to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('id')} doesn't refer to an existing site design, you will get
    a ${chalk.grey('File not found')} error.

    If all principals have rights revoked on the site design, the site design becomes viewable to everyone.

    If you try to revoke access for a user that doesn't have access granted to the specified site design
    you will get a ${chalk.grey('The specified user or domain group was not found')} error.

  Examples:
  
    Revoke access to the site design with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} from user
    with alias ${chalk.grey('PattiF')}. Will prompt for confirmation before revoking the access
      ${chalk.grey(config.delimiter)} ${this.name} --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --principals PattiF

    Revoke access to the site design with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} from users
    with aliases ${chalk.grey('PattiF')} and ${chalk.grey('AdeleV')} without prompting for confirmation
      ${chalk.grey(config.delimiter)} ${this.name} --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --principals PattiF,AdeleV --confirm

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignRightsRevokeCommand();