import auth from '../../AadAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../AadCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
  scope: string;
}

class Oauth2GrantSetCommand extends AadCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_SET;
  }

  public get description(): string {
    return 'Update OAuth2 permissions for the service principal';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<{ value: any[] }> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Updating OAuth2 permissions...`);
        }

        if (this.verbose) {
          cmd.log(`Updating OAuth2 permissions...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/myorganization/oauth2PermissionGrants/${encodeURIComponent(args.options.grantId)}?api-version=1.6`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'content-type': 'application/json'
          }),
          json: true,
          body: {
            "scope": args.options.scope
          }
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.patch(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res, null, 2));
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --grantId <grantId>',
        description: 'objectId of OAuth2 permission grant to update'
      },
      {
        option: '-s, --scope <scope>',
        description: 'Permissions to grant'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.grantId) {
        return 'Required option grantId missing';
      }

      if (!args.options.scope) {
        return 'Required option scope missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.OAUTH2GRANT_SET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to Azure Active Directory Graph,
      using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To update service principal's OAuth2 permissions, you have to first connect to Azure Active Directory
    Graph using the ${chalk.blue(commands.CONNECT)} command.

    Before you can update service principal's OAuth2 permissions, you need to get the ${chalk.grey('objectId')}
    of the permissions grant to update. You can retrieve it using the ${chalk.blue(commands.OAUTH2GRANT_LIST)} command.
   
  Examples:
  
    Update the existing OAuth2 permission grant with ID ${chalk.grey('YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek')}
    to the ${chalk.grey('Calendars.Read Mail.Read')} permissions
      ${chalk.grey(config.delimiter)} ${commands.OAUTH2GRANT_SET} --grantId YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek --scope "Calendars.Read Mail.Read"

  More information:
  
    Application and service principal objects in Azure Active Directory (Azure AD)
      https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new Oauth2GrantSetCommand();