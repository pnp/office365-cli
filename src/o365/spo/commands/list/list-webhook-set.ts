import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  notificationUrl?: string;
  expirationDateTime?: string;
  id: string;
}

class SpoListWebhookSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_SET;
  }

  public get description(): string {
    return 'Updates the specified webhook';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.notificationUrl = (!(!args.options.notificationUrl)).toString();
    telemetryProps.expirationDateTime = (!(!args.options.expirationDateTime)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving information about the specified webhook...`);
        }

        if (this.verbose) {
          const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
          cmd.log(`Updating webhook ${args.options.id} belonging to list ${list} located at site ${args.options.webUrl}...`);
        }

        let requestUrl: string = '';

        if (args.options.listId) {
          requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/Subscriptions('${encodeURIComponent(args.options.id)}')`;
        }
        else {
          requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/Subscriptions('${encodeURIComponent(args.options.id)}')`;
        }

        const requestBody: any = {};
        if (args.options.notificationUrl) {
          requestBody.notificationUrl = args.options.notificationUrl;
        }
        if (args.options.expirationDateTime) {
          requestBody.expirationDateTime = args.options.expirationDateTime;
        }

        const requestOptions: any = {
          url: requestUrl,
          method: 'PATCH',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: requestBody,
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.patch(requestOptions);
      })
      .then((res: any): void => {
         // REST patch call doesn't return anything
         cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to retrieve webhooks for is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list from which to retrieve the webhook. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list from which to retrieve the webhook. Specify either listId or listTitle but not both'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the webhook to retrieve'
      },
      {
        option: '-n, --notificationUrl [notificationUrl]',
        description: 'The new notification url'
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]',
        description: 'The new expiration date'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      if (!args.options.notificationUrl && !args.options.expirationDateTime) {
        return 'Specify notificationUrl, expirationDateTime or both, at least one is required';
      }

      const parsedDateTime = Date.parse(args.options.expirationDateTime as string)
      const chalk = vorpal.chalk;
      if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
        return `Provide the date in one of the following formats:
      ${chalk.grey('YYYY-MM-DD')}
      ${chalk.grey('YYYY-MM-DDThh:mm')}
      ${chalk.grey('YYYY-MM-DDThh:mmZ')}
      ${chalk.grey('YYYY-MM-DDThh:mm±hh:mm')}`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to SharePoint,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To update a webhook, you have to first log in to SharePoint
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('id')} doesn't refer to an existing webhook,
    you will get a ${chalk.grey('404 - "404 FILE NOT FOUND"')} error.
        
  Examples:
  
    Update the notification url of a webhook with ID ${chalk.grey('cc27a922-8224-4296-90a5-ebbc54da2e81')} which
    belongs to a list with ID ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/ninja')} to ${chalk.grey('https://contoso-funcions.azurewebsites.net/webhook')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_SET} --webUrl https://contoso.sharepoint.com/sites/ninja --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id cc27a922-8224-4296-90a5-ebbc54da2e81 --notificationUrl https://contoso-funcions.azurewebsites.net/webhook

    Update the expiration date of a webhook with ID ${chalk.grey('cc27a922-8224-4296-90a5-ebbc54da2e81')} which
    belongs to a list with title ${chalk.grey('Documents')} located in site 
    ${chalk.grey('https://contoso.sharepoint.com/sites/ninja')} to to October 9 2018
      ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_SET} --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e81 --expirationDateTime 2018-10-09
      `);
  }
}

module.exports = new SpoListWebhookSetCommand();