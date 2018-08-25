import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { ListItemInstance } from './ListItemInstance';
import { ContextInfo } from '../../spo';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  title?: string;
  query?: string;
  pageSize?: string;
  filter?: string;
  fields?: string;
}

class SpoListItemListCommand extends SpoCommand {

  public get name(): string {
    return commands.LISTITEM_LIST;
  }

  public get description(): string {
    return 'Get a list of items from the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const listIdArgument = args.options.id || '';
    const listTitleArgument = args.options.title || '';
    
    let siteAccessToken: string = '';
    let formDigestValue: string = '';

    const listRestUrl: string = (args.options.id ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise | Promise<any> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
          cmd.log(``);
          cmd.log(`auth object:`);
          cmd.log(auth);
        }

        if (args.options.query) {
          if (this.debug) {
            cmd.log(`getting request digest for query request`);
          }

          return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
        }
        else {
          return Promise.resolve();
        }
      })
      .then((res: ContextInfo): request.RequestPromise<any> => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }
        
        formDigestValue = args.options.query ? res['FormDigestValue'] : '';
        const rowLimit: string = args.options.pageSize ? `$top=${args.options.pageSize}` : ``
        const filter: string = args.options.filter ? `$filter=${encodeURIComponent(args.options.filter)}` : ``
        const fieldSelect: string = args.options.fields ?
          `?$select=${encodeURIComponent(args.options.fields)}&${rowLimit}&${filter}` :
          (
            (!args.options.output || args.options.output === 'text') ?
              `?$select=Id,Title&${rowLimit}&${filter}` :
              `?${rowLimit}&${filter}`
          )

        const requestBody: any = args.options.query ?
            {
              "query": { 
                "ViewXml": args.options.query 
              } 
            }
          : ``;
        
        const requestOptions: any = {
          url: `${listRestUrl}/${args.options.query ? `GetItems` : `items${fieldSelect}`}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata',
            'X-RequestDigest': formDigestValue
          }),
          json: true,
          body: requestBody
        };
        
        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return args.options.query ? request.post(requestOptions) : request.get(requestOptions);
      })
      .then((response: any): void => {
        (!args.options.output || args.options.output === 'text') && delete response["ID"];
        cmd.log(<ListItemInstance>response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>',
        description: 'URL of the site where the list from which to retrieve items is located'
      },
      {
        option: '-i, --id [listId]',
        description: 'ID of the list from which to retrieve items. Specify id or title but not both'
      },
      {
        option: '-t, --title [listTitle]',
        description: 'Title of the list from which to retrieve items. Specify id or title but not both'
      },
      {
        option: '-s, --pageSize [pageSize]',
        description: 'The number of items to retrieve per page request'
      },
      {
        option: '-q, --query [query]',
        description: 'CAML query to use to retrieve items. Will ignore pageSize if specified'
      },
      {
        option: '-f, --fields [fields]',
        description: 'Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested'
      },
      {
        option: '-l, --filter [odataFilter]',
        description: 'ODATA filter to use to query the list of items with. Specify query or filter but not both'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes {
    return {
      string: [
        'webUrl',
        'id',
        'title',
        'query',
        'pageSize',
        'fields',
        'filter',
      ],
    };
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.id && !args.options.title) {
        return `Specify list id or title`;
      }

      if (args.options.id && args.options.title) {
        return `Specify list id or title but not both`;
      }

      if (args.options.query && args.options.fields) {
        return `Specify query or fields but not both`;
      }

      if (args.options.query && args.options.pageSize) {
        return `Specify query or pageSize but not both`;
      }

      if (args.options.pageSize && isNaN(Number(args.options.pageSize))) {
        return `pageSize must be numeric`;
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} in option id is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To get a list of items from a list, you have to first connect to SharePoint using
    the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Get a list of items from list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x

    Get a list of items from list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} using the CAML query ${chalk.grey('<Query><View><Where><Eq><FieldRef Name=\'Title\' /><Value Type=\'Text\'>Demo list item</Value></Eq></Where></View></Query>')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --query "<Query><View><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo list item</Value></Eq></Where></View></Query>"
    
    Get a list of items from list with a GUID of ${chalk.grey('935c13a0-cc53-4103-8b48-c1d0828eaa7f')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_LIST} --id 935c13a0-cc53-4103-8b48-c1d0828eaa7f --webUrl https://contoso.sharepoint.com/sites/project-x

    Get a list of items from list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}, specifying fields ${chalk.grey('ID,Title,Modified')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified"

    Get a list of items from list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}, with an ODATA filter ${chalk.grey('Title eq \'Demo list item\'')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --filter "Title eq 'Demo list item'"

    Get a list of items from list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}, with a page size of ${chalk.grey('10')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --pageSize 10

   `);
  }

}

module.exports = new SpoListItemListCommand();