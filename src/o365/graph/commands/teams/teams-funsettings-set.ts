import auth from '../../GraphAuth';
import Utils from '../../../../Utils';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../GraphCommand';
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  allowGiphy: string;
  giphyContentRating: string;
  allowStickersAndMemes: string;
  allowCustomMemes: string;
}

class GraphTeamsFunSettingsSetCommand extends GraphCommand {

  private static booleanProps: string[] = [
    'allowGiphy',
    'allowStickersAndMemes',
    'allowCustomMemes'
  ];

  public get name(): string {
    return `${commands.TEAMS_FUNSETTINGS_SET}`;
  }

  public get description(): string {
    return 'Updates fun settings of a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {
        const body: any = {
          funSettings: {}
        };
        GraphTeamsFunSettingsSetCommand.booleanProps.forEach(p => {
          if (typeof (args.options as any)[p] !== 'undefined') {
            body.funSettings[p] = (args.options as any)[p] === 'true';
          }
        });

        if (args.options.giphyContentRating) {
          body.funSettings.giphyContentRating = args.options.giphyContentRating;
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          },
          body: body,
          json: true
        };

        return request.patch(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Teams team for which to update settings'
      },
      {
        option: '--allowGiphy [allowGiphy]',
        description: 'Set to true to allow giphy and to false to disable it'
      },
      {
        option: '--giphyContentRating [giphyContentRating]',
        description: 'Settings to set content rating for giphy. Allowed values Strict|Moderate'
      },
      {
        option: '--allowStickersAndMemes [allowStickersAndMemes]',
        description: 'Set to true to allow stickers and memes and to false to disable them'
      },
      {
        option: '--allowCustomMemes [allowCustomMemes]',
        description: 'Set to true to allow custom memes and to false to disable them'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      let isValid: boolean = true;
      let value, property: string = '';
      GraphTeamsFunSettingsSetCommand.booleanProps.every(p => {
        property = p;
        value = (args.options as any)[p];
        isValid = typeof value === 'undefined' ||
          value === 'true' ||
          value === 'false';
        return isValid;
      });

      if (!isValid) {
        return `Value ${value} for option ${property} is not a valid boolean`;
      }

      if (args.options.giphyContentRating) {
        const giphyContentRating = args.options.giphyContentRating.toLowerCase();
        if (giphyContentRating !== 'strict' && giphyContentRating !== 'moderate') {
          return `giphyContentRating value ${value} is not valid.  Please specify Strict or Moderate.`
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To set fun settings of a Microsoft Teams team, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Allow giphy usage within a given Microsoft Teams team, setting the content rating for giphy to Moderate
      ${chalk.grey(config.delimiter)} ${this.name} --teamId 83cece1e-938d-44a1-8b86-918cf6151957 --allowGiphy true --giphyContentRating Moderate
    
    Disable usage of giphy within the given Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId 83cece1e-938d-44a1-8b86-918cf6151957 --allowGiphy false

    Allow usage of Stickers and Memes within a given Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId 83cece1e-938d-44a1-8b86-918cf6151957 --allowStickersAndMemes true

    Disable usage Custom Memes within a given Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId 83cece1e-938d-44a1-8b86-918cf6151957 --allowCustomMemes false

`);
  }
}

module.exports = new GraphTeamsFunSettingsSetCommand();