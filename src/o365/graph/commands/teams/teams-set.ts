import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  displayName?: string;
  description?: string;
  mailNickName?: string;
  classification?: string;
  visibility?: string;
}

class GraphTeamsSetCommand extends GraphCommand {
  private static props: string[] = [
    'displayName',
    'description',
    'mailNickName',
    'classification',
    'visibility '
  ];

  public get name(): string {
    return `${commands.TEAMS_SET}`;
  }

  public get description(): string {
    return 'Updates settings of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    GraphTeamsSetCommand.props.forEach((p: string) => {
        telemetryProps[p] = typeof (args.options as any)[p] !== 'undefined';
    });
    return telemetryProps;
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    if (options.displayName) {
      requestBody.displayName = options.displayName;
    }
    if (options.description) {
      requestBody.description = options.description;
    }
    if (options.mailNickName) {
      requestBody.mailNickName = options.mailNickName;
    }
    if (options.classification) {
      requestBody.classification = options.classification;
    }
    if (options.visibility) {
      requestBody.visibility = options.visibility;
    }
    return requestBody;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {
        const body: any = this.mapRequestBody(args.options);

        const requestOptions: any = {
          url: `${auth.service.resource}/beta/groups/${encodeURIComponent(args.options.teamId)}`,
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
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team for which to update settings'
      },
      {
        option: '--displayName [displayName]',
        description: 'The display name for the Microsoft Teams team'
      },
      {
        option: '--description [description]',
        description: 'The description for the Microsoft Teams team'
      },
      {
        option: '--mailNickName [mailNickName]',
        description: 'The mail alias for the Microsoft Teams team'
      },
      {
        option: '--classification [classification]',
        description: 'The classification for the Microsoft Teams team with valid values HBI, MBI, LBI, GDPR'
      },
      {
        option: '--visibility [visibility]',
        description: 'Specifies the visibility of the Microsoft Teams team with valid values Private,Public'
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

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.

    To update the settings of the specified Microsoft Teams team, you have to
    first log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    Set Microsoft Teams team visibility as private
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --visibility private

    Set Microsoft Teams team clasiification as MBI
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --classification MBI
`);
  }
}

module.exports = new GraphTeamsSetCommand();