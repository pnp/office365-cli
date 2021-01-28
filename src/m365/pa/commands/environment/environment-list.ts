import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class PowerAppEnvironmentListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power App environments in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of Microsoft Power App environments...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.PowerApps/environments?api-version=2017-08-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ name: string, displayName: string; properties: { displayName: string } }] }>(requestOptions)
      .then((res: { value: [{ name: string, displayName: string; properties: { displayName: string } }] }): void => {
        if (res.value && res.value.length > 0) {
          res.value.forEach(e => {
            e.displayName = e.properties.displayName
          });

          logger.log(res.value);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new PowerAppEnvironmentListCommand();
