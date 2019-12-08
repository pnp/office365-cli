import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoReportOneDriveUsageAccountDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.SPO_REPORT_ONEDRIVEUSAGEACCOUNTDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageAccountDetail';
  }

  public get description(): string {
    return 'Gets details about OneDrive usage by account';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about OneDrive usage by account for the last week
      ${commands.SPO_REPORT_ONEDRIVEUSAGEACCOUNTDETAIL} --period D7

    Gets details about OneDrive usage by account for May 1, 2019
      ${commands.SPO_REPORT_ONEDRIVEUSAGEACCOUNTDETAIL} --date 2019-05-01

    Gets details about OneDrive usage by account for the last week
    and exports the report data in the specified path in text format
      ${commands.SPO_REPORT_ONEDRIVEUSAGEACCOUNTDETAIL} --period D7 --output text --outputFile 'C:/report.txt'

    Gets details about OneDrive usage by account for the last week
    and exports the report data in the specified path in json format
      ${commands.SPO_REPORT_ONEDRIVEUSAGEACCOUNTDETAIL} --period D7 --output json --outputFile 'C:/report.json'
`);
  }
}

module.exports = new SpoReportOneDriveUsageAccountDetailCommand();