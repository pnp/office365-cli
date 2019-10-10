import { Hash } from "../Hash";

export interface ConfigJson {
  $schema?: string;
  bundles?: Object;
  entries?: Entry[];
  localizedResources?: Hash;
  version?: string;
  externals?:ExternalConfiguration;
}

export interface Entry {
  entry: string;
  manifest: string;
  outputPath: string;
}
export interface ExternalConfiguration {
  [key: string]: External;
}
export interface External {
  path: string;
  globalName?: string;
  globalDependencies?: string[];
}