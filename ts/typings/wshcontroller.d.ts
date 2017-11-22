//
//  Types for WshController
//
// WshController    https://msdn.microsoft.com/ja-jp/library/cc364348.aspx

interface ActiveXObject {
  // Add specialized signatures
  new (s: 'WSHController'): WshController;
}

interface WshController {
  CreateScript(CommandLine: string, MachineName?: string): WshRemote;
}

interface WshRemote {
  readonly Status: WshRemote.RemoteState;
  readonly Error: WshRemoteError;
  Execute(): void;
  Terminate(): void;
}

interface WshRemoteError {
  readonly Description: string;
  readonly Line: number;
  readonly Character: number;
  readonly SourceText: string;
  readonly Source: string;
  readonly Number: number;
}

export namespace WshRemote {
  export const enum RemoteState {
    NoTask = 0,
    Running = 1,
    Finished = 2,
  }
}
