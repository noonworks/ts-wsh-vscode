//
//  Types for WshShell
//
// WshShell         https://msdn.microsoft.com/ja-jp/library/cc364436.aspx

interface ActiveXObject {
  // Add specialized signatures
  new (s: 'WScript.Shell'): WshShell;
}

interface WshShell {
  // properties
  CurrentDirectory: string;
  Environment: (strType?: WshShell.EnvironmentType) => WshEnvironment;
  SpecialFolders: (objWshSpecialFolders: WshShell.SpecialFolderType) => string;
  // methods
  AppActivate(title: string|number): void;
  CreateShortcut(strPathname: string): WshShortcut | WshURLShortcut;
  ExpandEnvironmentStrings(strString: string): string;
  LogEvent(intType: WshShell.EventType, strMessage: string, strTarget?: string): boolean;
  Popup(strText: string, nSecondsToWait?: number, strTitle?: string,
    nType?: WshShell.ButtonType | WshShell.IconType | WshShell.DefaultButton): WshShell.PopupResult;
  RegDelete(strName: string): void;
  RegRead(strName: string): string | number | string[] | number[];
  RegWrite(strName: string, anyValue: any, strType?: WshShell.RegType): void;
  Run(strCommand: string, intWindowStyle?: WshShell.RunWindowStyle, bWaitOnReturn?: boolean): number;
  SendKeys(key: string): void;
  Exec(strCommand: string): WshScriptExec;
}

interface WshEnvironment {
  readonly length: number;
  Item(key: string): string;
  Remove(strName: string): void;
}

interface WshShortcut {
  Arguments: string;
  Description: string;
  readonly FullName: string;
  Hotkey: string;
  IconLocation: string;
  TargetPath: string;
  WindowStyle: WshShell.ShortcutWindowStyle;
  WorkingDirectory: string;
  Save(): void;
}

interface WshURLShortcut {
  readonly FullName: string;
  TargetPath: string;
  Save(): void;
}

interface WshScriptExec {
  readonly Status: WshShell.ExecState;
  readonly StdErr: TextStreamReader;
  readonly StdIn: TextStreamWriter;
  readonly StdOut: TextStreamReader;
  Terminate(): void;
}

declare namespace WshShell {
  export const enum EnvironmentType {
    SYSTEM = 'SYSTEM',
    USER = 'USER',
    VOLATILE = 'VOLATILE',
    PROCESS = 'PROCESS',
  }

  export const enum SpecialFolderType {
    AllUsersDesktop = 'AllUsersDesktop',
    AllUsersStartMenu = 'AllUsersStartMenu',
    AllUsersPrograms = 'AllUsersPrograms',
    AllUsersStartup = 'AllUsersStartup',
    Desktop = 'Desktop',
    Favorites = 'Favorites',
    Fonts = 'Fonts',
    MyDocuments = 'MyDocuments',
    NetHood = 'NetHood',
    PrintHood = 'PrintHood',
    Programs = 'Programs',
    Recent = 'Recent',
    SendTo = 'SendTo',
    StartMenu = 'StartMenu',
    Startup = 'Startup',
    Templates = 'Templates',
  }

  export const enum EventType {
    SUCCESS = 0,
    ERROR = 1,
    WARNING = 2,
    INFORMATION = 4,
    AUDIT_SUCCESS = 8,
    AUDIT_FAILURE = 16,
  }

  export const enum ButtonType {
    OKOnly = 0,
    OKCancel = 1,
    AbortRetryIgnore = 2,
    YesNoCancel = 3,
    YesNo = 4,
    RetryCancel = 5,
  }

  export const enum IconType {
    None = 0,
    Critical = 16,
    Question = 32,
    Exclamation = 48,
    Information = 64,
  }

  export const enum DefaultButton {
    DefaultButton1 = 0,
    DefaultButton2 = 256,
    DefaultButton3 = 512,
  }

  export const enum PopupResult {
    OK = 1,
    Cancel = 2,
    Abort = 3,
    Retry = 4,
    Ignore = 5,
    Yes = 6,
    No = 7,
    Timeout = -1,
  }

  export const enum RegType {
    REG_SZ = 'REG_SZ',
    REG_EXPAND_SZ = 'REG_EXPAND_SZ',
    REG_DWORD = 'REG_DWORD',
    REG_BINARY = 'REG_BINARY',
  }

  export const enum RunWindowStyle {
    InvisibleDeactivate = 0,
    VisibleActivateRestore = 1,
    ActivateMinimize = 2,
    ActivateMaximize = 3,
    VisibleUpdate = 4,
    ActivateRestore = 5,
    MinimizeActivateNext = 6,
    Minimize = 7,
    VisibleRestore = 8,
    VisibleActivateRestoreFromMinimized = 9,
    BasedOnParent = 10
  }

  export const enum ShortcutWindowStyle {
    VisibleActivateRestore = 1,
    ActivateMaximize = 3,
    MinimizeActivateNext = 7,
  }

  export const enum ExecState {
    Running = 0,
    Finished = 1,
  }
}
