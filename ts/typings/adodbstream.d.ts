//
//  Types for ADODB.Stream
//
// ADODB.Stream     https://msdn.microsoft.com/ja-jp/library/cc364272.aspx

interface ActiveXObject {
  // Add specialized signatures
  new (s: 'ADODB.Stream'): ADODBStream;
}

interface ADODBStream {
  Charset: ADODBStream.Charsets | string;
  readonly EOS: boolean;
  LineSeparator: ADO.LineSeparatorsEnum;
  Mode: ADO.ConnectModeEnum;
  Position: number;
  readonly Size: number;
  State: ADO.ObjectStateEnum;
  Type: ADO.StreamTypeEnum;
  Cancel(): void;
  Close(): void;
  CopyTo(DestStream: ADODBStream, NumChars?: number): void;
  Flush(): void;
  LoadFromFile(FileName: string): void;
  Open(Source?: any, Mode?: ADO.ConnectModeEnum, OpenOptions?: ADO.StreamOpenOptionsEnum, UserName?: string, Password?: string): void;
  Read(NumBytes?: number | ADO.StreamReadEnum): any;
  ReadText(NumChars?: number | ADO.StreamReadEnum): string;
  SaveToFile(FileName: string, SaveOptions?: ADO.SaveOptionsEnum): void;
  SetEOS(): void;
  SkipLine(): void;
  Write(Buffer: any): void;
  WriteText(Data: string, Options?: ADO.StreamWriteEnum): void;
}

declare namespace ADODBStream {
  const enum Charsets {
    Unicode = 'Unicode',
    utf8 = 'utf-8',
    shift_jis = 'shift_jis',
    eucjp = 'euc-jp',
  }
}
