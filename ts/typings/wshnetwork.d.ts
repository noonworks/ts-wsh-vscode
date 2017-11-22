//
//  Types for WshNetwork
//
// WshNetwork       https://msdn.microsoft.com/ja-jp/library/cc364454.aspx

interface ActiveXObject {
  // Add specialized signatures
  new (s: 'WScript.Network'): WshNetwork;
}

interface WshNetwork {
  readonly ComputerName: string;
  readonly UserDomain: string;
  readonly UserName: string;
  AddWindowsPrinterConnection(strPrinterPath: string, strDriverName?: string, strPort?: string): void;
  AddPrinterConnection(strLocalName: string, strRemoteName: string, bUpdateProfile?: boolean, strUser?: string, strPassword?: string): void;
  EnumNetworkDrives(): { readonly length: number; Item(key: number): string; };
  EnumPrinterConnections(): { readonly length: number; Item(key: number): string; };
  MapNetworkDrive(strLocalName: string, strRemoteName: string, bUpdateProfile?: boolean, strUser?: string, strPassword?: string): void;
  RemoveNetworkDrive(strName: string, bForce?: boolean, bUpdateProfile?: boolean): void;
  RemovePrinterConnection(strName: string, bForce?: boolean, bUpdateProfile?: boolean): void;
  SetDefaultPrinter(strPrinterName: string): void;
}
