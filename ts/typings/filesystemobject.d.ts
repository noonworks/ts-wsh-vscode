//
//  Types for FileSystemObject
//
// FileSystemObject https://msdn.microsoft.com/ja-jp/library/cc428071.aspx

interface ActiveXObject {
  // Add specialized signatures
  new (s: 'Scripting.FileSystemObject'): FileSystemObject;
}

interface FileSystemObject {
  Drives: FileSystemObject.Drives;
  BuildPath(path: string, name: string): string;
  CopyFile(source: string, destination: string, overwrite?: boolean): void;
  CopyFolder(source: string, destination: string, overwrite?: boolean): void;
  CreateFolder(foldername: string): FileSystemObject.Folder;
  CreateTextFile(filename: string, overwrite?: boolean, unicode?: boolean): FileSystemObject.TextStream;
  DeleteFile(filespec: string, force?: boolean): void;
  DeleteFolder(folderspec: string, force?: boolean): void;
  DriveExists(drivespec: string): boolean;
  FileExists(filespec: string): boolean;
  FolderExists(folderspec: string): boolean;
  GetAbsolutePathName(pathspec: string): string;
  GetBaseName(path: string): string;
  GetDrive(drivespec: string): FileSystemObject.Drive;
  GetDriveName(path: string): string;
  GetExtensionName(path: string): string;
  GetFile(filespec: string): FileSystemObject.File;
  GetFileName(pathspec: string): string;
  GetFolder(folderspec: string): FileSystemObject.Folder;
  GetParentFolderName(path: string): string;
  GetSpecialFolder(folderspec: FileSystemObject.SpecialFolders): FileSystemObject.Folder;
  GetTempName(): string;
  MoveFile(source: string, destination: string): void;
  MoveFolder(source: string, destination: string): void;
  OpenTextFile(filename: string, iomode?: FileSystemObject.OpenMode,
    create?: boolean, format?: FileSystemObject.FileFormat): FileSystemObject.TextStream;
}

declare namespace FileSystemObject {
  export const enum OpenMode {
    ForReading = 1,
    ForWriting = 2,
    ForAppending = 8,
  }

  export const enum FileFormat {
    TristateUseDefault = -2,
    TristateTrue = -1,
    TristateFalse = 0,
  }

  export const enum SpecialFolders {
    WindowsFolder = 0,
    SystemFolder = 1,
    TemporaryFolder = 2,
  }

  export const enum FolderAttributes {
    Normal = 0,
    ReadOnly = 1,
    Hidden = 2,
    System = 4,
    Volume = 8,
    Directory = 16,
    Archive = 32,
    Alias = 64,
    Compressed = 128,
  }

  interface FSOItem {
    Attributes: FolderAttributes;
    readonly DateCreated: Date;
    readonly DateLastAccessed: Date;
    readonly DateLastModified: Date;
    readonly Drive: string;
    Name: string;
    readonly ParentFolder: Folder;
    readonly Path: string;
    readonly ShortName: string;
    readonly ShortPath: string;
    readonly Size: number;
    readonly Type: string;
    Copy(destination: string, overwrite?: boolean): void;
    Delete(force?: boolean): void;
    Move(destination: string): void;
  }

  export interface Drive {
    readonly AvailableSpace: number;
    readonly DriveLetter: string;
    readonly DriveType: number;
    readonly FileSystem: string;
    readonly FreeSpace: number;
    readonly IsReady: boolean;
    readonly Path: string;
    readonly RootFolder: string;
    readonly SerialNumber: string;
    readonly ShareName: string;
    readonly TotalSize: number;
    VolumeName: string;
  }

  export interface Folder extends FSOItem {
    readonly Files: Files;
    readonly IsRootFolder: boolean;
    readonly SubFolders: Folders;
  }

  export interface File extends FSOItem {
    OpenAsTextStream(iomode?: OpenMode, format?: FileFormat): TextStream;
  }

  interface FSOCollection<T> {
    readonly Count: number;
    Item(index: number): T;
  }

  export interface Drives extends FSOCollection<Drive> {}

  export interface Files extends FSOCollection<File> {}

  export interface Folders extends FSOCollection<Folder> {
    Add(folderName: string): Folder;
  }

  export interface TextStream extends TextStreamReader, TextStreamWriter {}
}
