//
//  Types for ADO
//
// ADO Const Enums  https://msdn.microsoft.com/ja-jp/library/cc408225.aspx

declare namespace ADO {
  const enum ConnectModeEnum {
    adModeRead = 1,
    adModeReadWrite = 3,
    adModeRecursive = 0x400000,
    adModeShareDenyNone = 16,
    adModeShareDenyRead = 4,
    adModeShareDenyWrite = 8,
    adModeShareExclusive = 12,
    adModeUnknown = 0,
    adModeWrite = 2,
  }

  const enum LineSeparatorsEnum {
    adCR = 13,
    adCRLF = -1,
    adLF = 10,
  }

  const enum ObjectStateEnum {
    adStateClosed = 0,
    adStateOpen = 1,
    adStateConnecting = 2,
    adStateExecuting = 4,
    adStateFetching = 8,
  }

  const enum SaveOptionsEnum {
    adSaveCreateNotExist = 1,
    adSaveCreateOverWrite = 2,
  }

  const enum StreamOpenOptionsEnum {
    adOpenStreamAsync = 1,
    adOpenStreamFromRecord = 4,
    adOpenStreamUnspecified = -1,
  }

  const enum StreamReadEnum {
    adReadAll = -1,
    adReadLine = -2,
  }

  const enum StreamTypeEnum {
    adTypeBinary = 1,
    adTypeText = 2,
  }

  const enum StreamWriteEnum {
    adWriteChar = 0,
    adWriteLine = 1,
  }
}
