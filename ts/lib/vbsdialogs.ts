module VBSDialogs {
  declare function escape(s: string): string;
  type ValidateFunction = (v: any) => any;

  interface MSScriptControl {
    Language: string;
    AddCode(code: string): void;
    Eval(code: string): any;
    Run(procedureName: string, parameters: any[]): any;
  }

  function toVBStringParam(str: string): string {
    return 'Unescape("' + escape(str + "") + '")';
  }

  function createLangObj(): MSScriptControl {
    const obj: MSScriptControl = WScript.CreateObject('MSScriptControl.ScriptControl');
    obj.Language = 'vbscript';
    return obj;
  }

  export function inputBox(msg: string, title: string, placeholder: string = '', cancel_value: string = '', validator: ValidateFunction = null): any {
    // InputBox(prompt, title, default) https://msdn.microsoft.com/ja-jp/library/cc410238.aspx
    placeholder = placeholder.replace(/[\n\r\\]/g, '');
    const obj= createLangObj();
    obj.AddCode('function g() g = InputBox(' + toVBStringParam(msg) + ',' + toVBStringParam(title) + ',' + toVBStringParam(placeholder) + ') end function');
    const val = obj.Eval('g');
    if (typeof(val) == 'undefined') return cancel_value;
    if (validator != null) return validator(val);
    return val;
  }

  export function msgBox(msg: string, title: string, buttons: string = 'vbOKOnly'): any {
    // MsgBox(prompt, buttons, title) https://msdn.microsoft.com/ja-jp/library/cc410277.aspx
    const obj = createLangObj();
    obj.AddCode('function g() g = MsgBox(' + toVBStringParam(msg) + ',' + buttons + ',' + toVBStringParam(title) + ') end function');
    return obj.Eval('g');
  }
}
