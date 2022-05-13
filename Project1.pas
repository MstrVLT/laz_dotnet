program Project1;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Classes, SysUtils, CustApp, ActiveX, ComObj, dynlibs
  { you can add units after this };

type

  { TMyApplication }

  TMyApplication = class(TCustomApplication)
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
  end;


  PDotnetEnvironmentSDKInfo = ^TDotnetEnvironmentSDKInfo;
  TDotnetEnvironmentSDKInfo = packed record
    Size    : SIZE_T;
    version : PWideChar;
    path    : PWideChar;
  end;

  PDotnetEnvironmentFrameworkInfo = ^TDotnetEnvironmentFrameworkInfo;
  TDotnetEnvironmentFrameworkInfo = packed record
    Size    : SIZE_T;
    name    : PWideChar;
    version : PWideChar;
    path    : PWideChar;
  end;

  PDotnetEnvironmentInfo = ^TDotnetEnvironmentInfo;
  TDotnetEnvironmentInfo = packed record
    Size                : SIZE_T;
    hostfxr_version     : PWideChar;
    hostfxr_commit_hash : PWideChar;
    sdk_count           : SIZE_T;
    sdks                : PDotnetEnvironmentSDKInfo;
    framework_count     : SIZE_T;
    frameworks          : PDotnetEnvironmentFrameworkInfo;
  end;

  PComActivationContext = ^TComActivationContext;
  TComActivationContext = packed record
    ClassId      : TGuid;
    InterfaceId  : TGuid;
    AssemblyPath : PWideChar;
    AssemblyName : PWideChar;
    TypeName     : PWideChar;
    ComActivator : PUnknown;
  end;

  PInitializeParameters = ^TInitializeParameters;
  TInitializeParameters = packed record
    Size       : SIZE_T;
    HostPath   : PWideChar;
    DotnetRoot : PWideChar;
  end;

  // ComActivator.cs
  TComDelegateFn = function(const cxt: PComActivationContext): Int32; stdcall;

  // hostfxr.h
  TErrorWriterFn = procedure(const Msg: PWideChar); cdecl;

  TSetErrorWriterFn = function(
      ErrorWriter: TErrorWriterFn): Int32; cdecl;

  TGetDotnetEnvironmentInfo = procedure(
      DotnetRoot : PWideChar;
      Reserved: Pointer;
      Result: Pointer;
      ResultContext: Pointer); cdecl;

  TInitializeForRuntimeConfigFn = function(
      const RuntimeConfigPath: PWideChar;
      const Parameters: PInitializeParameters;
      var HostContextHandle: Pointer): Int32; cdecl;

  TGetRuntimeDelegateFn = function(
      const HostContextHandle: Pointer;
      DelegateType: Int32;
      var Delegate: TComDelegateFn): Int32; cdecl;

  TCloseFn = function(
      const HostContextHandle: Pointer): Int32; cdecl;

  ICalculator = interface(IUnknown)
    ['{158C2477-6F7D-4B7A-BC58-DDF7AC098109}']
    function Sum(A, B: Integer; C: PInteger): HRESULT; stdcall;
  end;

  procedure ErrorWriter(const Msg: PWideChar); cdecl;
  begin
    writeln('Error: ', WideCharToString(Msg));
  end;

{ TMyApplication }

procedure TMyApplication.DoRun;
var
  ErrorMsg: String;

  Handle: HMODULE;
  SetErrorWriter: TSetErrorWriterFn;
  GetDotnetEnvironmentInfo: TGetDotnetEnvironmentInfo;
  InitializeForRuntimeConfig: TInitializeForRuntimeConfigFn;
  GetRuntimeDelegateFn: TGetRuntimeDelegateFn;
  CloseFn: TCloseFn;

  result: integer;
  ctx: pointer;
  ComDelegate: TComDelegateFn;

  ComActivator: IUnknown;
  ClassFactory: IClassFactory;

  Calculator: ICalculator;
  rec: TComActivationContext;
  c: Integer;

  InitializeParameters: TInitializeParameters;
begin
  Handle := LoadLibrary('{ full path } \dotnet\host\fxr\5.0.16\hostfxr.dll');
  if Handle <> dynlibs.NilHandle then
  begin
    try

      Pointer(SetErrorWriter) := GetProcAddress(Handle, 'hostfxr_set_error_writer');
      Pointer(GetDotnetEnvironmentInfo) := GetProcAddress(Handle, 'hostfxr_get_dotnet_environment_info');
      Pointer(InitializeForRuntimeConfig) := GetProcAddress(Handle, 'hostfxr_initialize_for_runtime_config');
      Pointer(GetRuntimeDelegateFn) := GetProcAddress(Handle, 'hostfxr_get_runtime_delegate');
      Pointer(CloseFn) := GetProcAddress(Handle, 'hostfxr_close');

      ctx := nil;

      InitializeParameters.Size := SizeOf(TInitializeParameters);

      InitializeParameters.HostPath := '{ full path } \ClassLibrary1.dll';
      InitializeParameters.DotnetRoot := '{ full path } \dotnet';

      result := SetErrorWriter(@ErrorWriter);
      result := InitializeForRuntimeConfig('{ full path } \ClassLibrary1.runtimeconfig.json', @InitializeParameters, ctx);

      if ((result <> 0) or not Assigned(ctx)) then
      begin
        CloseFn(ctx); Exit;
      end;
                                                                     
      ComDelegate := nil;
      result := GetRuntimeDelegateFn(ctx, 0{hdt_com_activation}, ComDelegate);

      if ((result <> 0) or not Assigned(ComDelegate)) then
      begin
        CloseFn(ctx); Exit;
      end;

      result := CloseFn(ctx);

      rec.ClassId := GUID_NULL;
      rec.InterfaceId := IClassFactory;
      rec.AssemblyPath := '{ full path } \ClassLibrary1.dll';
      rec.AssemblyName := 'ClassLibrary1';
      rec.TypeName := 'ClassLibrary1.Calculator';

      ComActivator := nil;
      rec.ComActivator := @ComActivator;

      ComDelegate(@rec);

      if Supports(ComActivator, IClassFactory, ClassFactory) then
      begin
        ComActivator := nil;
        if Succeeded(ClassFactory.CreateInstance(nil, ICalculator, Calculator)) then
        begin
          ClassFactory := nil;
          if Succeeded(Calculator.Sum(1, 2, @C)) then
            writeln('1 + 2 = ', c);
        end;
      end;
    finally
      FreeLibrary(Handle);
    end;
  end;

  // stop program loop
  Terminate;
end;

constructor TMyApplication.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  StopOnException:=True;
end;

destructor TMyApplication.Destroy;
begin
  inherited Destroy;
end;

var
  Application: TMyApplication;
begin
  Application:=TMyApplication.Create(nil);
  Application.Title:='My Application';
  Application.Run;
  Application.Free;
end.

