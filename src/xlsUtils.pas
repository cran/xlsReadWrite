unit xlsUtils;

{ Helpers.
                              ---
  The contents of this file may be used under the terms of the GNU General
  Public License Version 2 (the "GPL"). As a special exception I (copyright
  holder) allow to link against flexcel (http://www.tmssoftware.com/flexcel.htm).
                              ---
  The software is provided in the hope that it will be useful but without any
  express or implied warranties, including, but not without limitation, the
  implied warranties of merchantability and fitness for a particular purpose.
                              ---
  Copyright (C) 2006 by Hans-Peter Suter, Treetron GmbH, Switzerland.
  All rights reserved.
                              ---                                              }

{==============================================================================}
interface
uses
  SysUtils, rhRInternals, rhTypesAndConsts;

type
  ExlsReadWrite = class( Exception );

  aOutputType = ( otUndefined, otDouble, otInteger, otLogical
                , otCharacter, otDataFrame, otNumeric );
  aRowNameKind =( rnNA, rnTrue, rnFalse, rnSupplied );

const
  TheNAString =     'NA';   // see remark in pro vesion
  TheNaNString =    'NaN';

  TheOutputType: array[aOutputType] of string
                 = ( 'undefined', 'double', 'integer', 'logical'
                   , 'character', 'data.frame', 'numeric' );
  TheRowNameKind:array[aRowNameKind] of string
                 = ( 'NA', 'True', 'False', 'Supplied' );

function DateTimeToStrFmt( const _format: string; _dateTime: TDateTime ): string;
                 
function StrToOutputType( const _type: string ): aOutputType;
function AllOutputTypes(): string;

function VarAsBool( const _v: variant; _default: boolean ): boolean;
function VarAsDouble( const _v: variant; _default: double ): double; overload;
function VarAsDouble( const _v: variant; _default: double; _nullvalue: double ): double; overload;
function VarAsInt( const _v: variant; _default: integer ): integer;
function VarAsString( const _v: variant ): string; overload;
function VarAsString( const _v: variant; const _def: string ): string; overload;

function ProgFilename: string;

function GetFileVersion( const FileName : string;
    var Major, Minor, Release, Build : integer;
    var PreRelease, Debug : boolean;
    var Description : string ) : boolean;
function ReplaceVersionAndBuild( const _s: string ): string;


function GetScalarString( _val: pSExp; const _err: string ): string;

function AsFactor( _val: pSExp ): pSExp; cdecl;
function MakeNames( _names: pSExp ): pSExp; cdecl;

{==============================================================================}
implementation
uses
  Windows, Variants;

function DateTimeToStrFmt( const _format: string; _dateTime: TDateTime ): string;
  begin
    DateTimeToString( result, _format, _dateTime );
  end;

function StrToOutputType( const _type: string ): aOutputType;
  var
    i: aOutputType;
  begin
    for i:= Low( aOutputType) to High( aOutputType ) do begin
      if SameText( _type, theOutputType[i] ) then begin
        result:= i;
        Exit;
      end;
    end {for};
    result:= otUndefined;
  end {StrToOutputType};

function AllOutputTypes(): string;
  var
    i: aOutputType;
  begin
    result:= '';
    for i:= Succ( Low( aOutputType) ) to High( aOutputType ) do begin
      result:= result + ' / ' + theOutputType[i];
    end {for};
    if Length( result ) > 3 then Delete( result, 1, 3 );
  end {AllOutputTypes};

function VarAsBool( const _v: variant; _default: boolean ): boolean;
  begin
    case VarType( _v ) of
      varBoolean:       result:= _v;

      varSmallint,
      varInteger,
      varInt64,
      varByte,
      varWord,
      varLongWord:      result:= _v <> 0;

      varSingle,
      varDouble,
      varCurrency,
      varDate:          result:= Trunc( _v ) <> 0;

      varOleStr,
      varString:        result:= StrToBoolDef( _v, _default );
    else
      result:= _default;
    end;
  end {VarAsBool};

function VarAsInt( const _v: variant; _default: integer ): integer;
  begin
    case VarType( _v ) of
      varShortInt,
      varSmallint,
      varInteger,
      varInt64,
      varByte,
      varWord,
      varLongWord,
      varBoolean:       result:= _v;

      varSingle,
      varDouble,
      varCurrency,
      varDate:          result:= Trunc( _v );

      varOleStr,
      varString:        result:= StrToIntDef( _v, _default );
    else
      result:= _default;
    end {case};
  end {VarAsInt};

function VarAsDouble( const _v: variant; _default: double ): double;
  begin
    result:= VarAsDouble( _v, _default, _default );
  end {VarAsDouble};

function VarAsDouble( const _v: variant; _default, _nullvalue: double ): double;
  begin
    case VarType( _v ) of
      varSmallint,
      varInteger,
      varSingle,
      varDouble,
      varCurrency,
      varDate,
      varBoolean,
      varShortInt,
      varByte,
      varWord,
      varLongWord,
      varInt64:           result:= _v;

      varEmpty,
      varNull: 	          result:= _nullvalue;
    else
      result:= _default;
    end {case};
  end {VarAsDouble};

function VarAsString( const _v: variant ): string;
  begin
    result:= VarAsString( _v, '' );
  end {VarAsString};

function VarAsString( const _v: variant; const _def: string ): string;
  begin
    if VarIsNull( _v ) or VarIsEmpty( _v ) or (VarType(_v) = varError) then begin
      result:= _def;
    end else if VarType(_v) = varDate then begin
      result:= DateTimeToStr( VarToDateTime( _v ) );
    end else begin
      result:= string(_v);
    end;
  end {VarAsString};

function ProgFilename: string;
  begin
    SetLength( result, 255 );
    if IsLibrary then begin
      Windows.GetModuleFileName( HInstance, pChar(result), 255 );
    end else begin
      Windows.GetModuleFileName( 0, pChar(result), 255 );
    end;
    SetLength( result, StrLen( pChar(result) ) );
  end;

function GetFileVersion( const FileName : string;
    var Major, Minor, Release, Build : integer;
    var PreRelease, Debug : boolean;
    var Description : string ) : boolean;
  var
    zero: DWORD;       // set to 0 by GetFileVersionInfoSize
    versionInfoSize: DWORD;
    pVersionData: pointer;
    pFixedFileInfo: PVSFixedFileInfo;
    fixedFileInfoLength: UINT;
    fileFlags: WORD;
  begin
    result:= False;
      // ask Windows how big a data buffer to allocate to hold this EXE or DLL version info
    versionInfoSize:=
        GetFileVersionInfoSize( pChar(FileName), zero );
      // if no version info in the EXE or DLL
    if versionInfoSize = 0 then begin
        Exit;
    end;
      // allocate memory needed to hold version info
    pVersionData:= AllocMem( versionInfoSize );
    try
        // load version resource out of EXE or DLL into our buffer
      if GetFileVersionInfo( pChar(FileName), 0, versionInfoSize,
        pVersionData ) = FALSE then begin
        Exit;
      end;

        // get the fixed file info portion of the resource in buffer
      if VerQueryValue( pVersionData, '\', pointer(pFixedFileInfo),
          fixedFileInfoLength ) = False
      then begin
          // no fixed file info in this version resource !
        Exit;
      end;
        // extract the info from the the fixed file data structure
      Major := pFixedFileInfo^.dwFileVersionMS shr 16;
      Minor := pFixedFileInfo^.dwFileVersionMS and $FFFF;
      Release := pFixedFileInfo^.dwFileVersionLS shr 16;
      Build := pFixedFileInfo^.dwFileVersionLS and $FFFF;

      fileFlags :=  pFixedFileInfo^.dwFileFlags;
      PreRelease := (VS_FF_PRERELEASE and fileFlags) <> 0;
      Debug := (VS_FF_DEBUG and fileFlags) <> 0;

      Description := Format(
          'Ver %d.%d, Release %d Build %d',
          [Major, Minor, Release, Build] );

      if PreRelease then begin
          Description := Description + ' Beta';
      end;
      if Debug then begin
          Description := Description + ' Debug';
      end;
      result := True;
    finally
      FreeMem( pVersionData );
    end;
  end;

function ReplaceVersionAndBuild( const _s: string ): string;
  var
    v, b: string;
    major, minor, release, build : integer;
    preRelease, debug : boolean;
    description : string;
  begin
    if GetFileVersion( ProgFileName, major, minor, release, build,
        preRelease, debug, description )
    then begin
      v:= IntToStr( major ) + '.' + IntToStr( minor ) + '.' + IntToStr( release );
      if preRelease then v:= v + '-BETA';
      b:= IntToStr( build );
    end else begin
      v:= '<unknown>';
      b:= '<missing>';
    end;
    result:= StringReplace( _s, '@version@', v, [] );
    result:= StringReplace( result, '@build@', b, [] );
  end;

function GetScalarString( _val: pSExp; const _err: string ): string;
  begin
    if riIsString( _val ) and (riLength( _val ) = 1) then begin
      result:= riChar( riStringElt( _val, 0 ) );
    end else begin
      raise ExlsReadWrite.Create( _err );
    end;
  end;

function AsFactor( _val: pSExp ): pSExp; cdecl;
  var
    fcall: pSExp;
  begin
    fcall:= riProtect( riLang2( riInstall( 'as.factor' ), _val ) );
    result:= riProtect( riEval( fcall, RGlobalEnv ) );
    riUnprotect( 2 );
  end {AsFactor};

function MakeNames( _names: pSExp ): pSExp; cdecl;
  var
    fcall: pSExp;
  begin
    fcall:= riProtect( riLang2( riInstall( 'make.names' ), _names ) );
    result:= riProtect( riEval( fcall, RGlobalEnv ) );
    riUnprotect( 2 );
  end {MakeNames};

end {xlsUtils}.
