unit rhuLoadRVars;

{ Initialize global variables of the R header files. LoadPredefinedRVars()
  loads the variables according to the constants in the implementation section.
  Warning: It (probably) only works if the code runs in the same process.
                              ---
  R is a computer language for statistical data analysis and is published
  under the GPL, the above-mentioned headerfile is distributed under LGPL.
  Copyright (C) 1995, 1996  Robert Gentleman and Ross Ihaka
  Copyright (C) 1999-2006   The R Development Core Team.
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
  SysUtils, rhTypesAndConsts;

type
  EInitRVarsError = class( Exception );


function LoadRVars( const _arr: aRVarsArr; _hdl: THandle = 0 ): boolean;
procedure RequireRVars(const _arr: aRVarsArr; _hdl: THandle = 0);

{==============================================================================}
implementation
uses
  Windows, rhRInternals;

const

{---------------------------------------------------------- Define variables here }

  { Important: the types have to be pSExp !! }

    { from rhRDefines }
  TheLoadNullUserObject = True; // will be initialized with RNilValue if true


{---------------------------------------------------------- load functions }

function LoadRVars( const _arr: aRVarsArr; _hdl: THandle = 0 ): boolean;
  var
    dllHdl: THandle;
    i: integer;
    myppSExp: ^pSExp;
    mypDouble: pDouble;
    mypInteger: pInteger;
  begin
    result:= True;

    if _hdl <> 0 then dllHdl:= _hdl else dllHdl:= Windows.LoadLibrary( TheRDll );

    if dllHdl <> 0 then begin
      for i:= 0 to Length( _arr ) - 1 do begin
        if _arr[i].gvType = vtSExp then begin
          myppSExp:= GetProcAddress( dllHdl, pChar(_arr[i].gvName) );
          if integer(myppSExp) <> 0 then begin
            pointer(_arr[i].gvPointer^):= myppSExp^;
          end else begin
            result:= False;
          end;
        end else if _arr[i].gvType = vtDouble then begin
          mypDouble:= GetProcAddress( dllHdl, pChar(_arr[i].gvName) );
          if integer(mypDouble) <> 0 then begin
            double(_arr[i].gvPointer^):= mypDouble^;
          end else begin
            result:= False;
          end;
        end else if _arr[i].gvType = vtInteger then begin
          mypInteger:= GetProcAddress( dllHdl, pChar(_arr[i].gvName) );
          if integer(mypInteger) <> 0 then begin
            integer(_arr[i].gvPointer^):= mypInteger^;
          end else begin
            result:= False;
          end;
        end;
      end {for};

        { only free handle if loaded here }
      if _hdl = 0 then FreeLibrary( dllHdl );
    end else begin
      result:= False;
    end {if handle};
  end {LoadRVars};

procedure RequireRVars(const _arr: aRVarsArr; _hdl: THandle = 0);
  begin
    if not LoadRVars( _arr, _hdl ) then begin
      raise EInitRVarsError.Create( 'RequireRVars: Could not load a R variable' );
    end;
  end {RequireRVars};


end {rhuLoadRVars}.
