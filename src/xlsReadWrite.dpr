library xlsReadWrite;

{ The contents of this file may be used under the terms of the GNU General 
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

{$R *.RES}  // include version information
uses
  Windows,
  rhTypesAndConsts,
  rhR,
  rhRInternals,
  rhuLoadRVars,
  xlsRegister in 'xlsRegister.pas',
  xlsUtils in 'xlsUtils.pas',
  xlsWrite in 'xlsWrite.pas',
  xlsRead in 'xlsRead.pas';

var
  DllProcNext: procedure( _reason: integer ) = nil;

const
  theStartupMsg =
    'xlsReadWrite version @version@ (Build @build@)' + #13#10 +
    'Copyright (C) 2006, Hans-Peter Suter, Treetron, Switzerland.' + #13#10 +
    '' + #13#10 +
    'This package can be freely distributed and used for any' + #13#10 +
    'purpose. It comes with ABSOLUTELY NO GUARANTEE at all.' + #13#10 +
    'Code written by myself is licensed under GPLv2 with an ' + #13#10 +
    'exception to link certain proprietary code. See LICENSE.' + #13#10 +
    '' + #13#10 +
    'Kindly travel to http://treetron.googlepages.com for' + #13#10 +
    'infos, updates, the pro version, suggestions, support' + #13#10 +
    'and YES, why not make a donation if it serves you well...' + #13#10#13#10;

procedure MyDllProc( _reason: integer );
  var
    loadok: boolean;
  begin
    case _reason of
      DLL_PROCESS_ATTACH: begin
        loadok:= LoadRVars( ToRVarsArr( [vriRNilValue, vriRDimnamesSymbol,
            vriRRowNamesSymbol, vriRNamesSymbol, vriRLevelsSymbol] ) );
        if not LoadRVars( ToRVarsArr( [varRNaN, varRNaInt] ) ) then loadok:= False;
        if not loadok then begin
          rRprintf( 'Load xlsReadWrite.dll: Could not initialize RNilValue/RNaN' );
        end;
        rRprintf( pChar(ReplaceVersionAndBuild( theStartupMsg )), #13#10 );
      end;
      DLL_PROCESS_DETACH: begin
        { Here the console is already gone and a message window pops up if
          we use rRprintf. We have to use the R_unload_xlsReadWrite procedure
          which we register the same way as ReadXls and WriteXls (in R2.2.1
          the proc for some reason doesn't get called, but in R2.3.1 it is ok) }
      end;
    end {case};
    if Assigned( DllProcNext ) then DllProcNext( _reason );
  end {DllMain};


{ Exports }

exports ReadXls;
exports WriteXls;

exports R_init_xlsReadWrite;
exports R_unload_xlsReadWrite;

{==============================================================================}
begin {InitializeDemo}
  DllProcNext:= pointer( InterlockedExchange( integer(@DllProc), integer(@MyDllProc) ));
  MyDllProc( DLL_PROCESS_ATTACH );
end {xlsReadWrite}.
