unit xlsRead;

{ Read functionality.
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
  rhRInternals, rhTypesAndConsts;

function ReadXls( _file, _sheet, _type, _colNames, _skipLines: pSExp ): pSExp; cdecl;

{==============================================================================}
implementation
uses
  SysUtils, Variants, Classes, xlsUtils, UFlexCelImport, XlsAdapter, rhR;

function ReadXls( _file, _sheet, _type, _colNames, _skipLines: pSExp ): pSExp; cdecl;
  var
    reader: TFlexCelImport;
    colcnt, rowcnt, offsetRow: integer;
    colnames: array of string;
    hasColNames: boolean;

procedure SelectSheet();
  var
    i, sheetIdx: integer;
    sheetName: string;
  begin
    if riIsNumeric(_sheet) then begin
      sheetIdx:= riInteger( riCoerceVector( _sheet, setIntSxp ) )[0];
      if (sheetIdx < 1) or (sheetIdx > reader.SheetCount) then begin
        raise ExlsReadWrite.Create('Sheet index must be between 1 and number of sheets');
      end;
      reader.ActiveSheet:= sheetIdx;
    end else if riIsString(_sheet) then begin
      sheetName:= riChar(riStringElt(_sheet, 0));
      for i:= 1 to reader.SheetCount do begin
        reader.ActiveSheet:= i;
        if SameText(reader.ActiveSheetName, sheetName) then Break;
      end;
      if not SameText(reader.ActiveSheetName, sheetName) then begin
        raise ExlsReadWrite.CreateFmt('There is no sheet "%s" in the file "%s"',
            [sheetName, riChar( riStringElt( _file, 0 ) )]);
      end;
    end else begin
      raise ExlsReadWrite.Create('sheet must be of type numeric or string');
    end {if};
  end {SelectSheet};

procedure ReadAndSetHeader( _idx: integer );
  var
    headername: pSExp;
  begin
    headername:= riProtect( riAllocVector( setStrSxp, 1 ));
    riSetStringElt( headername, 0, riMkChar(
        pChar(VarAsString( reader.CellValue[_idx, 1], '' )) ) );
    riUnprotect( 1 );
  end {SetHeader};

procedure ReadColNames( _idx: integer );
  var
    i: integer;
  begin
    SetLength( colnames, colcnt );
    for i:= 0 to colcnt - 1 do begin
      if hasColNames then begin
        colnames[i]:= VarAsString( reader.CellValue[_idx, i + 1], '' );
      end else begin
        colnames[i]:= '';
      end;
    end;
  end {SetColNames};

procedure SetColNames(_idx: integer);
  var
    dim, col: pSExp;
    i: integer;
  begin
    col:= riProtect( riAllocVector( setStrSxp, colcnt ));
    for i:= 0 to colcnt - 1 do begin
      riSetStringElt( col, i, riMkChar( pChar(colnames[i]) ) );
    end;
    dim:= riProtect( riAllocVector( setVecSxp, 2 ) );
    riSetVectorElt( dim, 0, RNilValue );
    riSetVectorElt( dim, 1, col );
    riSetAttrib( result, RDimNamesSymbol, dim );
    riUnprotect( 2 );
  end {SetColNames};

function ReadDouble(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setRealSxp, rowcnt, colcnt ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 do begin
        riReal( result )[r + rowcnt*c]:=
            VarAsDouble( reader.CellValue[r + 1 + offsetRow, c + 1], RNaN, 0 );
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadDouble};

function ReadInteger(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setIntSxp, rowcnt, colcnt ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 do begin
        riInteger( result )[r + rowcnt*c]:=
            VarAsInt( reader.CellValue[r + 1 + offsetRow, c + 1], RNaInt );
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadInteger};

function ReadLogical(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setLglSxp, rowcnt, colcnt ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 do begin
        riLogical( result )[r + rowcnt*c]:=
            integer(VarAsBool( reader.CellValue[r + 1 + offsetRow, c + 1], False ));
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadLogical};

function ReadString(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setStrSxp, rowcnt, colcnt ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 do begin
        riSetStringElt( result, r + rowcnt*c, riMkChar(
            pChar(VarAsString( reader.CellValue[r + 1 + offsetRow, c + 1], '' )) ) );
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadString};

function ReadDataframe(): pSExp; cdecl;
  var
    coltypes: array of aSExpType;
    r, c: integer;
    v: variant;
    myrownames, myclass, mynames: pSExp;
    tempname: string;
    firstColAsRowName: boolean;
  begin
      { first column is used as rowname if
        - there are column header, - the first columnheader is empty and
        - first value in the first column is *not* the string '1' }
    firstColAsRowName:= hasColNames and (colnames[0] = '') and
        VarIsStr( reader.CellValue[1 + offsetRow, 1] ) and
        (reader.CellValue[1 + offsetRow, 1] <> '1');

    SetLength( coltypes, colcnt - integer(firstColAsRowName) );
    result:= riProtect( riAllocVector( setVecSxp, colcnt - integer(firstColAsRowName)) );
    mynames:= riProtect( riAllocVector( setStrSxp, colcnt - integer(firstColAsRowName) ) );
    myrownames:= riProtect( riAllocVector( setStrSxp, rowcnt ) );

    { loop columns (get type and name) }

    for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin

      v:= reader.CellValue[1 + offsetRow, c + 1 + integer(firstColAsRowName)];
      case VarType( v ) of
        varSmallint,
        varInteger,
        varShortInt,
        varByte,
        varWord,
        varLongWord,
        varInt64: begin
          coltypes[c]:= setIntSxp;
          riSetVectorElt( result, c, riAllocVector( setIntSxp, rowcnt ) );
        end;
        varSingle,
        varDouble,
        varCurrency: begin
          coltypes[c]:= setRealSxp;
          riSetVectorElt( result, c, riAllocVector( setRealSxp, rowcnt ) );
        end;
        varDate: begin
          coltypes[c]:= setCplxSxp; // WARNING: misuse of setCplxSxp !!!
          riSetVectorElt( result, c, riAllocVector( setRealSxp, rowcnt ) );
        end;
        varBoolean: begin
          coltypes[c]:= setLglSxp;
          riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
        end;
        varOleStr,
        varString: begin
          coltypes[c]:= setStrSxp;
          riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
        end;
        else
          tempname:= '';
          for r:= 0 to colcnt - 1 do tempname:= tempname + ', "' + colnames[r] + '"';
          if Length( tempname ) > 2 then Delete( tempname, 1, 2 );
          rWarning( pChar('Could not determine a column type.' + #13#10 +
              'The first data row *must* have valid entries for all columns. Infos:' + #13#10 +
              '- colCnt: ' + IntToStr( colcnt ) + ', rowCnt: ' + IntToStr( rowcnt ) + ', ' +
              'rowIdx of data row: ' + IntToStr( 0 + offsetRow + 1 ) + ', variant type of value: "' + VarTypeAsText( VarType( v ) ) + '"' + #13#10 +
              '- colHeaders: ' + tempname + ')' + #13#10 +
              '- colIdx: ' + IntToStr( c + 1 )  + #13#10 +
              '"LOGICAL" will be assumed and all values will be NA' + #13#10 +
              '(Maybe it works if you delete the superfluous columns (not only the cell content))' + #13#10#13#10 ));
          coltypes[c]:= setNilSxp;
          riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
      end {case};

        { set mynames (colnames) }
      tempname:= colnames[c + integer(firstColAsRowName)];
      if tempname = '' then tempname:= 'V' + IntToStr( c + 1 );
      riSetStringElt( mynames, c, riMkChar( pChar(tempname) ) );
    end {for each column};


    { loop rows (read data) }

    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin
        case coltypes[c] of
          setIntSxp: begin
            riInteger( riVectorElt( result, c ) )[r]:= VarAsInt(
                reader.CellValue[r + 1 + offsetRow, c + 1 + integer(firstColAsRowName)], RNaInt );
          end;
          setRealSxp, setCplxSxp: begin   // setCplxSxp used for Date (but currently not treated specially)
            riReal( riVectorElt( result, c ) )[r]:= VarAsDouble(
                reader.CellValue[r + 1 + offsetRow, c + 1 + integer(firstColAsRowName)], RNaN, 0 );
          end;
          setLglSxp: begin
            riLogical( riVectorElt( result, c ) )[r]:= integer(VarAsBool(
                reader.CellValue[r + 1 + offsetRow, c + 1 + integer(firstColAsRowName)], False ));
          end;
          setStrSxp: begin
            riSetStringElt( riVectorElt( result, c ), r, riMkChar(pChar(VarAsString(
                reader.CellValue[r + 1 + offsetRow, c + 1 + integer(firstColAsRowName)], '' )) ) );
          end;      
          setNilSxp: begin
            riLogical( riVectorElt( result, c ) )[r]:= RNaInt;
          end;
        else
          assert( False, 'ReadDataframe: coltype not supported (bug)' );
        end {case};
      end {for each column};

      if firstColAsRowName then begin
        riSetStringElt( myrownames, r, riMkChar( pChar(VarAsString(
            reader.CellValue[r + 1 + offsetRow, 1], IntToStr( r + 1 ) ) )) );
      end else begin
        riSetStringElt( myrownames, r, riMkChar( pChar(IntToStr( r + 1 )) ) );
      end;
    end {for each row};

    { make the frame }

    riSetAttrib( result, RNamesSymbol, mynames );
    myclass:= riProtect( riMkString( 'data.frame' ) );
    riClassgets( result, myclass );
    riUnprotect( 1 );
    riSetAttrib( result, RRowNamesSymbol, myrownames );
    riUnprotect( 3 );
  end {ReadDataframe};

  var
    outputtype: aOutputType;
    skipLines: integer;
  begin {ReadXls}
    result:= RNilValue;
    try
      hasColNames:= riLogical( _colNames )[0] <> 0;
      skipLines:= riInteger( riCoerceVector( _skipLines, setIntSxp ) )[0];

        { create reader }
      reader:= TFlexCelImport.Create();
      reader.Adapter:= TXLSAdapter.Create();
      try
          { open existing file }
        reader.OpenFile( riChar( riStringElt( _file, 0 ) ) );
        SelectSheet();

          { counts and offsets }
        offsetRow:= skipLines;
        rowcnt:= reader.MaxRow;
        colcnt:= reader.MaxCol;
        if hasColNames then Inc( offsetRow );
        rowcnt:= rowcnt - offsetRow;

          { read column header (empty if not hasColNames ) }
        ReadColNames( skiplines + 1 );

          { read matrix }
        outputtype:= otUndefined;
        if (rowcnt > 0) and (colcnt > 0) then begin
          outputtype:= StrToOutputType( riChar( riStringElt( _type, 0 ) ) );
          case outputtype of
            otDouble:         result:= ReadDouble();
            otInteger:        result:= ReadInteger();
            otLogical:        result:= ReadLogical();
            otCharacter:      result:= ReadString();
            otDataFrame:      result:= ReadDataframe();
            else begin
              raise ExlsReadWrite.Create( 'The types "' + AllOutputTypes +
                  '" are supported right now. (Your input was: ' +
                  riChar( riStringElt( _type, 0 ) ) + ')' );
            end {else};
          end {case};
        end else begin
          result:= RNilValue;
        end {if};

          { header and column header }
        if result <> RNilValue then begin
          if (outputtype <> otDataFrame) and hasColNames then begin
            SetColNames( skiplines + 1 );
          end;
        end {if};

        reader.CloseFile;
      finally
        reader.Free;
      end {try};
    except
      on E: ExlsReadWrite do begin
        rError( pChar(E.Message) );
      end;
      on E: Exception do begin
        rError( pChar('Unexpected error. Message: ' + E.Message) );
      end;
    end {try};
  end {ReadXls};

end {xlsRead}.
