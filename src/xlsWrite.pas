unit xlsWrite;

{ Write functionality.
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

function WriteXls( _data, _file, _sheet, _colNames, _skipLines: pSExp ): pSExp; cdecl;

{==============================================================================}
implementation
uses
  Windows, SysUtils, Variants, Classes, UFlexCelImport, xlsUtils, XlsAdapter,
  rhR;

type
  aColheadertype = ( chtNone, chtLogical, chtString );

function WriteXls( _data, _file, _sheet, _colNames, _skipLines: pSExp ): pSExp; cdecl;
  var
    writer: TFlexCelImport;
    colcnt, rowcnt, offsetRow: integer;
    colheadertype: aColheadertype;
    rowNameAsFirstCol: boolean;

procedure SelectOrInsertSheet();
  var
    i, sheetIdx: integer;
    sheetName: string;
  begin
    if riIsNumeric( _sheet ) then begin
      sheetIdx := riInteger( riCoerceVector( _sheet, setIntSxp ) )[0];
      if (sheetIdx < 1) or (sheetIdx > writer.SheetCount) then begin
        raise ExlsReadWrite.Create('Sheet index must be between 1 and number of sheets');
      end;
      writer.ActiveSheet := sheetIdx;
    end else if riIsString(_sheet) then begin
      sheetName:= riChar( riStringElt( _sheet, 0 ) );
      if sheetName = '' then begin
        writer.ActiveSheet:= 1;
      end else begin
        for i:= 1 to writer.SheetCount do begin
          writer.ActiveSheet:= i;
          if SameText(writer.ActiveSheetName, sheetName) then Break;
        end;
        if not SameText(writer.ActiveSheetName, sheetName) then begin
          writer.InsertEmptySheets( 1, 1 );
          writer.ActiveSheet:= 1;
          writer.ActiveSheetName:= sheetName;
        end {if};
      end {if};
    end else begin
      raise ExlsReadWrite.Create('sheet must be of type numeric or string');
    end {if};
  end {SelectOrInsertSheet};

procedure WriteDouble(); cdecl;
  var
    r, c: integer;
  begin
    for r := 0 to rowcnt - 1 do begin
      for c := 0 to colcnt - 1 do begin
        writer.CellValue[r + 1 + offsetRow, c + 1]:= riReal( _data )[r + rowcnt*c];
      end {for};
    end {for};
  end {WriteDouble};

procedure WriteInteger(); cdecl;
  var
    r, c: integer;
  begin
    for r := 0 to rowcnt - 1 do begin
      for c := 0 to colcnt - 1 do begin
        writer.CellValue[r + 1 + offsetRow, c + 1]:= riInteger( _data )[r + rowcnt*c];
      end {for};
    end {for};
  end {WriteInteger};

procedure WriteLogical(); cdecl;
  var
    r, c: integer;
  begin
    for r := 0 to rowcnt - 1 do begin
      for c := 0 to colcnt - 1 do begin
        writer.CellValue[r + 1 + offsetRow, c + 1]:= riLogical( _data )[r + rowcnt*c];
      end {for};
    end {for};
  end {WriteLogical};

procedure WriteString(); cdecl;
  var
    r, c: integer;
  begin
    for r := 0 to rowcnt - 1 do begin
      for c := 0 to colcnt - 1 do begin
        writer.CellValue[r + 1 + offsetRow, c + 1]:=
          string(riChar( riStringElt( _data, r + rowcnt*c ) ));
      end {for};
    end {for};
  end {WriteString};

procedure WriteDataframe(); cdecl;
  var
    coltypes: array of aSExpType;
    lev: array of pSExp;
    r, c: integer;
    myrownames: pSExp;
  begin
    myrownames:= nil;                  
      { are there rownames which could become the first column }
    if colheadertype <> chtNone then begin
      myrownames:= riGetAttrib( _data, RRowNamesSymbol );
      rowNameAsFirstCol:= (not riIsNull( myrownames)) and
          (riTypeOf( myrownames ) = setStrSxp) and
          (string(riChar( riStringElt( myrownames, 0 ) )) <> '1' );
    end;
    SetLength( coltypes, colcnt + integer(rowNameAsFirstCol) );
    SetLength( lev, colcnt + integer(rowNameAsFirstCol) );

      { loop columns (set type) }
    if rowNameAsFirstCol then begin
      coltypes[0]:= setStrSxp;
      lev[0]:= nil;
    end;
    for c:= integer(rowNameAsFirstCol) to colcnt - 1 + integer(rowNameAsFirstCol) do begin
      lev[c - integer(rowNameAsFirstCol)]:= nil;
      coltypes[c - integer(rowNameAsFirstCol)]:=
          riTypeOf( riVectorElt( _data, c - integer(rowNameAsFirstCol) ) );
        { treat factors separately }
      if coltypes[c - integer(rowNameAsFirstCol)] = setIntSxp then begin
        if riIsFactor( riVectorElt( _data, c - integer(rowNameAsFirstCol) ) ) then begin
          coltypes[c - integer(rowNameAsFirstCol)]:= setCplxSxp;  // WARNING: misuse of setCplxSxp !!!
          lev[c - integer(rowNameAsFirstCol)]:=
              riGetAttrib( riVectorElt( _data, c - integer(rowNameAsFirstCol) ), RLevelsSymbol );
        end;
      end;
    end {for};

      { loop rows (write data) }
    for r:= 0 to rowcnt - 1 do begin
        { first column (rownames) }
      if Assigned( myrownames ) and rowNameAsFirstCol then begin
        writer.CellValue[r + 1 + offsetRow, 1]:=
          string(riChar( riStringElt( myrownames, r ) ));
      end;
        { data columns }
      for c:= integer(rowNameAsFirstCol) to colcnt - 1 + integer(rowNameAsFirstCol) do begin
        case coltypes[c - integer(rowNameAsFirstCol)] of
          setIntSxp: begin
            writer.CellValue[r + 1 + offsetRow, c + 1]:=
                riInteger( riVectorElt( _data, c - integer(rowNameAsFirstCol) ) )[r];
          end;
          setCplxSxp: begin  // setCplxSxp used for factors (WARNING: levels 1-based, riStringElt 0-based)
            writer.CellValue[r + 1 + offsetRow, c + 1]:=
                string(riChar( riStringElt( lev[c - integer(rowNameAsFirstCol)],
                riInteger( riVectorElt( _data, c - integer(rowNameAsFirstCol) ) )[r] - 1 ) ));
          end;
          setRealSxp: begin
            writer.CellValue[r + 1 + offsetRow, c + 1]:=
                riReal( riVectorElt( _data, c - integer(rowNameAsFirstCol) ) )[r];
          end;
          setLglSxp: begin
            writer.CellValue[r + 1 + offsetRow, c + 1]:=
                riLogical( riVectorElt( _data, c - integer(rowNameAsFirstCol) ) )[r];
          end;
          setStrSxp: begin
            writer.CellValue[r + 1 + offsetRow, c + 1]:=
                string(riChar( riStringElt( riVectorElt( _data, c - integer(rowNameAsFirstCol) ), r ) ));
          end;
        else
          assert( True, 'WriteDataframe: coltype not supported' );
        end {case};
      end {for each column};
    end {for each row};
  end {WriteDataframe};

  var
    i: integer;
    colnames, dimnames: pSExp;
    filename: string;
    skipLines: integer;
  begin {WriteXls}
    result:= RNilValue;
    try
      filename:= riChar( riStringElt( _file, 0 ) );
      colheadertype:= chtNone;
      if riIsLogical( _colNames ) then begin
        if riLogical( _colNames )[0] <> 0 then colheadertype:= chtLogical;
      end else if riIsString( _colNames ) then begin
        if riLength( _colNames ) = colcnt then begin
          colheadertype:= chtString;
        end else begin
          rWarning( 'Length of character colheader must be equal to length of ' +
              'columns (%d). Logical colheader will be used.', [colcnt] );
          colheadertype:= chtLogical;
        end;
      end else begin
        raise ExlsReadWrite.Create('colHeader must be of type logical or string');
      end {if colHeader};

      skipLines:= riInteger( riCoerceVector( _skipLines, setIntSxp ) )[0];

        { create writer }
      writer:= TFlexCelImport.Create();
      writer.Adapter:= TXLSAdapter.Create();
      try
          { new or existing file }
        writer.OpenFile( ExtractFileDir( ProgFilename ) + '\TemplateNew.xls' );
        SelectOrInsertSheet();

          { row and column count }
        offsetRow:= skipLines;
        if riIsFrame( _data ) then begin
          rowcnt:= riLength( riVectorElt( _data, 0 ) );
        end else begin
          rowcnt:= riNrows( _data );
        end;
        if rowcnt > 65536 then raise ExlsReadWrite.CreateFmt( 'Only up to %d rows supported (Excel <V2007))', [65536] );
        if riIsFrame( _data ) then begin
          colcnt:= riLength( _data );
        end else begin
          colcnt:= riNcols( _data );
        end;
        if colcnt > 256 then raise ExlsReadWrite.CreateFmt( 'Only up to %f columns supported (Excel <V2007))', [256] );
        if colheadertype <> chtNone then Inc( offsetRow );

        { -- write matrix }

        case riTypeOf( _data ) of
          setRealSxp:   WriteDouble();
          setIntSxp:    WriteInteger();
          setLglSxp:    WriteLogical();
          setStrSxp:    WriteString();
          setVecSxp:    begin
            if not riIsFrame( _data ) then begin
              raise ExlsReadWrite.Create( 'Currently the following types ' +
                  'are supported: ' + #13#13 +
                  'REALSXP (double), INTSXP (integer), LGLSXP (logical) ' +
                  'STRSXP (character), VECSXP (data.frame)' );
            end;
            WriteDataframe();
          end {setVecSxp}
        else
          raise ExlsReadWrite.Create( 'Currently only the following types ' +
              'are supported: ' + #13#13 +
              'REALSXP (double), INTSXP (integer), LGLSXP (logical) ' +
              'STRSXP (character), VECSXP (data.frame or list)' );
        end {case};

          { column header }
        if colheadertype <> chtNone then begin
          colnames:= nil;
          if colheadertype = chtString then begin
            colnames:= _colNames;
          end else if colheadertype = chtLogical then begin
            if riTypeOf( _data ) = setVecSxp then begin
                { frame }
              colnames:= riGetAttrib( _data, RNamesSymbol );
              if riIsNull( colnames ) then colnames:= nil;
            end else begin
                { matrix }
              dimnames:= riGetAttrib( _data, RDimNamesSymbol );
              if not riIsNull( dimnames ) then begin
                colnames:= riVectorElt( dimnames, 1 );
                if riIsNull( colnames ) then colnames:= nil;
              end;
            end;
          end {if};

          for i:= integer(rowNameAsFirstCol) to colcnt - 1 + integer(rowNameAsFirstCol) do begin
            if Assigned( colnames ) then begin
              writer.CellValue[skiplines + 1, i + 1]:=
                  string(riChar( riStringElt( colnames, i - integer(rowNameAsFirstCol) ) ));
            end else begin
              writer.CellValue[skiplines + 1, i + 1]:=
                  'V' + IntToStr( i - integer(rowNameAsFirstCol) + 1 );
            end {if};
          end {for};
        end {if};

          { close }
        if FileExists(FileName) then SysUtils.DeleteFile( filename );
        writer.Save( filename );
        writer.CloseFile;
      finally
        writer.Free;
      end {try};
    except
      on E: ExlsReadWrite do begin
        rError( pChar(E.Message) );
      end;
      on E: Exception do begin
        rError( pChar('Unexpected error. Message: ' + E.Message) );
      end;
    end {try};
  end;

end {xlsWrite}.
