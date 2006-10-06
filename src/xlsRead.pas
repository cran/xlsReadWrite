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

function ReadXls( _file, _sheet, _type, _colNames, _skipLines, _colClasses: pSExp ): pSExp; cdecl;

{==============================================================================}
implementation
uses
  SysUtils, Variants, Classes, xlsUtils, UFlexCelImport, XlsAdapter, rhR;

function ReadXls( _file, _sheet, _type, _colNames, _skipLines, _colClasses: pSExp ): pSExp; cdecl;
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

procedure SetColNames(_colNames: pSExp);
  var
    i: integer;
  begin
    if riIsLogical( _colNames ) then begin
      hasColNames:= riLogical( _colNames )[0] <> 0;
      SetLength( colNames, 0 );
    end else if riIsString( _colNames ) then begin
      SetLength( colNames, riLength( _colNames ) );
      for i := 0 to riLength( _colNames ) - 1 do begin
        colNames[i]:= string(riChar( riStringElt( _colNames, i ) ));
      end;
      hasColNames:= Length( colNames ) > 0;
    end else begin
      raise ExlsReadWrite.Create('SetColNames: "colNames" must be of type logical or string');
    end {if colHeader};
  end;

procedure ReadColNames( _idx: integer );
  var
    i: integer;
  begin
    if hasColNames and (Length( colNames ) > 0) then begin
      if Length( colNames ) <> colcnt then begin
        raise EXlsReadWrite.CreateFmt( 'colNames must be a vector of ' +
          'equal length than the column count (length: %d/colcnt: %d)',
          [Length( colNames ), colCnt] );
      end;
      Exit;
    end;
    SetLength( colnames, colcnt );
    for i:= 0 to colcnt - 1 do begin
      if hasColNames then begin
        colnames[i]:= VarAsString( reader.CellValue[_idx, i + 1], '' );
      end else begin
        colnames[i]:= '';
      end;
    end;
  end {SetColNames};

procedure ApplyColNames(_idx: integer);
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
    hasColClasses: boolean;
    firstColAsRowName: boolean;

    { it's a bit a hack but I don't have the nice classes
      from pro and don't want to change too many things }
  procedure SetColClasses(_colClasses: pSExp);

    function StrToColType( const _type: string ): aSExpType;
      begin
        if _type = 'double' then begin
          result:= setRealSxp;
        end else if _type = 'integer' then begin
          result:= setIntSxp;
        end else if _type = 'logical' then begin
          result:= setLglSxp;
        end else if (_type = 'character') or (_type = 'factor') then begin
          result:= setStrSxp;
        end else if _type = 'rowname' then begin
          result:= setRawSxp;  // misuse !!
        end else if _type = 'NA' then begin
          result:= setNilSxp;  // 1st try to find a type, 2nd use RNaInt
        end else begin
          raise EXlsReadWrite.CreateFmt( '"%s" is not a valid colClasses entry ' +
              '(use double, integer, logical, character, factor or NA)', [_type] );
        end;
      end {StrToColType};

    var
      i: integer;
    begin {SetColClasses}

        { check if is NA scalar }
      if (riLength( _colClasses ) = 1) and
         (riTypeOf( _colClasses ) in [setLglSxp, setRealSxp]) and
         (rIsNa( riReal( riCoerceVector( _colClasses, setRealSxp ) )[0] ) <> 0)
      then begin
        hasColClasses:= False;
      end else begin

        if riIsString(_colClasses) then begin
          hasColClasses:= True;
          SetLength( coltypes, colcnt );
            { scalar }
          if riLength( _colClasses ) = 1 then begin
            if string(riChar( riStringElt( _colClasses, 0 ) )) = 'rowname' then begin
              raise EXlsReadWrite.Create( '"rowname" may not be used for a scalar colClasses argument' );
            end;
            for i:= 0 to colcnt - 1 do begin
              coltypes[i]:= StrToColType( string(riChar( riStringElt( _colClasses, 0 ) )) );
            end;
            { vector }
          end else if riLength( _colClasses ) = colCnt then begin
            for i:= 0 to colcnt - 1 do begin
              coltypes[i]:= StrToColType( string(riChar( riStringElt( _colClasses, i ) )) );
            end;
          end else begin
            raise EXlsReadWrite.CreateFmt( 'colClasses must be a scalar or a vector of ' +
              'equal length than the column count (length: %d/colcnt: %d)',
              [riLength( _colClasses ), colCnt] );
          end;
        end else begin
           raise ExlsReadWrite.Create( 'colClasses must be NA or a string (vector)' );
        end {if};
      end;
    end {SetColClasses};

  var
    r, c, i: integer;
    v: variant;
    myrownames, myclass, mynames: pSExp;
    tempname: string;

  begin {ReadDataframe}
    SetLength( coltypes, 0 );
    SetColClasses( _colClasses );

      { support rowname }
    if hasColClasses then begin
        { check in colClasses }
      assert( Length( coltypes ) > 0, 'ReadDataframe: coltypes must be longer than zero' );
      firstColAsRowName:= coltypes[0] = setRawSxp;
      if firstColAsRowName then coltypes:= Copy( coltypes, 1, Length( coltypes ) - 1 );
        { rowname must be at the beginning }
      for i:= 0 to Length( coltypes ) - 1 do if coltypes[i] = setRawSxp then begin
        raise ExlsReadWrite.Create( '"rownames" can only be indicated at the first position in colClasses' );
      end;
    end else begin
      { check for autorow-column: -there are column header, - the first columnheader
        is empty and - first value in the first column is *not* the string '1' }
      firstColAsRowName:= hasColNames and (colnames[0] = '') and
          (length( colnames ) > 1) and
          VarIsStr( reader.CellValue[1 + offsetRow, 1] ) and
          (reader.CellValue[1 + offsetRow, 1] <> '1');
      SetLength( coltypes, colcnt - integer(firstColAsRowName) );
      for i:= 0 to Length( coltypes ) - 1 do coltypes[i]:= setNilSxp;
    end {if};

      { allocate }
    result:= riProtect( riAllocVector( setVecSxp, colcnt - integer(firstColAsRowName) ) );
    mynames:= riProtect( riAllocVector( setStrSxp, colcnt - integer(firstColAsRowName) ) );
    myrownames:= riProtect( riAllocVector( setStrSxp, rowcnt ) );

    { loop columns (get type and name) }

    for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin

      if coltypes[c] <> setNilSxp then begin

          { type already determined }
        case coltypes[c] of
          setRealSxp:    riSetVectorElt( result, c, riAllocVector( setRealSxp, rowcnt ) );
          setIntSxp:     riSetVectorElt( result, c, riAllocVector( setIntSxp, rowcnt ) );
          setLglSxp:     riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
          setStrSxp:     riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
        else
          assert( False, 'coltype not supported (bug)' );
        end {case};
      end else begin

          { read row and determine type }
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
            coltypes[c]:= setNilSxp;  // riLogical and RNaInt will be used;
            riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
        end {case};
      end {if};

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
  begin {ReadXls}
    result:= RNilValue;
    SetLength( colnames, 0 );
    try
      offsetRow:= riInteger( riCoerceVector( _skipLines, setIntSxp ) )[0];
      SetColNames( _colNames );

        { create reader }
      reader:= TFlexCelImport.Create();
      reader.Adapter:= TXLSAdapter.Create();
      try
          { open existing file }
        reader.OpenFile( riChar( riStringElt( _file, 0 ) ) );
        SelectSheet();

          { counts and offsets }
        rowcnt:= reader.MaxRow;
        colcnt:= reader.MaxCol;
        if hasColNames and (Length( colnames ) = 0) then Inc( offsetRow );
        rowcnt:= rowcnt - offsetRow;

          { read column header (empty if not hasColNames ) }
        ReadColNames( offsetRow );

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
            ApplyColNames( offsetRow );
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
