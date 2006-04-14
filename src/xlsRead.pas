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
  rhRInternals, rhTypesAndConsts, xlsUtils;

function ReadXls( _file, _colNames, _sheet, _type,
    _from, _rowNames, _colClasses,
    _checkNames, _dateTimeAs,
    _stringsAsFactors: pSExp ): pSExp; cdecl;

{==============================================================================}
implementation
uses
  SysUtils, Variants, Classes, Math, UFlexCelImport, XlsAdapter, rhR, UFlxNumberFormat,
  UFlxFormats;

function ReadXls( _file, _colNames, _sheet, _type,
    _from, _rowNames, _colClasses,
    _checkNames, _dateTimeAs,
    _stringsAsFactors: pSExp ): pSExp; cdecl;
  var
    reader: TFlexCelImport;
    colcnt, rowcnt, from: integer;
    hasColNames: boolean;
    colnames: array of string;
    rownames: array of string;
    rownameKind: aRowNameKind;
    checkNames: boolean;
    dateTimeAsNumeric: boolean;
    stringsAsFactors: boolean;
    firstColAsRowName: boolean;

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
      sheetName:= GetScalarString(_sheet, 'sheet must be a character string' );
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

procedure SetTrueFalse(_val: pSExp; var _res: boolean; const _what: string );
  begin
    if riIsLogical( _val ) and (riLength( _val ) = 1) then begin
      _res:= riLogical( _val )[0] <> 0;
    end else begin
      raise ExlsReadWrite.Create( '"' + _what + '" must be TRUE or FALSE' );
    end;
  end;

procedure SetDateTimeAs( _dateTimeAs: pSExp );
  begin
    dateTimeAsNumeric:= GetScalarString( _dateTimeAs,
        'dateTimeAs must be a character string' ) = 'numeric';
  end;

procedure SetRowNames(_colNames: pSExp);
  var
    i: integer;
  begin
        { check if is NA scalar }
    if (riLength( _rowNames ) = 1) and
       (riTypeOf( _rowNames ) in [setLglSxp, setRealSxp]) and
       (rIsNa( riReal( riCoerceVector( _rowNames, setRealSxp ) )[0] ) <> 0)
    then begin
      rownameKind:= rnNA;
    end else if riIsLogical( _rowNames ) then begin
      if riLogical( _rowNames )[0] <> 0 then rownameKind:= rnTrue else rownameKind:= rnFalse;
      SetLength( rowNames, 0 );
    end else if riIsString( _rowNames ) then begin
      rownameKind:= rnSupplied;
      SetLength( rowNames, riLength( _rowNames ) );
      for i := 0 to riLength( _rowNames ) - 1 do begin
        rowNames[i]:= string(riChar( riStringElt( _rowNames, i ) ));
      end;
      if not Length( rowNames ) > 0 then rownameKind:= rnFalse;
    end else begin
      raise ExlsReadWrite.Create('SetRowNames: "rowNames" must be of type logical or string');
    end;
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
        colnames[i]:= VarAsString( reader.CellValue[_idx - 1, i + 1], '' );
      end else begin
        colnames[i]:= '';
      end;
    end;
  end {SetColNames};

procedure ApplyMatrixRowColNames( _result: pSExp);
  var
    dim: pSExp;
    i: integer;
    r: integer;
    myrownames, mycolnames: pSExp;
  begin
    dim:= riProtect( riAllocVector( setVecSxp, 2 ) );

      { rownames }
    if firstColAsRowName or (rownameKind = rnSupplied) then begin
      myrownames:= riProtect( riAllocVector( setStrSxp, rowcnt ) );
      if rownameKind = rnSupplied then begin
        for r:= 0 to rowcnt - 1 do begin
          riSetStringElt( myrownames, r, riMkChar( pChar(rownames[r]) ) );
        end;
      end else begin
        for r:= 0 to rowcnt - 1 do begin
          if firstColAsRowName then begin
            riSetStringElt( myrownames, r, riMkChar( pChar(VarAsString(
                reader.CellValue[r + from, 1], IntToStr( r + 1 ) ) )) );
          end else begin
            riSetStringElt( myrownames, r, riMkChar( pChar(IntToStr( r + 1 )) ) );
          end;
        end {for};
      end;
    end else begin
      myrownames:= RNilValue;
    end;

      { colnames }
    if hasColNames then begin
      mycolnames:= riProtect( riAllocVector( setStrSxp, colcnt - integer(firstColAsRowName) ));
      for i:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin
        riSetStringElt( mycolnames, i, riMkChar( pChar(colnames[i + integer(firstColAsRowName)]) ) );
      end;
      if checkNames then mycolnames:= riProtect( MakeNames( mycolnames ) );
    end else begin
      mycolnames:= RNilValue;
    end;

      { apply }
    riSetVectorElt( dim, 0, myrownames );
    riSetVectorElt( dim, 1, mycolnames );
    riSetAttrib( result, RDimNamesSymbol, dim );

    riUnprotect( 1 + integer(CheckNames) );
    if firstColAsRowName or (rownameKind = rnSupplied) then riUnprotect( 1 );
    if hasColNames then riUnprotect( 1 );
  end {ApplyColNames};

procedure ApplyFrameRowNames( _result: pSExp);
  var
    r: integer;
    myrownames: pSExp;
  begin
    myrownames:= riProtect( riAllocVector( setStrSxp, rowcnt ) );
    if rownameKind = rnSupplied then begin
      for r:= 0 to rowcnt - 1 do begin
        riSetStringElt( myrownames, r, riMkChar( pChar(rownames[r]) ) );
      end;
    end else begin
      for r:= 0 to rowcnt - 1 do begin
        if firstColAsRowName then begin
          riSetStringElt( myrownames, r, riMkChar( pChar(VarAsString(
              reader.CellValue[r + from, 1], IntToStr( r + 1 ) ) )) );
        end else begin
          riSetStringElt( myrownames, r, riMkChar( pChar(IntToStr( r + 1 )) ) );
        end;
      end {for};
    end;
    riSetAttrib( _result, RRowNamesSymbol, myrownames );
    riUnprotect( 1 );
  end {ApplyRowNames};

function CheckForAutoRow: boolean;
  begin
    result:= hasColNames and (colnames[0] = '') and
        (length( colnames ) > 1) and
        VarIsStr( reader.CellValue[from, 1] ) and
        (reader.CellValue[from, 1] <> '1');
  end {CheckForAutoRow};

function ReadDouble(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setRealSxp, rowcnt, colcnt - integer(firstColAsRowName) ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin
        riReal( result )[r + rowcnt*c]:=
            VarAsDouble( reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], RNaN, RNaReal );
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadDouble};

function ReadInteger(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setIntSxp, rowcnt, colcnt - integer(firstColAsRowName) ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin
        riInteger( result )[r + rowcnt*c]:=
            VarAsInt( reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], RNaInt );
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadInteger};

function ReadLogical(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setLglSxp, rowcnt, colcnt - integer(firstColAsRowName) ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin
        riLogical( result )[r + rowcnt*c]:=
            integer(VarAsBool( reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], False ));
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadLogical};

function ReadString(): pSExp; cdecl;
  var
    r, c: integer;
  begin
    result:= riProtect( riAllocMatrix( setStrSxp, rowcnt, colcnt - integer(firstColAsRowName) ) );
    for r:= 0 to rowcnt - 1 do begin
      for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin
        riSetStringElt( result, r + rowcnt*c, riMkChar(
            pChar(VarAsString( reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], '' )) ) );
      end {for};
    end {for};
    riUnprotect( 1 );
  end {ReadString};

function ReadDataframe(): pSExp; cdecl;
  var
    coltypes: array of aSExpType;
    hasColClasses: boolean;

    { it's a bit a hack but I don't have the nice classes
      from pro and don't want to change too many things }
  procedure SetColClasses(_colClasses: pSExp);

    function StrToColType( const _type: string ): aSExpType;
      begin
        if (_type = 'double') or (_type = 'numeric' ) then begin
          result:= setRealSxp;
        end else if _type = 'integer' then begin
          result:= setIntSxp;
        end else if _type = 'logical' then begin
          result:= setLglSxp;
        end else if (_type = 'character') then begin
          result:= setStrSxp;
        end else if (_type = 'factor') then begin
          result:= setSpecialSxp;  // misuse of aSExpType
        end else if (_type = 'isodate') then begin
          result:= setCplxSxp;     // misuse of aSExpType
        end else if (_type = 'isotime') then begin
          result:= setDotSxp;      // misuse of aSExpType
        end else if (_type = 'isodatetime') then begin
          result:= setAnySxp;      // misuse of aSExpType
        end else if _type = 'NA' then begin
          result:= setNilSxp;      // 1st try to find a type, 2nd use RNaInt
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
          if (riLength( _colClasses ) = colCnt) or
              (riLength( _colClasses ) = colCnt - integer(firstColAsRowName))
          then begin
            SetLength( coltypes, riLength( _colClasses ) );
            for i:= 0 to Length( coltypes ) - 1 do begin
              coltypes[i]:= StrToColType( string(riChar( riStringElt( _colClasses, i ) )) );
            end;
          end else if riLength( _colClasses ) = 1 then begin
            SetLength( coltypes, colCnt );
            coltypes[0]:= StrToColType( string(riChar( riStringElt( _colClasses, 0 ) )) );
            for i:= 1 to colCnt - 1 do begin
              coltypes[i]:= coltypes[0];
            end;
          end else begin
            raise EXlsReadWrite.CreateFmt( 'colClasses must be a scalar or a vector of ' +
              'equal (or optionally with rownames equal - 1) length than the column count (length: %d/colcnt: %d)',
              [riLength( _colClasses ), colCnt] );
          end;
        end else begin
           raise ExlsReadWrite.Create( 'colClasses must be NA or a string (vector)' );
        end {if};
      end;
      if Length( coltypes ) = 0 then begin
        SetLength( coltypes, colcnt - integer(firstColAsRowName) );
        for i:= 0 to Length( coltypes ) - integer(firstColAsRowName) - 1 do begin
          coltypes[i]:= setNilSxp;
        end;
      end;
    end {SetColClasses};

  var
    r, c: integer;
    v: variant;
    myclass, mynames: pSExp;
    flexfmt: TFlxFormat;
    hasDate, hasTime: boolean;
    tempname: string;
    tempdat: double;
    formerfactors: pSExp;
    consideredRows: integer;
  begin {ReadDataframe}
    SetLength( coltypes, 0 );

    { support rowname }

    firstColAsRowName:= False;
    if rownameKind = rnTrue then begin
      firstColAsRowName:= True;
    end else if rownameKind = rnNA then begin
      if CheckForAutoRow then begin
        firstColAsRowName:= True;
      end;
    end;
    if firstColAsRowName and (colcnt < 2) then begin
      raise ExlsReadWrite.CreateFmt('If a column is used for the rownames, there must be  ' +
              'at least 2 columns in the Excelfile (actual number: %d)', [colcnt]);
    end;

    SetColClasses( _colClasses );
    if firstColAsRowName and (Length( coltypes ) = colcnt - integer(firstColAsRowName) + 1) then begin
        { need to remove column for rownames }
      coltypes:= Copy( coltypes, 1, Length( coltypes ) - 1 );
    end;

      { allocate }
    result:= riProtect( riAllocVector( setVecSxp, colcnt - integer(firstColAsRowName) ) );
    mynames:= riProtect( riAllocVector( setStrSxp, colcnt - integer(firstColAsRowName) ) );

    { loop columns (get type and name) }

   consideredRows:= Min( rowcnt - 1, 15 );

   for c:= 0 to colcnt - 1 - integer(firstColAsRowName) do begin

      if coltypes[c] <> setNilSxp then begin

          { type already determined }
        case coltypes[c] of
          setRealSxp:    riSetVectorElt( result, c, riAllocVector( setRealSxp, rowcnt ) );
          setIntSxp:     riSetVectorElt( result, c, riAllocVector( setIntSxp, rowcnt ) );
          setLglSxp:     riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
          setCplxSxp,
          setDotSxp,
          setAnySxp,
          setSpecialSxp,
          setStrSxp:     riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
        else
          assert( False, 'coltype not supported (bug)' );
        end {case};
      end else begin

          { read row and determine type }
        for r:= 0 to consideredRows do begin
          v:= reader.CellValue[from + r, c + 1 + integer(firstColAsRowName)];
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
              Break;
            end;
            varSingle,
            varDouble,
            varCurrency,
            varDate: begin
              if dateTimeAsNumeric then begin
                coltypes[c]:= setRealSxp;
                riSetVectorElt( result, c, riAllocVector( setRealSxp, rowcnt ) );
                Break;
              end else begin
                reader.GetCellFormatDef( from, c + 1 + integer(firstColAsRowName), flexfmt );
                HasXlsDateTime( flexfmt.Format, hasDate, hasTime );
                if hasDate and hasTime then begin
                  coltypes[c]:= setAnySxp; // WARNING: misuse
                  riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
                end else if hasDate then begin
                  coltypes[c]:= setCplxSxp; // WARNING: misuse
                  riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
                end else if hasTime then begin
                  coltypes[c]:= setDotSxp; // WARNING: misuse
                  riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
                end else begin
                  coltypes[c]:= setRealSxp;
                  riSetVectorElt( result, c, riAllocVector( setRealSxp, rowcnt ) );
                  Break;
                end;
              end {if use oleDateTime};
            end;
            varBoolean: begin
              coltypes[c]:= setLglSxp;
              riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
              Break;
            end;
            varOleStr,
            varString: begin
              if stringsAsFactors then begin
                coltypes[c]:= setSpecialSxp;  // misuse of aSExpType
              end else begin
                coltypes[c]:= setStrSxp;
              end;
              riSetVectorElt( result, c, riAllocVector( setStrSxp, rowcnt ) );
              Break;
            end;
          end {case};
        end {for considered rows};

        if coltypes[c] = setNilSxp then begin
          rWarning( pChar('Could not determine a column type from first ' + IntToStr( consideredRows ) + ' rows. Infos:' + #13#10 +
              '- colCnt: ' + IntToStr( colcnt ) + ', rowCnt: ' + IntToStr( rowcnt ) + ', ' +
              'rowIdx of first data row: ' + IntToStr( from ) + #13#10 +
              '- colIdx: ' + IntToStr( c + 1 ) + #13#10 +
              '"LOGICAL" will be assumed and all values will be NA' + #13#10 +
              '(Often it works if you delete the superfluous rows/columns (*not only* the cell *content*))' + #13#10#13#10));
          coltypes[c]:= setNilSxp;  // riLogical and RNaInt will be used;
          riSetVectorElt( result, c, riAllocVector( setLglSxp, rowcnt ) );
        end {if nothing found in considered rows};
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
                reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], RNaInt );
          end;
          setRealSxp: begin
            riReal( riVectorElt( result, c ) )[r]:= VarAsDouble(
                reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], RNaN, RNaReal );
          end;
          setCplxSxp, setDotSxp, setAnySxp: begin   // misused for Date
            tempdat:= VarAsDouble( reader.CellValue[r + from, c + 1 +
                integer(firstColAsRowName)], NaN, NaN );
            if not IsNan( tempdat ) then begin
              if coltypes[c] = setCplxSxp then begin
                tempname:= DateTimeToStrFmt( 'yyyy-mm-dd', tempdat );
              end else if coltypes[c] = setDotSxp then begin
                tempname:= DateTimeToStrFmt( 'hh:nn:ss', tempdat );
              end else begin
                tempname:= DateTimeToStrFmt( 'yyyy-mm-dd hh:nn:ss', tempdat );
              end;
            end else begin
              tempname:= '';
            end;
            riSetStringElt( riVectorElt( result, c ), r, riMkChar( pChar(tempname) ) );
          end;
          setLglSxp: begin
            riLogical( riVectorElt( result, c ) )[r]:= integer(VarAsBool(
                reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], False ));
          end;
          setSpecialSxp, setStrSxp: begin
            riSetStringElt( riVectorElt( result, c ), r, riMkChar(pChar(VarAsString(
                reader.CellValue[r + from, c + 1 + integer(firstColAsRowName)], '' )) ) );
          end;
          setNilSxp: begin
            riLogical( riVectorElt( result, c ) )[r]:= RNaInt;
          end;
        else
          assert( False, 'ReadDataframe: coltype not supported (bug)' );
        end {case};
      end {for each column};
    end {for each row};

      { convert to factor if necessary }
    for c:= 0 to Length( coltypes ) - 1 do begin
      if coltypes[c] = setSpecialSxp then begin
        formerfactors:= riVectorElt( result, c );
        riSetVectorElt( result, c, AsFactor( formerfactors ) );
      end;
    end {for each col};

    { make the frame }

    if checkNames then mynames:= riProtect( MakeNames( mynames ) );
    riSetAttrib( result, RNamesSymbol, mynames );
    myclass:= riProtect( riMkString( 'data.frame' ) );
    riClassgets( result, myclass );
    riUnprotect( 3 + integer(checkNames) );
  end {ReadDataframe};

  var
    outputtype: aOutputType;
  begin {ReadXls}
    result:= RNilValue;
    SetLength( colnames, 0 );
    SetLength( rownames, 0 );
    try
      from:= riInteger( riCoerceVector( _from, setIntSxp ) )[0];
      SetColNames( _colNames );
      SetRowNames( _rowNames );
      SetDateTimeAs( _dateTimeAs );
      SetTrueFalse( _checkNames, checkNames, 'checkNames' );
      SetTrueFalse( _stringsAsFactors, stringsAsFactors, 'stringsAsFactors' );

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
        if hasColNames and (Length( colnames ) = 0) then Inc( from );
        rowcnt:= rowcnt - from + 1;
        if (rownameKind = rnSupplied) and (not (length( rownames ) = rowcnt)) then begin
          raise ExlsReadWrite.CreateFmt('Number of supplied rownames doesn''t match ' +
              'the number of determined rows (%d/%d)', [length( rownames ), rowcnt]);
        end;

          { read column header (empty if not hasColNames ) }
        ReadColNames( from );

          { read matrix }
        if (rowcnt > 0) and (colcnt > 0) then begin
          outputtype:= StrToOutputType( riChar( riStringElt( _type, 0 ) ) );

            { data.frame }
          if outputtype = otDataFrame then begin
            result:= ReadDataframe();
            ApplyFrameRowNames( result )

            { matrix }
          end else begin
            if outputtype in [otDouble, otInteger, otLogical, otCharacter] then begin
              firstColAsRowName:= rownameKind = rnTrue;
              if not firstColAsRowName then begin
                if (rownameKind = rnNA) and CheckForAutoRow() then begin
                  firstColAsRowName:= True;
                end;
              end;

              case outputtype of
                otDouble, otNumeric:   result:= ReadDouble();
                otInteger:             result:= ReadInteger();
                otLogical:             result:= ReadLogical();
                otCharacter:           result:= ReadString();
              end;

                { colNames and rownames }
              if hasColNames or firstColAsRowName or (rownameKind = rnSupplied) then begin
                ApplyMatrixRowColNames( result );
              end;

            end else begin
              raise ExlsReadWrite.Create( 'The types "' + AllOutputTypes +
                  '" are supported right now. (Your input was: ' +
                  riChar( riStringElt( _type, 0 ) ) + ')' );
            end {if matrix};
          end {if data.frame};
        end else begin
          result:= RNilValue;
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
