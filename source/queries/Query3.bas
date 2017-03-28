dbMemo "SQL" ="SELECT t.ID, Version, IsSupported, FieldCheck, Context, Syntax, TemplateName, Da"
    "taScope, Params, Template, Remarks, NumRecords\015\012FROM tsys_Db_Templates AS "
    "t LEFT JOIN NumRecords AS n ON n.ID = t.ID\015\012WHERE IsSupported > 0 \015\012"
    "AND (EffectiveDate < Date() OR EffectiveDate = Date())\015\012AND (RetireDate > "
    "Date() OR RetireDate IS NULL)\015\012AND FieldCheck = 1;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdc7877d58ec3b6459724dc7e4fcb1c78
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="DataScope"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Params"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Template"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remarks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumRecords"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fuels_1000.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Fuels_Transects.FuelsTransect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Belt_Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfCensus_ID"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic_Freq_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fuels_1000_Hr_Count"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Version"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsSupported"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldCheck"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Context"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Syntax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TemplateName"
        dbLong "AggregateType" ="-1"
    End
End
