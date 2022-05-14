Operation =1
Option =0
Begin InputTables
    Name ="tbl_VH_Intercept"
End
Begin OutputColumns
    Expression ="tbl_VH_Intercept.Intercept_ID"
    Expression ="tbl_VH_Intercept.Transect_ID"
    Expression ="tbl_VH_Intercept.Point"
    Expression ="tbl_VH_Intercept.Wood"
    Expression ="tbl_VH_Intercept.WHeight"
    Expression ="tbl_VH_Intercept.Herb"
    Expression ="tbl_VH_Intercept.HHeight"
End
Begin OrderBy
    Expression ="tbl_VH_Intercept.Point"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xea1ccc01aa3a24449981aa7406c55f5a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_VH_Intercept.Herb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.HHeight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Point"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Intercept_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Wood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.WHeight"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1648
    Bottom =853
    Left =-1
    Top =-1
    Right =1536
    Bottom =410
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =251
        Top =0
        Name ="tbl_VH_Intercept"
        Name =""
    End
End
