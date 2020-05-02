Operation =1
Option =0
Begin InputTables
    Name ="tbl_VH_Transect"
End
Begin OutputColumns
    Expression ="tbl_VH_Transect.Transect_ID"
    Expression ="tbl_VH_Transect.Event_ID"
    Expression ="tbl_VH_Transect.Transect"
    Expression ="tbl_VH_Transect.Visit_Date"
    Expression ="tbl_VH_Transect.Observer"
    Expression ="tbl_VH_Transect.Recorder"
End
Begin OrderBy
    Expression ="tbl_VH_Transect.Transect"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb9306c3f1b982a4bb4b1935ba4a83494
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_VH_Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Visit_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Recorder"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =16
    Top =140
    Right =1000
    Bottom =453
    Left =-1
    Top =-1
    Right =960
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =47
        Top =6
        Right =191
        Bottom =150
        Top =0
        Name ="tbl_VH_Transect"
        Name =""
    End
End
