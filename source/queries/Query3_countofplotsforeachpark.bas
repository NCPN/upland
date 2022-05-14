Operation =1
Option =0
Begin InputTables
    Name ="tbl_Revisit_List"
End
Begin OutputColumns
    Expression ="tbl_Revisit_List.PARK"
    Alias ="CountOfPlot"
    Expression ="Count(tbl_Revisit_List.Plot)"
End
Begin Groups
    Expression ="tbl_Revisit_List.PARK"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xedc013d70061c44c98c97f6780aad665
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Revisit_List.PARK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfPlot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c5b5e46d7d72140814f3f0fa8f6da27
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1564
    Bottom =836
    Left =-1
    Top =-1
    Right =1540
    Bottom =391
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Revisit_List"
        Name =""
    End
End
