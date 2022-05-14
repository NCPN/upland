Operation =4
Option =0
Begin InputTables
    Name ="tbl_Revisit_List"
End
Begin OutputColumns
    Name ="tbl_Revisit_List.ParkPlot"
    Expression ="[PARK] & [Plot]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x9da29831eeca694babd015068358e290
End
Begin
    Begin
        dbText "Name" ="tbl_Revisit_List.ParkPlot"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-535
    Top =189
    Right =1029
    Bottom =985
    Left =-1
    Top =-1
    Right =1540
    Bottom =391
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =507
        Top =195
        Right =651
        Bottom =339
        Top =0
        Name ="tbl_Revisit_List"
        Name =""
    End
End
