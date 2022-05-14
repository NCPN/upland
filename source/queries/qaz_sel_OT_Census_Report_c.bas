Operation =1
Option =0
Begin InputTables
    Name ="qaz_sel_OT_Census_Report_b"
End
Begin OutputColumns
    Expression ="qaz_sel_OT_Census_Report_b.ParkPlot"
    Alias ="MaxOfVisit_Year"
    Expression ="Max(qaz_sel_OT_Census_Report_b.Visit_Year)"
End
Begin Groups
    Expression ="qaz_sel_OT_Census_Report_b.ParkPlot"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x390638cb740c0144bb59bb1d93762174
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_b.ParkPLot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfVisit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x99836f0a100a9c4eb83658d10986aa79
        End
    End
End
Begin
    State =0
    Left =-6
    Top =28
    Right =1437
    Bottom =751
    Left =-1
    Top =-1
    Right =1419
    Bottom =267
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =418
        Top =71
        Right =562
        Bottom =215
        Top =0
        Name ="qaz_sel_OT_Census_Report_b"
        Name =""
    End
End
