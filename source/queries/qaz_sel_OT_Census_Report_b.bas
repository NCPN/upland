Operation =1
Option =2
Begin InputTables
    Name ="qaz_sel_OT_Census_Report_a"
End
Begin OutputColumns
    Expression ="qaz_sel_OT_Census_Report_a.ParkPlot"
    Expression ="qaz_sel_OT_Census_Report_a.Visit_Year"
    Alias ="MaxOfLofN2"
    Expression ="Max(qaz_sel_OT_Census_Report_a.LofN2)"
End
Begin OrderBy
    Expression ="qaz_sel_OT_Census_Report_a.ParkPlot"
    Flag =0
    Expression ="qaz_sel_OT_Census_Report_a.Visit_Year"
    Flag =0
End
Begin Groups
    Expression ="qaz_sel_OT_Census_Report_a.ParkPlot"
    GroupLevel =0
    Expression ="qaz_sel_OT_Census_Report_a.Visit_Year"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0c5d147069703e43bc79a47c6a1cfb23
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_a.ParkPlot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_a.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfLofN2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x343e8887bdd0b547a868f059d9be5c06
        End
    End
End
Begin
    State =0
    Left =-34
    Top =74
    Right =1409
    Bottom =797
    Left =-1
    Top =-1
    Right =1419
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =407
        Top =74
        Right =610
        Bottom =218
        Top =0
        Name ="qaz_sel_OT_Census_Report_a"
        Name =""
    End
End
