Operation =1
Option =0
Begin InputTables
    Name ="qaz_sel_OT_Census_Report_c"
    Name ="qaz_sel_OT_Census_Report_b"
End
Begin OutputColumns
    Expression ="qaz_sel_OT_Census_Report_b.ParkPlot"
    Expression ="qaz_sel_OT_Census_Report_b.Visit_Year"
    Alias ="LongestNote"
    Expression ="qaz_sel_OT_Census_Report_b.MaxOfLofN2"
End
Begin Joins
    LeftTable ="qaz_sel_OT_Census_Report_c"
    RightTable ="qaz_sel_OT_Census_Report_b"
    Expression ="qaz_sel_OT_Census_Report_c.ParkPlot = qaz_sel_OT_Census_Report_b.ParkPlot"
    Flag =1
    LeftTable ="qaz_sel_OT_Census_Report_c"
    RightTable ="qaz_sel_OT_Census_Report_b"
    Expression ="qaz_sel_OT_Census_Report_c.MaxOfVisit_Year = qaz_sel_OT_Census_Report_b.Visit_Ye"
        "ar"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4bdd9c9a600caa4da545a099172b61ba
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="LongestNote"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3afc65cdb8b9f44dbcd8bc0767ddb997
        End
    End
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_b.ParkPlot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_b.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =133
    Top =63
    Right =1114
    Bottom =677
    Left =-1
    Top =-1
    Right =957
    Bottom =284
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =137
        Top =67
        Right =320
        Bottom =211
        Top =0
        Name ="qaz_sel_OT_Census_Report_c"
        Name =""
    End
    Begin
        Left =439
        Top =110
        Right =583
        Bottom =254
        Top =0
        Name ="qaz_sel_OT_Census_Report_b"
        Name =""
    End
End
