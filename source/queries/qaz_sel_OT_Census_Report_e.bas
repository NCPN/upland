Operation =1
Option =0
Where ="(((qaz_sel_OT_Census_Report_d.LongestNote)>74))"
Begin InputTables
    Name ="tbl_Revisit_List"
    Name ="qaz_sel_OT_Census_Report_d"
End
Begin OutputColumns
    Expression ="qaz_sel_OT_Census_Report_d.ParkPlot"
    Expression ="qaz_sel_OT_Census_Report_d.Visit_Year"
    Expression ="qaz_sel_OT_Census_Report_d.LongestNote"
End
Begin Joins
    LeftTable ="tbl_Revisit_List"
    RightTable ="qaz_sel_OT_Census_Report_d"
    Expression ="tbl_Revisit_List.ParkPlot = qaz_sel_OT_Census_Report_d.ParkPlot"
    Flag =1
End
Begin OrderBy
    Expression ="qaz_sel_OT_Census_Report_d.LongestNote"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1273046cc1ec444a92d133cab7f41466
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_d.ParkPlot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_d.LongestNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qaz_sel_OT_Census_Report_d.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-529
    Top =192
    Right =914
    Bottom =915
    Left =-1
    Top =-1
    Right =1419
    Bottom =216
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Revisit_List"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =457
        Bottom =156
        Top =0
        Name ="qaz_sel_OT_Census_Report_d"
        Name =""
    End
End
