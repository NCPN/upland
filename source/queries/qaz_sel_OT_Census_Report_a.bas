Operation =1
Option =0
Begin InputTables
    Name ="qry_sel_OT_Census_Report"
End
Begin OutputColumns
    Expression ="qry_sel_OT_Census_Report.Unit_Code"
    Expression ="qry_sel_OT_Census_Report.Plot_ID"
    Alias ="ParkPlot"
    Expression ="[Unit_Code] & [Plot_ID]"
    Expression ="qry_sel_OT_Census_Report.Quad"
    Expression ="qry_sel_OT_Census_Report.Tag_No"
    Expression ="qry_sel_OT_Census_Report.Visit_Year"
    Expression ="qry_sel_OT_Census_Report.Notes"
    Alias ="LofN"
    Expression ="Len([Notes])"
    Alias ="LofN2"
    Expression ="IIf([Notes] Is Null,0,Len([Notes]))"
End
Begin OrderBy
    Expression ="qry_sel_OT_Census_Report.Unit_Code"
    Flag =0
    Expression ="qry_sel_OT_Census_Report.Plot_ID"
    Flag =0
    Expression ="qry_sel_OT_Census_Report.Tag_No"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2b1b7ba014816244867063659f0f1f1f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_sel_OT_Census_Report.Tag_No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_sel_OT_Census_Report.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_sel_OT_Census_Report.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_sel_OT_Census_Report.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_sel_OT_Census_Report.Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4875"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ParkPlot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa4f94729ab108d47963342b75fa6a710
        End
    End
    Begin
        dbText "Name" ="qry_sel_OT_Census_Report.Quad"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LofN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x000000000000000080c4da0090605225
        End
    End
    Begin
        dbText "Name" ="LofN2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf01545fb7f057848a961d909abbf5040
        End
    End
End
Begin
    State =0
    Left =-159
    Top =130
    Right =1398
    Bottom =800
    Left =-1
    Top =-1
    Right =1386
    Bottom =344
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =228
        Bottom =267
        Top =0
        Name ="qry_sel_OT_Census_Report"
        Name =""
    End
End
