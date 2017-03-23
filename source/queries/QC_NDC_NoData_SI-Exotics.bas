Operation =1
Option =0
Having ="(((tbl_Events.Start_Date) Is Not Null) AND ((Count(tbl_Dist_Exotic.Exotic_ID))=0"
    "))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Site_Impact"
    Name ="tbl_Dist_Exotic"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Alias ="SI_Exotic_Count"
    Expression ="Count(tbl_Dist_Exotic.Exotic_ID)"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Site_Impact"
    Expression ="tbl_Events.Event_ID = tbl_Site_Impact.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Site_Impact"
    RightTable ="tbl_Dist_Exotic"
    Expression ="tbl_Site_Impact.Impact_ID = tbl_Dist_Exotic.Impact_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tbl_Events.Start_Date"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Plot_ID"
    GroupLevel =0
    Expression ="tbl_Events.Start_Date"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xacc27a755c0c0740adeb9d3d1908a008
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcdd840b5b966fe4eaa3e8b40e7bfb565
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ce68970dc5d4048891c92b3e0d31799
        End
    End
    Begin
        dbText "Name" ="SI_Exotic_Count"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x80e8050208bf2400f4be240080f24203
        End
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1005
    Bottom =533
    Left =-1
    Top =-1
    Right =989
    Bottom =112
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tbl_Site_Impact"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_Dist_Exotic"
        Name =""
    End
End
