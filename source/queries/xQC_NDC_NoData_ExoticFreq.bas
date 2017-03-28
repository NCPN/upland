Operation =1
Option =0
Having ="(((tbl_Events.Start_Date) Is Not Null) AND ((Count(tbl_LP_Exotic_Freq.Exotic_ID)"
    ")=0))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Exotic_Freq"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Expression ="tbl_LP_Belt_Transect.Transect"
    Alias ="Exotic_Freq_Count"
    Expression ="Count(tbl_LP_Exotic_Freq.Exotic_ID)"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Belt_Transect"
    Expression ="tbl_Events.Event_ID = tbl_LP_Belt_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Exotic_Freq"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Exotic_Freq.Transect_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tbl_Events.Start_Date"
    Flag =0
    Expression ="tbl_LP_Belt_Transect.Transect"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Plot_ID"
    GroupLevel =0
    Expression ="tbl_Events.Start_Date"
    GroupLevel =0
    Expression ="tbl_LP_Belt_Transect.Transect"
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
            0x8bcc4a0d481759489fda18adb770fbe3
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x225b7db5513f40489ff1b220ed36e78a
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa61228be78678469cacb57697d467f4
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Belt_Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb9b632c880056e40a39103af4ddb3be8
        End
    End
    Begin
        dbText "Name" ="Exotic_Freq_Count"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa1ffb32245b78478abfe38d59fbc806
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
    Bottom =134
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
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_LP_Exotic_Freq"
        Name =""
    End
End
