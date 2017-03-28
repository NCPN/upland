Operation =1
Option =0
Having ="(((tbl_Events.Start_Date) Is Not Null) AND ((Count(tbl_LP_Shrub.Shrub_ID))=0) AN"
    "D ((tbl_Locations.Vegetation_Type)<>\"oak scrub\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Shrub"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Expression ="tbl_LP_Belt_Transect.Transect"
    Alias ="Shrub_Count"
    Expression ="Count(tbl_LP_Shrub.Shrub_ID)"
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
    RightTable ="tbl_LP_Shrub"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Shrub.Transect_ID"
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
    Expression ="tbl_Locations.Vegetation_Type"
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
            0x5028ac28450aec4ea3b4f36c726f5601
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ba91b10da05cf41a5cf57ca32886759
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc940449b23b7df4eac502795bd7c5541
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Belt_Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x72c0a69def5a1e489b6f3d8f369f1ae6
        End
    End
    Begin
        dbText "Name" ="Shrub_Count"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa334677a5386641afe5402b0c14b9d5
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
    Bottom =135
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
        Name ="tbl_LP_Shrub"
        Name =""
    End
End
