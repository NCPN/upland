﻿Operation =1
Option =2
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Shrub"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tlu_NCPN_Plants.Master_Family"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Alias ="Year"
    Expression ="Year([Start_Date])"
End
Begin Joins
    LeftTable ="tbl_LP_Shrub"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_LP_Shrub.Species = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =1
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Shrub"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Shrub.Transect_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Belt_Transect"
    Expression ="tbl_Events.Event_ID = tbl_LP_Belt_Transect.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "Filter" ="([qry_Sp_Rpt_BT_Shrubs].[Utah_Species] Is Null OR [qry_Sp_Rpt_BT_Shrubs].[Utah_S"
    "pecies]=\"\")"
dbBinary "GUID" = Begin
    0xc66c9a3a05a04248980101fb238b3d6a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Family"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e582aaed247d94e9979568ca14cd2fa
        End
    End
End
Begin
    State =0
    Left =47
    Top =43
    Right =1267
    Bottom =356
    Left =-1
    Top =-1
    Right =1196
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =120
        Top =0
        Name ="tbl_LP_Shrub"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =705
        Bottom =120
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
