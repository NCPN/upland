Operation =1
Option =2
Where ="(((tbl_Quadrat_Species.Plant_Code) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Transect"
    Name ="tbl_Quadrat"
    Name ="tbl_Quadrat_Species"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Present_Year"
    Expression ="Year(tbl_Quadrat_Transect.visit_Date)"
    Expression ="tbl_Quadrat_Species.Plant_Code"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Transect"
    Expression ="tbl_Events.Event_ID = tbl_Quadrat_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Quadrat_Transect"
    RightTable ="tbl_Quadrat"
    Expression ="tbl_Quadrat_Transect.Transect_ID = tbl_Quadrat.Transect_ID"
    Flag =2
    LeftTable ="tbl_Quadrat"
    RightTable ="tbl_Quadrat_Species"
    Expression ="tbl_Quadrat.Quadrat_ID = tbl_Quadrat_Species.Quadrat_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="Year(tbl_Quadrat_Transect.visit_Date)"
    Flag =0
    Expression ="tbl_Quadrat_Species.Plant_Code"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x565ecbd5cdd3c0419a27e5c6a7e894d8
End
Begin
    Begin
        dbText "Name" ="Present_Year"
        dbBinary "GUID" = Begin
            0xc80d16b839c94443ad9a0d4de4bf8a00
        End
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1002
    Bottom =327
    Left =-1
    Top =-1
    Right =977
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =1
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
        Left =308
        Top =6
        Right =485
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Transect"
        Name =""
    End
    Begin
        Left =517
        Top =12
        Right =613
        Bottom =126
        Top =4
        Name ="tbl_Quadrat"
        Name =""
    End
    Begin
        Left =694
        Top =6
        Right =790
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Species"
        Name =""
    End
End
