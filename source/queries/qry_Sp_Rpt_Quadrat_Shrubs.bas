Operation =1
Option =2
Begin InputTables
    Name ="tlu_NCPN_Plants"
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Transect"
    Name ="tbl_Quadrat"
    Name ="tbl_Quadrat_Shrubs"
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
    LeftTable ="tlu_NCPN_Plants"
    RightTable ="tbl_Quadrat_Shrubs"
    Expression ="tlu_NCPN_Plants.Master_PLANT_Code = tbl_Quadrat_Shrubs.Plant_Code"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Transect"
    Expression ="tbl_Events.Event_ID = tbl_Quadrat_Transect.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Quadrat_Transect"
    RightTable ="tbl_Quadrat"
    Expression ="tbl_Quadrat_Transect.Transect_ID = tbl_Quadrat.Transect_ID"
    Flag =1
    LeftTable ="tbl_Quadrat"
    RightTable ="tbl_Quadrat_Shrubs"
    Expression ="tbl_Quadrat.Quadrat_ID = tbl_Quadrat_Shrubs.Quadrat_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tlu_NCPN_Plants.Master_Family"
    Flag =0
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Flag =0
    Expression ="Year([Start_Date])"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xaa337786da5b2a49aa7e2ec9d3696a72
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Year"
        dbBinary "GUID" = Begin
            0xbe3b2b2c16a8ef4994e181921ca7c327
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
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =708
        Top =6
        Right =842
        Bottom =120
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
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
        Name ="tbl_Quadrat_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =120
        Top =0
        Name ="tbl_Quadrat"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =670
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Shrubs"
        Name =""
    End
End
