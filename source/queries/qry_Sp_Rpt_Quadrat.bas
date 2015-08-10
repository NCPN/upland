Operation =1
Option =2
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Transect"
    Name ="tbl_Quadrat"
    Name ="tbl_Quadrat_Species"
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
    LeftTable ="tbl_Quadrat_Species"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Quadrat_Species.Plant_Code = tlu_NCPN_Plants.Master_PLANT_Code"
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
    RightTable ="tbl_Quadrat_Species"
    Expression ="tbl_Quadrat.Quadrat_ID = tbl_Quadrat_Species.Quadrat_ID"
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
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x82d71889a5572f42a699e4f5a8e699d3
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc1b77001a3704d4aa2e4f71ab824da46
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ee3790763ef2240888bf02dc6cb2590
        End
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4be07dd1ee043c4991bf08978a876248
        End
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x080857f9c560434abb5962db9b57f270
        End
    End
End
Begin
    State =0
    Left =50
    Top =369
    Right =1135
    Bottom =755
    Left =-1
    Top =-1
    Right =1078
    Bottom =170
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =11
        Top =12
        Right =155
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =185
        Top =4
        Right =329
        Bottom =148
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =352
        Top =5
        Right =496
        Bottom =119
        Top =0
        Name ="tbl_Quadrat_Transect"
        Name =""
    End
    Begin
        Left =533
        Top =10
        Right =677
        Bottom =154
        Top =0
        Name ="tbl_Quadrat"
        Name =""
    End
    Begin
        Left =695
        Top =14
        Right =839
        Bottom =158
        Top =0
        Name ="tbl_Quadrat_Species"
        Name =""
    End
    Begin
        Left =867
        Top =17
        Right =1011
        Bottom =161
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
