Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_SL_Transect"
    Name ="tbl_SLI_Gaps"
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
    LeftTable ="tbl_Events"
    RightTable ="tbl_SL_Transect"
    Expression ="tbl_Events.Event_ID = tbl_SL_Transect.Event_ID"
    Flag =1
    LeftTable ="tbl_SLI_Gaps"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_SLI_Gaps.Species = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =1
    LeftTable ="tbl_SL_Transect"
    RightTable ="tbl_SLI_Gaps"
    Expression ="tbl_SL_Transect.Transect_ID = tbl_SLI_Gaps.Transect_ID"
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
dbMemo "Filter" ="([qry_Sp_Rpt_SLI].[Utah_Species] Is Null OR [qry_Sp_Rpt_SLI].[Utah_Species]=\"\""
    ")"
dbBinary "GUID" = Begin
    0x2b952490b5cca84b8808ceebab0e92ea
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Family"
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
            0xaf736d8023d59d479000a1268b2cb70a
        End
    End
End
Begin
    State =0
    Left =76
    Top =72
    Right =1301
    Bottom =385
    Left =-1
    Top =-1
    Right =1201
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
        Name ="tbl_SL_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =120
        Top =0
        Name ="tbl_SLI_Gaps"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =670
        Bottom =120
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
