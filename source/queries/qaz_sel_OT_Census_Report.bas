Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_OT_Census"
    Name ="tlu_Crown_Health_Class"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_OT_Census.Tag_No"
    Expression ="tbl_OT_Census.Notes"
    Alias ="Visit_Year"
    Expression ="DatePart(\"yyyy\",[tbl_events].[Start_Date])"
    Alias ="ParkPLot"
    Expression ="[Unit_Code] & [Plot_Id]"
    Expression ="tbl_OT_Census.Species"
End
Begin Joins
    LeftTable ="tbl_OT_Census"
    RightTable ="tlu_Crown_Health_Class"
    Expression ="tbl_OT_Census.Crown_Health = tlu_Crown_Health_Class.Crown_Health_Class"
    Flag =2
    LeftTable ="tbl_OT_Census"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_OT_Census.Species = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_OT_Census"
    Expression ="tbl_Events.Event_ID = tbl_OT_Census.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="DatePart(\"yyyy\",[tbl_events].[Start_Date])"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7f8a0ac0aa327740af95947fd5e482e8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbInteger "RowHeight" ="810"
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa2e1794c64a4754689fe8cab1629d7db
        End
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Species"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="7500"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ParkPLot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x51f26708a44b954893027969f857d227
        End
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Tag_No"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =92
    Top =39
    Right =1191
    Bottom =732
    Left =-1
    Top =-1
    Right =1075
    Bottom =282
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =175
        Top =16
        Right =354
        Bottom =200
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =399
        Top =5
        Right =529
        Bottom =243
        Top =0
        Name ="tbl_OT_Census"
        Name =""
    End
    Begin
        Left =701
        Top =8
        Right =881
        Bottom =96
        Top =0
        Name ="tlu_Crown_Health_Class"
        Name =""
    End
    Begin
        Left =564
        Top =144
        Right =698
        Bottom =247
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
