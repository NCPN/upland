Operation =1
Option =0
Where ="(((tbl_OT_Census.Crown_Health)<>6))"
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
    Expression ="tbl_OT_Census.Quad"
    Expression ="tbl_OT_Census.Tag_No"
    Expression ="tbl_OT_Census.Species"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Expression ="tbl_OT_Census.DBH"
    Expression ="tbl_OT_Census.Crown_Class"
    Expression ="tlu_Crown_Health_Class.Class_Description"
    Expression ="tbl_OT_Census.Notes"
    Alias ="Visit_Year"
    Expression ="DatePart(\"yyyy\",tbl_events.Start_Date)"
    Expression ="tbl_OT_Census.DType"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_OT_Census"
    Expression ="tbl_Events.Event_ID = tbl_OT_Census.Event_ID"
    Flag =2
    LeftTable ="tbl_OT_Census"
    RightTable ="tlu_Crown_Health_Class"
    Expression ="tbl_OT_Census.Crown_Health = tlu_Crown_Health_Class.Crown_Health_Class"
    Flag =2
    LeftTable ="tbl_OT_Census"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_OT_Census.Species = tlu_NCPN_Plants.Master_PLANT_Code"
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
    Expression ="tbl_OT_Census.Quad"
    Flag =0
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
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x371e7fa38d126a4aab7f99d86b2b93f5
        End
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
        dbText "Name" ="tbl_OT_Census.Quad"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Tag_No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Crown_Health_Class.Class_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.DType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =169
    Right =923
    Bottom =670
    Left =-1
    Top =-1
    Right =1229
    Bottom =110
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
        Left =172
        Top =6
        Right =268
        Bottom =109
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =436
        Bottom =109
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
        Left =501
        Top =7
        Right =635
        Bottom =110
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
