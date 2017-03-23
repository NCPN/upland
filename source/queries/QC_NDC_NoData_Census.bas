Operation =1
Option =0
Having ="(((tbl_Events.Start_Date) Is Not Null) AND ((Count(tbl_OT_Census.Census_ID))=0))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_OT_Census"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Alias ="CountOfCensus_ID"
    Expression ="Count(tbl_OT_Census.Census_ID)"
End
Begin Joins
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
            0x5d58bd36933e164abde3d8841eb77890
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf9824c3656d1c54682991ff3e04f8dc3
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4156efbf1bc6874b86fce79b94f6667e
        End
    End
    Begin
        dbText "Name" ="CountOfCensus_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf22c816058cfd49b81a1588ab6e0bf9
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
    Bottom =133
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
        Name ="tbl_OT_Census"
        Name =""
    End
End
