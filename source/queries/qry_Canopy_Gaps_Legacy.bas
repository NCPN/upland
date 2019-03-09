Operation =1
Option =0
Where ="(((tbl_Events.Start_Date)<#1/1/2011#))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Canopy_Transect"
    Name ="tbl_Canopy_Gaps"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Expression ="tbl_Canopy_Transect.Transect"
    Expression ="tbl_Canopy_Gaps.Gap_ID"
    Expression ="tbl_Canopy_Gaps.Class"
    Expression ="tbl_Canopy_Gaps.Start"
    Expression ="tbl_Canopy_Gaps.Gap_End"
End
Begin Joins
    LeftTable ="tbl_Canopy_Transect"
    RightTable ="tbl_Canopy_Gaps"
    Expression ="tbl_Canopy_Transect.Transect_ID = tbl_Canopy_Gaps.Transect_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Canopy_Transect"
    Expression ="tbl_Events.Event_ID = tbl_Canopy_Transect.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tbl_Events.Start_Date"
    Flag =0
    Expression ="tbl_Canopy_Transect.Transect"
    Flag =0
    Expression ="tbl_Canopy_Gaps.Gap_ID"
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
            0x1563cc25baa3aa4ba2ac1f5f775c6e92
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x47be9b96602c0e4bac7b8c76d002177b
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xecd63fa4794656409db59ea7d7e2aa1b
        End
    End
    Begin
        dbText "Name" ="tbl_Canopy_Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0423fe554de3234a9a16df84d7ae3ccc
        End
    End
    Begin
        dbText "Name" ="tbl_Canopy_Gaps.Gap_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x41af6a91703be749a8ece1d50225f798
        End
    End
    Begin
        dbText "Name" ="tbl_Canopy_Gaps.Class"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x642766f6299bc54692c979ebf5a9e8dc
        End
    End
    Begin
        dbText "Name" ="tbl_Canopy_Gaps.Start"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x770c2cf354ad7849949d4f09d82ff95e
        End
    End
    Begin
        dbText "Name" ="tbl_Canopy_Gaps.Gap_End"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8f5dcf681c39654cb65a2dcc7ce15b37
        End
    End
End
Begin
    State =0
    Left =43
    Top =38
    Right =923
    Bottom =475
    Left =-1
    Top =-1
    Right =848
    Bottom =140
    Left =0
    Top =0
    ColumnsShown =539
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
        Name ="tbl_Canopy_Transect"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_Canopy_Gaps"
        Name =""
    End
End
