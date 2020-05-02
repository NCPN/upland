Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_VH_Transect"
    Name ="tbl_VH_Intercept"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Expression ="tbl_VH_Transect.Transect_ID"
    Expression ="tbl_VH_Transect.Event_ID"
    Expression ="tbl_VH_Transect.Transect"
    Expression ="tbl_VH_Transect.Visit_Date"
    Expression ="tbl_VH_Transect.Observer"
    Expression ="tbl_VH_Transect.Recorder"
    Expression ="tbl_VH_Intercept.Intercept_ID"
    Expression ="tbl_VH_Intercept.Transect_ID"
    Expression ="tbl_VH_Intercept.Point"
    Expression ="tbl_VH_Intercept.Wood"
    Expression ="tbl_VH_Intercept.WHeight"
    Expression ="tbl_VH_Intercept.Herb"
    Expression ="tbl_VH_Intercept.HHeight"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_VH_Transect"
    Expression ="tbl_Events.Event_ID = tbl_VH_Transect.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_VH_Transect"
    RightTable ="tbl_VH_Intercept"
    Expression ="tbl_VH_Transect.Transect_ID = tbl_VH_Intercept.Transect_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tbl_Events.Start_Date"
    Flag =0
    Expression ="tbl_VH_Transect.Transect"
    Flag =0
    Expression ="tbl_VH_Intercept.Point"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa03561aa7fc5ef478324c74c833a8204
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_VH_Intercept.HHeight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Visit_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.WHeight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Recorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Transect.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Intercept_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Point"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Wood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_VH_Intercept.Herb"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =53
    Right =982
    Bottom =504
    Left =-1
    Top =-1
    Right =930
    Bottom =196
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
        Name ="tbl_VH_Transect"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_VH_Intercept"
        Name =""
    End
End
