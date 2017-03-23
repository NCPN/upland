Operation =1
Option =2
Where ="(((tbl_Events.Start_Date) Is Not Null) AND ((tbl_Locations.Vegetation_Type)=\"fo"
    "rest\" Or (tbl_Locations.Vegetation_Type)=\"woodland\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Fuels_1000"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Expression ="tbl_Fuels_1000.Transect"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Fuels_1000"
    Expression ="tbl_Events.Event_ID = tbl_Fuels_1000.Event_ID"
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
    Expression ="tbl_Fuels_1000.Transect"
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
            0x0cb0ffb1656a9f4c9c21c918d70b239b
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa4b5edab61de8144b972e90ee726879f
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3669b825b9429b499a3f7b8e0bb2fd44
        End
    End
    Begin
        dbText "Name" ="tbl_Fuels_1000.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b90c73f1f4836418e8800ef07471c41
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
    Bottom =116
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
        Name ="tbl_Fuels_1000"
        Name =""
    End
End
