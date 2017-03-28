Operation =1
Option =0
Where ="(((tbl_Locations.Vegetation_Type)=\"forest\" Or (tbl_Locations.Vegetation_Type)="
    "\"woodland\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tlu_Fuels_Transects"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Events.Start_Date"
    Expression ="tlu_Fuels_Transects.FuelsTransect"
End
Begin Joins
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
    Expression ="tlu_Fuels_Transects.FuelsTransect"
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
            0x2ff2e2019d940342873609c7ba1126d6
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6bcf45fed415df4fb9051693a3193658
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3c819d18bb3d9649a11786688a9af289
        End
    End
    Begin
        dbText "Name" ="tlu_Fuels_Transects.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb8242f8efc7f954fa6d73461d096a57f
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
        Name ="tlu_Fuels_Transects"
        Name =""
    End
End
