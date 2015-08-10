Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Fuels"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Fuels.Fuels_ID"
    Expression ="tbl_Fuels.Event_ID"
    Expression ="tbl_Fuels.[1HR_A]"
    Expression ="tbl_Fuels.[1HR_B]"
    Expression ="tbl_Fuels.[1HR_C]"
    Expression ="tbl_Fuels.[1HR_D]"
    Expression ="tbl_Fuels.[10HR_A]"
    Expression ="tbl_Fuels.[10HR_B]"
    Expression ="tbl_Fuels.[10HR_C]"
    Expression ="tbl_Fuels.[10HR_D]"
    Expression ="tbl_Fuels.[100HR_A]"
    Expression ="tbl_Fuels.[100HR_B]"
    Expression ="tbl_Fuels.[100HR_C]"
    Expression ="tbl_Fuels.[100HR_D]"
    Expression ="tbl_Fuels.[2Litter_A]"
    Expression ="tbl_Fuels.[4Litter_A]"
    Expression ="tbl_Fuels.[6Litter_A]"
    Expression ="tbl_Fuels.[8Litter_A]"
    Expression ="tbl_Fuels.[10Litter_A]"
    Expression ="tbl_Fuels.[12Litter_A]"
    Expression ="tbl_Fuels.[14Litter_A]"
    Expression ="tbl_Fuels.[2Litter_B]"
    Expression ="tbl_Fuels.[4Litter_B]"
    Expression ="tbl_Fuels.[6Litter_B]"
    Expression ="tbl_Fuels.[8Litter_B]"
    Expression ="tbl_Fuels.[10Litter_B]"
    Expression ="tbl_Fuels.[12Litter_B]"
    Expression ="tbl_Fuels.[14Litter_B]"
    Expression ="tbl_Fuels.[2Litter_C]"
    Expression ="tbl_Fuels.[4Litter_C]"
    Expression ="tbl_Fuels.[6Litter_C]"
    Expression ="tbl_Fuels.[8Litter_C]"
    Expression ="tbl_Fuels.[10Litter_C]"
    Expression ="tbl_Fuels.[12Litter_C]"
    Expression ="tbl_Fuels.[14Litter_C]"
    Expression ="tbl_Fuels.[2Litter_D]"
    Expression ="tbl_Fuels.[4Litter_D]"
    Expression ="tbl_Fuels.[6Litter_D]"
    Expression ="tbl_Fuels.[8Litter_D]"
    Expression ="tbl_Fuels.[10Litter_D]"
    Expression ="tbl_Fuels.[12Litter_D]"
    Expression ="tbl_Fuels.[14Litter_D]"
    Expression ="tbl_Fuels.[2Duff_A]"
    Expression ="tbl_Fuels.[4Duff_A]"
    Expression ="tbl_Fuels.[6Duff_A]"
    Expression ="tbl_Fuels.[8Duff_A]"
    Expression ="tbl_Fuels.[10Duff_A]"
    Expression ="tbl_Fuels.[12Duff_A]"
    Expression ="tbl_Fuels.[14Duff_A]"
    Expression ="tbl_Fuels.[2Duff_B]"
    Expression ="tbl_Fuels.[4Duff_B]"
    Expression ="tbl_Fuels.[6Duff_B]"
    Expression ="tbl_Fuels.[8Duff_B]"
    Expression ="tbl_Fuels.[10Duff_B]"
    Expression ="tbl_Fuels.[12Duff_B]"
    Expression ="tbl_Fuels.[14Duff_B]"
    Expression ="tbl_Fuels.[2Duff_C]"
    Expression ="tbl_Fuels.[4Duff_C]"
    Expression ="tbl_Fuels.[6Duff_C]"
    Expression ="tbl_Fuels.[8Duff_C]"
    Expression ="tbl_Fuels.[10Duff_C]"
    Expression ="tbl_Fuels.[12Duff_C]"
    Expression ="tbl_Fuels.[14Duff_C]"
    Expression ="tbl_Fuels.[2Duff_D]"
    Expression ="tbl_Fuels.[4Duff_D]"
    Expression ="tbl_Fuels.[6Duff_D]"
    Expression ="tbl_Fuels.[8Duff_D]"
    Expression ="tbl_Fuels.[10Duff_D]"
    Expression ="tbl_Fuels.[12Duff_D]"
    Expression ="tbl_Fuels.[14Duff_D]"
    Expression ="tbl_Locations.Bearing_A"
    Expression ="tbl_Locations.Bearing_B"
    Expression ="tbl_Locations.Bearing_C"
    Expression ="tbl_Locations.Bearing_D"
    Expression ="tbl_Locations.Slope_A"
    Expression ="tbl_Locations.Slope_B"
    Expression ="tbl_Locations.Slope_C"
    Expression ="tbl_Locations.Slope_D"
End
Begin Joins
    LeftTable ="tbl_Fuels"
    RightTable ="tbl_Events"
    Expression ="tbl_Fuels.Event_ID=tbl_Events.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2648e2f7c7c9c8439bcf0b7bb6b845ad
End
Begin
End
Begin
    State =0
    Left =61
    Top =43
    Right =1206
    Bottom =367
    Left =-1
    Top =-1
    Right =1126
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="tbl_Fuels"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =124
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =124
        Top =91
        Name ="tbl_Locations"
        Name =""
    End
End
