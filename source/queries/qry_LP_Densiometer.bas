Operation =1
Option =0
Begin InputTables
    Name ="tbl_LP_Densiometer"
    Name ="tlu_Densiometer_Point"
End
Begin OutputColumns
    Expression ="tbl_LP_Densiometer.SD_ID"
    Expression ="tbl_LP_Densiometer.Transect_ID"
    Expression ="tbl_LP_Densiometer.Point"
    Expression ="tbl_LP_Densiometer.Total1"
    Expression ="tbl_LP_Densiometer.Total2"
    Expression ="tbl_LP_Densiometer.Total3"
    Expression ="tbl_LP_Densiometer.Total4"
    Expression ="tlu_Densiometer_Point.Sort_Seq"
End
Begin Joins
    LeftTable ="tbl_LP_Densiometer"
    RightTable ="tlu_Densiometer_Point"
    Expression ="tbl_LP_Densiometer.Point=tlu_Densiometer_Point.Point"
    Flag =2
End
Begin OrderBy
    Expression ="tlu_Densiometer_Point.Sort_Seq"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8b9241bb49d970409c74ddf14e6dfc97
End
Begin
End
Begin
    State =0
    Left =47
    Top =69
    Right =987
    Bottom =382
    Left =-1
    Top =-1
    Right =933
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =206
        Bottom =120
        Top =0
        Name ="tbl_LP_Densiometer"
        Name =""
    End
    Begin
        Left =244
        Top =6
        Right =412
        Bottom =90
        Top =0
        Name ="tlu_Densiometer_Point"
        Name =""
    End
End
