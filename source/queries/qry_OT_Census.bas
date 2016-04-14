Operation =1
Option =0
Begin InputTables
    Name ="tbl_OT_Census"
End
Begin OutputColumns
    Expression ="tbl_OT_Census.Census_ID"
    Expression ="tbl_OT_Census.Event_ID"
    Expression ="tbl_OT_Census.Quad"
    Expression ="tbl_OT_Census.Tag_No"
    Expression ="tbl_OT_Census.Species"
    Expression ="tbl_OT_Census.DBH"
    Expression ="tbl_OT_Census.Crown_Health"
    Expression ="tbl_OT_Census.Crown_Class"
    Expression ="tbl_OT_Census.Notes"
    Expression ="tbl_OT_Census.DType"
End
Begin OrderBy
    Expression ="tbl_OT_Census.Quad"
    Flag =0
    Expression ="tbl_OT_Census.Tag_No"
    Flag =0
    Expression ="tbl_OT_Census.Species"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdc9c52bdbf0e2a4d9f14caeb98fac382
End
Begin
End
Begin
    State =0
    Left =47
    Top =43
    Right =987
    Bottom =367
    Left =-1
    Top =-1
    Right =925
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =168
        Bottom =109
        Top =4
        Name ="tbl_OT_Census"
        Name =""
    End
End
