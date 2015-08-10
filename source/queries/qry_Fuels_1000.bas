Operation =1
Option =0
Begin InputTables
    Name ="tbl_Fuels_1000"
End
Begin OutputColumns
    Expression ="tbl_Fuels_1000.Fuels_1000_ID"
    Expression ="tbl_Fuels_1000.Event_ID"
    Expression ="tbl_Fuels_1000.Transect"
    Expression ="tbl_Fuels_1000.Diameter"
    Expression ="tbl_Fuels_1000.Decay_Class"
End
Begin OrderBy
    Expression ="tbl_Fuels_1000.Transect"
    Flag =0
    Expression ="tbl_Fuels_1000.Diameter"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2a7a0bab6d9a4c409fcc9706dfd0b86b
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
        Right =134
        Bottom =120
        Top =0
        Name ="tbl_Fuels_1000"
        Name =""
    End
End
