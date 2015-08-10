Operation =1
Option =0
Begin InputTables
    Name ="tbl_SL_Transect"
End
Begin OutputColumns
    Expression ="tbl_SL_Transect.Transect_ID"
    Expression ="tbl_SL_Transect.Event_ID"
    Expression ="tbl_SL_Transect.Transect"
    Expression ="tbl_SL_Transect.Visit_Date"
    Expression ="tbl_SL_Transect.Observer"
    Expression ="tbl_SL_Transect.Recorder"
End
Begin OrderBy
    Expression ="tbl_SL_Transect.Transect"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x69b42bcb3ee26949a1c4f832717ae7cc
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
        Right =173
        Bottom =120
        Top =0
        Name ="tbl_SL_Transect"
        Name =""
    End
End
