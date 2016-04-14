Operation =1
Option =0
Begin InputTables
    Name ="tbl_Quadrat_Transect"
End
Begin OutputColumns
    Expression ="tbl_Quadrat_Transect.Transect_ID"
    Expression ="tbl_Quadrat_Transect.Event_ID"
    Expression ="tbl_Quadrat_Transect.Transect"
    Expression ="tbl_Quadrat_Transect.Visit_Date"
End
Begin OrderBy
    Expression ="tbl_Quadrat_Transect.Transect"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0b0f1349626e8b44a95e4c02a43aeb49
End
Begin
End
Begin
    State =0
    Left =18
    Top =14
    Right =1002
    Bottom =327
    Left =-1
    Top =-1
    Right =977
    Bottom =149
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =239
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Transect"
        Name =""
    End
End
