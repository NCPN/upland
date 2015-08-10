Operation =1
Option =0
Begin InputTables
    Name ="tbl_LP_Belt_Transect"
End
Begin OutputColumns
    Expression ="tbl_LP_Belt_Transect.Transect_ID"
    Expression ="tbl_LP_Belt_Transect.Event_ID"
    Expression ="tbl_LP_Belt_Transect.Transect"
    Expression ="tbl_LP_Belt_Transect.Visit_Date"
    Expression ="tbl_LP_Belt_Transect.Observer"
    Expression ="tbl_LP_Belt_Transect.Recorder"
End
Begin OrderBy
    Expression ="tbl_LP_Belt_Transect.Transect"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2c8a13982df735489fdba71cadd28ea8
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
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =2
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
End
