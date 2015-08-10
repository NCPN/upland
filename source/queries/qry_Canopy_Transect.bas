Operation =1
Option =0
Begin InputTables
    Name ="tbl_Canopy_Transect"
End
Begin OutputColumns
    Expression ="tbl_Canopy_Transect.Transect_ID"
    Expression ="tbl_Canopy_Transect.Event_ID"
    Expression ="tbl_Canopy_Transect.Transect"
    Expression ="tbl_Canopy_Transect.Visit_Date"
    Expression ="tbl_Canopy_Transect.Observer"
    Expression ="tbl_Canopy_Transect.Recorder"
    Expression ="tbl_Canopy_Transect.Species"
End
Begin OrderBy
    Expression ="tbl_Canopy_Transect.Transect"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfed2d1cdb4a65b4faf927a1d95034d43
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
        Top =0
        Name ="tbl_Canopy_Transect"
        Name =""
    End
End
