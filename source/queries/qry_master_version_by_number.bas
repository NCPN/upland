Operation =1
Option =0
Begin InputTables
    Name ="tbl_master_version"
End
Begin OutputColumns
    Expression ="tbl_master_version.*"
End
Begin OrderBy
    Expression ="tbl_master_version.version_key_number"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7d100e01cee0d2468f01c9fbd8d8dc30
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =40
    Right =895
    Bottom =353
    Left =-1
    Top =-1
    Right =853
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =193
        Bottom =120
        Top =0
        Name ="tbl_master_version"
        Name =""
    End
End
