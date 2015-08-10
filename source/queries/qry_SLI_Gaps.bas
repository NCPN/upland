Operation =1
Option =0
Begin InputTables
    Name ="tbl_SLI_Gaps"
End
Begin OutputColumns
    Expression ="tbl_SLI_Gaps.SLI_ID"
    Expression ="tbl_SLI_Gaps.Transect_ID"
    Expression ="tbl_SLI_Gaps.Species"
    Expression ="tbl_SLI_Gaps.Shrub_Start"
    Expression ="tbl_SLI_Gaps.Shrub_End"
End
Begin OrderBy
    Expression ="tbl_SLI_Gaps.Shrub_Start"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf16e4ed191e22c42903db870794d992e
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
        Name ="tbl_SLI_Gaps"
        Name =""
    End
End
