Operation =1
Option =0
Where ="(((tlu_NCPN_Plants.Utah_PLANT_Code) Is Not Null) AND ((tlu_NCPN_Plants.Lifeform)"
    " Is Null))"
Begin InputTables
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tlu_NCPN_Plants.UT_Family"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Expression ="tlu_NCPN_Plants.Utah_PLANT_Code"
    Expression ="tlu_NCPN_Plants.Master_PLANT_Code"
    Expression ="tlu_NCPN_Plants.Master_Species"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Expression ="tlu_NCPN_Plants.Unique_Species"
End
Begin OrderBy
    Expression ="tlu_NCPN_Plants.UT_Family"
    Flag =0
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x709c469139023d4f88212016cab8b9a8
End
Begin
End
Begin
    State =0
    Left =47
    Top =43
    Right =1269
    Bottom =356
    Left =-1
    Top =-1
    Right =1215
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =203
        Bottom =120
        Top =36
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
