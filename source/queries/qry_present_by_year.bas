Operation =1
Option =0
Begin InputTables
    Name ="qry_sel_Plant_Year"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="qry_sel_Plant_Year.Unit_Code"
    Expression ="qry_sel_Plant_Year.Plot_ID"
    Expression ="tlu_NCPN_Plants.Master_Family"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Alias ="2006"
    Expression ="IIf([present_Year]=2006,\"X\",\" \")"
    Alias ="2007"
    Expression ="IIf([present_Year]=2007,\"X\",\" \")"
End
Begin Joins
    LeftTable ="qry_sel_Plant_Year"
    RightTable ="tlu_NCPN_Plants"
    Expression ="qry_sel_Plant_Year.Plant_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
End
Begin OrderBy
    Expression ="qry_sel_Plant_Year.Unit_Code"
    Flag =0
    Expression ="qry_sel_Plant_Year.Plot_ID"
    Flag =0
    Expression ="tlu_NCPN_Plants.Master_Family"
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
    0x6a1819cf520bcd4f82e4ab16b2e62a90
End
Begin
    Begin
        dbText "Name" ="2006"
        dbBinary "GUID" = Begin
            0xabe625522f900c418b2a04c1095761ab
        End
    End
    Begin
        dbText "Name" ="2007"
        dbBinary "GUID" = Begin
            0xbc17553da1dd194099a8a778487c8eaa
        End
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1002
    Bottom =327
    Left =-1
    Top =-1
    Right =973
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
        Name ="qry_sel_Plant_Year"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =332
        Bottom =120
        Top =5
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
