Operation =1
Option =0
Begin InputTables
    Name ="qry_Park_Species"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="qry_Park_Species.Unit_Code"
    Expression ="qry_Park_Species.Plant_Code"
    Expression ="tlu_NCPN_Plants.*"
End
Begin Joins
    LeftTable ="qry_Park_Species"
    RightTable ="tlu_NCPN_Plants"
    Expression ="qry_Park_Species.Plant_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9091131fff3da64cb9fc70144fbb36e2
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
        Bottom =90
        Top =0
        Name ="qry_Park_Species"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
