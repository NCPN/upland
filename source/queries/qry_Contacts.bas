Operation =1
Option =0
Where ="(((tlu_Contacts.Active)=1))"
Begin InputTables
    Name ="tlu_Contacts"
End
Begin OutputColumns
    Expression ="tlu_Contacts.Contact_ID"
    Expression ="tlu_Contacts.Last_Name"
End
Begin OrderBy
    Expression ="tlu_Contacts.Last_Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x41bc97097d339447b1061cd31a29c371
End
Begin
End
Begin
    State =0
    Left =61
    Top =43
    Right =1400
    Bottom =367
    Left =-1
    Top =-1
    Right =1324
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="tlu_Contacts"
        Name =""
    End
End
