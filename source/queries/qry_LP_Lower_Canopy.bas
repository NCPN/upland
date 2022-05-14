Operation =1
Option =0
Begin InputTables
    Name ="tbl_LP_Lower_Canopy"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tbl_LP_Lower_Canopy.LC_ID"
    Alias ="Expr2"
    Expression ="tbl_LP_Lower_Canopy.Intercept_ID"
    Alias ="Expr3"
    Expression ="tbl_LP_Lower_Canopy.Sequence"
    Alias ="Expr4"
    Expression ="tbl_LP_Lower_Canopy.Species"
    Alias ="Expr5"
    Expression ="tbl_LP_Lower_Canopy.Alive"
End
Begin OrderBy
    Expression ="tbl_LP_Lower_Canopy.Sequence"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x48bfea5f5077f24496bd9a8442b8c243
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbBinary "GUID" = Begin
            0xa626eb2211cf8047a61f353bd96c489d
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbBinary "GUID" = Begin
            0x430af5227b2ddd4eb7c25ab73074a5ab
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbBinary "GUID" = Begin
            0x19fa8ad3b90f9545b17cb81d6085776b
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbBinary "GUID" = Begin
            0x7b6a26cd4a24074f84f8192d7b314828
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbBinary "GUID" = Begin
            0xa19c51bdcdcfd547812e51a7fda38339
        End
    End
End
Begin
    State =0
    Left =47
    Top =43
    Right =1002
    Bottom =356
    Left =-1
    Top =-1
    Right =931
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tbl_LP_Lower_Canopy"
        Name =""
    End
End
