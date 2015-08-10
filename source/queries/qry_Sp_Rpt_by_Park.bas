Operation =1
Option =0
Begin InputTables
    Name ="qry_Sp_Rpt_All"
End
Begin OutputColumns
    Expression ="qry_Sp_Rpt_All.Unit_Code"
    Expression ="qry_Sp_Rpt_All.Plot_ID"
    Expression ="qry_Sp_Rpt_All.Master_Family"
    Expression ="qry_Sp_Rpt_All.Utah_Species"
    Alias ="Visit_Year"
    Expression ="qry_Sp_Rpt_All.Year"
End
Begin OrderBy
    Expression ="qry_Sp_Rpt_All.Plot_ID"
    Flag =0
    Expression ="qry_Sp_Rpt_All.Master_Family"
    Flag =0
    Expression ="qry_Sp_Rpt_All.Utah_Species"
    Flag =0
    Expression ="qry_Sp_Rpt_All.Year"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd909755e3918684f970330a5fe90d71a
End
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbBinary "GUID" = Begin
            0xa21283488f1c4148a5f5d05639e52a87
        End
    End
End
Begin
    State =0
    Left =73
    Top =119
    Right =940
    Bottom =432
    Left =-1
    Top =-1
    Right =856
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =318
        Bottom =120
        Top =1
        Name ="qry_Sp_Rpt_All"
        Name =""
    End
End
