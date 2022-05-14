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
dbMemo "Filter" ="((([qry_Sp_Rpt_by_Park].[Unit_Code]=\"ARCH\"))) AND ([qry_Sp_Rpt_by_Park].[Plot_"
    "ID] In (85,90))"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa21283488f1c4148a5f5d05639e52a87
        End
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =151
    Top =111
    Right =1462
    Bottom =809
    Left =-1
    Top =-1
    Right =1287
    Bottom =106
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =318
        Bottom =120
        Top =0
        Name ="qry_Sp_Rpt_All"
        Name =""
    End
End
