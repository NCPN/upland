dbMemo "SQL" ="SELECT qry_Sp_Rpt_All_AddField.Unit_Code, qry_Sp_Rpt_All_AddField.Plot_ID, qry_S"
    "p_Rpt_All_AddField.Master_Family, qry_Sp_Rpt_All_AddField.Utah_Species, qry_Sp_R"
    "pt_All_AddField.Year\015\012FROM tbl_Revisit_List INNER JOIN qry_Sp_Rpt_All_AddF"
    "ield ON tbl_Revisit_List.ParkPlot = qry_Sp_Rpt_All_AddField.ParkPlot;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x12fad52bc758e944a9005bec8e4e15a0
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_Sp_Rpt_All_AddField.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All_AddField.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All_AddField.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All_AddField.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All_AddField.Year"
        dbLong "AggregateType" ="-1"
    End
End
