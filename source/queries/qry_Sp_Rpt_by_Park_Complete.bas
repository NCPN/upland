dbMemo "SQL" ="SELECT qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Year, qry_Sp_Rpt_All.Plot_ID, qr"
    "y_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All.Utah_Species, (qry_Sp_Rpt_All.Utah_Sp"
    "ecies+\"-\"+CStr(qry_Sp_Rpt_All.Year)) AS SpeciesYear, (qry_Sp_Rpt_All.Unit_Code"
    "+\"-\"+CStr(qry_Sp_Rpt_All.Utah_Species)) AS ParkSpecies\015\012FROM qry_Sp_Rpt_"
    "All\015\012ORDER BY qry_Sp_Rpt_All.Plot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp"
    "_Rpt_All.Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0x4421a2a8dab2aa4e84d97c0e6aa83aa2
End
Begin
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
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesYear"
        dbInteger "ColumnWidth" ="2796"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkSpecies"
        dbInteger "ColumnWidth" ="2796"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
