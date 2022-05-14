dbMemo "SQL" ="SELECT DISTINCT tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tlu_NCPN_Plants."
    "Master_Family, tlu_NCPN_Plants.Utah_Species, Year([Start_Date]) AS [Year]\015\012"
    "FROM tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_LP_Transect INNER JOIN"
    " (tbl_LP_Intercept INNER JOIN tlu_NCPN_Plants ON tbl_LP_Intercept.Top = tlu_NCPN"
    "_Plants.Master_PLANT_Code) ON tbl_LP_Transect.Transect_ID = tbl_LP_Intercept.Tra"
    "nsect_ID) ON tbl_Events.Event_ID = tbl_LP_Transect.Event_ID) ON tbl_Locations.Lo"
    "cation_ID = tbl_Events.Location_ID\015\012ORDER BY tbl_Locations.Unit_Code, tbl_"
    "Locations.Plot_ID, tlu_NCPN_Plants.Master_Family, tlu_NCPN_Plants.Utah_Species, "
    "Year([Start_Date]);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x29d0e58963f1b54682af44ddcfd75e75
End
dbMemo "Filter" ="([qry_Sp_Rpt_PI_TC].[Utah_Species] Is Null OR [qry_Sp_Rpt_PI_TC].[Utah_Species]="
    "\"\")"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb15460e311978f48a8f0f68570bd33ff
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
End
