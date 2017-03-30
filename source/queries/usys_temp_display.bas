dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start_Date, Co"
    "unt(tbl_OT_Census.Census_ID) AS CountOfCensus_ID\015\012FROM tbl_Locations LEFT "
    "JOIN (tbl_Events LEFT JOIN tbl_OT_Census ON tbl_Events.Event_ID = tbl_OT_Census."
    "Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID\015\012GROUP BY "
    "tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start_Date\015\012HAV"
    "ING (((tbl_Events.Start_Date) Is Not Null) AND ((Count(tbl_OT_Census.Census_ID))"
    "=0))\015\012AND tbl_Locations.Unit_Code = 'CEBR' \015\012AND tbl_Locations.Plot_"
    "ID = 133\015\012ORDER BY tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tbl_Eve"
    "nts.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
