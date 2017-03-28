dbMemo "SQL" ="SELECT DISTINCT tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start"
    "_Date, tbl_Fuels_1000.Transect\015\012FROM tbl_Locations INNER JOIN (tbl_Events "
    "INNER JOIN tbl_Fuels_1000 ON tbl_Events.Event_ID = tbl_Fuels_1000.Event_ID) ON t"
    "bl_Locations.Location_ID = tbl_Events.Location_ID\015\012WHERE (((tbl_Events.Sta"
    "rt_Date) Is Not Null) \015\012AND ((tbl_Locations.Vegetation_Type)=\"forest\" Or"
    " (tbl_Locations.Vegetation_Type)=\"woodland\"))\015\012AND tbl_Locations.Unit_Co"
    "de = 'CEBR' \015\012AND tbl_Locations.Plot_ID = 133\015\012ORDER BY tbl_Location"
    "s.Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start_Date, tbl_Fuels_1000.Transe"
    "ct;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
