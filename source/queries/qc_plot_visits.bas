dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start_Date, tb"
    "l_Locations.Vegetation_Type\015\012FROM tbl_Locations INNER JOIN tbl_Events ON t"
    "bl_Locations.Location_ID = tbl_Events.Location_ID\015\012ORDER BY tbl_Locations."
    "Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
