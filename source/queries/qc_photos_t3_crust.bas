dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, tbl_Events.Start_Date, tb"
    "l_Locations.Vegetation_Type, tbl_Photos.Transect, tbl_Photos.Location\015\012FRO"
    "M tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Photos ON tbl_Events.Event"
    "_ID = tbl_Photos.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID"
    "\015\012WHERE (((tbl_Photos.Transect)=\"T3 - 40m - crust\") \015\012AND ((tbl_Ph"
    "otos.Location)=40))\015\012ORDER BY tbl_Locations.Unit_Code, tbl_Locations.Plot_"
    "ID, tbl_Events.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
