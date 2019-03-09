﻿dbMemo "SQL" ="SELECT L.Unit_Code, L.Plot_ID, L.Vegetation_Type, E.Start_Date, LPBT.Transect, P"
    "LANTS.LU_Code AS Species\015\012FROM ((tbl_Locations AS L INNER JOIN (tbl_Events"
    " AS E INNER JOIN tbl_LP_Belt_Transect AS LPBT ON E.Event_ID = LPBT.Event_ID) ON "
    "L.Location_ID = E.Location_ID) INNER JOIN tbl_LP_Exotic_Freq AS LPEF ON LPBT.Tra"
    "nsect_ID = LPEF.Transect_ID) LEFT JOIN tlu_NCPN_Plants AS PLANTS ON LPEF.Species"
    " = PLANTS.Master_PLANT_Code\015\012WHERE (((L.Vegetation_Type<>\"oak scrub\") AN"
    "D (LPEF.M0=False) AND (LPEF.M5=False) AND (LPEF.M10=False) AND (LPEF.M15=False) "
    "AND (LPEF.M20=False) AND (LPEF.M25=False) AND (LPEF.M30=False) AND (LPEF.M35=Fal"
    "se) AND (LPEF.M40=False) AND (LPEF.M45=False)) OR ((L.Vegetation_Type=\"oak scru"
    "b\") AND (LPEF.Oak0=False) AND (LPEF.Oak2=False) AND (LPEF.Oak4=False) AND (LPEF"
    ".Oak6=False) AND (LPEF.Oak8=False) AND (LPEF.Oak10=False) AND (LPEF.Oak12=False)"
    " AND (LPEF.Oak14=False) AND (LPEF.Oak16=False) AND (LPEF.Oak18=False)) OR ((LPEF"
    ".Species Is Null) Or (LPEF.Species=\"\") Or (LPEF.Species=\" \")))\015\012AND L."
    "Unit_Code = 'ARCH'\015\012AND E.Start_Date = #5/2/2018#\015\012AND L.Plot_ID = 1"
    "\015\012ORDER BY L.Unit_Code, L.Plot_ID, E.Start_Date, LPBT.Transect, PLANTS.LU_"
    "Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
