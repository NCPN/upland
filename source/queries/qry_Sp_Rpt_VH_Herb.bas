﻿dbMemo "SQL" ="SELECT DISTINCT L.Unit_Code, L.Plot_ID, NP.Master_Family, NP.Utah_Species, Year("
    "E.Start_Date) AS [Year]\015\012FROM (((tbl_Locations AS L INNER JOIN tbl_Events "
    "AS E ON L.Location_ID = E.Location_ID) INNER JOIN tbl_VH_Transect AS VHT ON E.Ev"
    "ent_ID = VHT.Event_ID) INNER JOIN tbl_VH_Intercept AS VHI ON VHT.Transect_ID = V"
    "HI.Transect_ID) INNER JOIN tlu_NCPN_Plants AS NP ON VHI.Herb = NP.Master_PLANT_C"
    "ode;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x02ca7c33c140684e9f920e6a8fc8fb59
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="L.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="L.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NP.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NP.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xea17b4d1e113fc4dbcc9c81052fefd16
        End
    End
End
