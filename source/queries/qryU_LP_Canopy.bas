dbMemo "SQL" ="SELECT tlu_LP_Soil_Surface.Surface_Code AS Master_Plant_Code,  tlu_LP_Soil_Surfa"
    "ce.Surface_Code AS LU_Code,  (tlu_LP_Soil_Surface.Surface_Code & \" - \" & tlu_L"
    "P_Soil_Surface.Surface_Description) AS Utah_Species, 1 as Sort_Seq\015\012FROM t"
    "lu_LP_Soil_Surface WHERE  tlu_LP_Soil_Surface.LC_Code = 1\015\012UNION SELECT tl"
    "u_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.LU_Code,  tlu_NCPN_Plants.Utah_"
    "Species, 2 as Sort_Seq\015\012FROM tlu_NCPN_Plants WHERE tlu_NCPN_Plants.Utah_PL"
    "ANT_Code is not null\015\012UNION SELECT tbl_Unknown_Species.Unknown_Code AS Mas"
    "ter_Plant_Code,  tbl_Unknown_Species.Unknown_Code AS LU_Code, (tbl_Unknown_Speci"
    "es.Unknown_Code & \"   \" & tbl_Unknown_Species.Plant_Description) AS Utah_Speci"
    "es, 2 AS Sort_Seq\015\012FROM tbl_Unknown_Species\015\012ORDER BY Sort_Seq, LU_C"
    "ode;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd79084802d22d74f8cfc464f7aa34376
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Utah_Species"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9e0b6af8c83bb240ae123d6b5ede9676
        End
    End
    Begin
        dbText "Name" ="Master_Plant_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x24012dfa53733a408f214b9b6f21abeb
        End
    End
    Begin
        dbText "Name" ="Sort_Seq"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67120e199fb13942a940862d8a20e80b
        End
    End
    Begin
        dbText "Name" ="LU_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9d8cb3f338f049459e652272c7f29ae6
        End
    End
End
