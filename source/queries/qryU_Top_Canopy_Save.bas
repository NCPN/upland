dbMemo "SQL" ="SELECT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Utah_PLANT_Code,  tlu_"
    "NCPN_Plants.Utah_Species, 2 as Sort_Seq\015\012FROM tlu_NCPN_Plants WHERE tlu_NC"
    "PN_Plants.Utah_PLANT_Code is not null\015\012UNION SELECT tbl_Unknown_Species.Un"
    "known_Code AS Master_Plant_Code,  tbl_Unknown_Species.Unknown_Code AS Utah_Plant"
    "_Code, (tbl_Unknown_Species.Unknown_Code & \"   \" & tbl_Unknown_Species.Plant_D"
    "escription) AS Utah_Species, 2 AS Sort_Seq\015\012FROM tbl_Unknown_Species\015\012"
    "ORDER BY Sort_Seq, Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Sort_Seq"
        dbBinary "GUID" = Begin
            0x67120e199fb13942a940862d8a20e80b
        End
    End
End
