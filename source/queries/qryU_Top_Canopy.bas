dbMemo "SQL" ="SELECT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.LU_Code,  tlu_NCPN_Pla"
    "nts.Utah_Species, 2 as Sort_Seq\015\012FROM tlu_NCPN_Plants WHERE tlu_NCPN_Plant"
    "s.LU_Code is not null\015\012UNION SELECT tbl_Unknown_Species.Unknown_Code AS Ma"
    "ster_Plant_Code,  tbl_Unknown_Species.Unknown_Code AS LU_Code, (tbl_Unknown_Spec"
    "ies.Unknown_Code & \"   \" & tbl_Unknown_Species.Plant_Description) AS Utah_Spec"
    "ies, 2 AS Sort_Seq\015\012FROM tbl_Unknown_Species\015\012ORDER BY Sort_Seq, Uta"
    "h_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x51b9647c0cf3d64bb6c497459459cc34
End
Begin
    Begin
        dbText "Name" ="Sort_Seq"
        dbBinary "GUID" = Begin
            0x67120e199fb13942a940862d8a20e80b
        End
    End
End
