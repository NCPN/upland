dbMemo "SQL" ="SELECT DISTINCT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.LU_Code,  tlu"
    "_NCPN_Plants.Utah_Species, 2 as Sort_Seq, Lifeform, Nativity\015\012FROM tlu_NCP"
    "N_Plants WHERE tlu_NCPN_Plants.LU_Code IS NOT NULL\015\012UNION SELECT DISTINCT "
    "tbl_Unknown_Species.Unknown_Code AS Master_Plant_Code,  tbl_Unknown_Species.Unkn"
    "own_Code AS LU_Code, (tbl_Unknown_Species.Plant_Type & \" - \" & tbl_Unknown_Spe"
    "cies.Plant_Description) AS Utah_Species, 2 AS Sort_Seq, Plant_Type AS Lifeform, "
    "'' AS Nativity\015\012FROM tbl_Unknown_Species\015\012ORDER BY Sort_Seq, Utah_Sp"
    "ecies;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x51b9647c0cf3d64bb6c497459459cc34
End
dbMemo "OrderBy" ="[qryU_Top_Canopy].[Master_PLANT_Code]"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Sort_Seq"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67120e199fb13942a940862d8a20e80b
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2328"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Lifeform"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x97146031993d1f4383bc1df2c3d6ef80
        End
    End
    Begin
        dbText "Name" ="Nativity"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x011c90156cc8804d801fee424851ed99
        End
    End
    Begin
        dbText "Name" ="Master_PLANT_Code"
        dbBinary "GUID" = Begin
            0xde50ad42518b3a4e94936bb974a260bd
        End
    End
    Begin
        dbText "Name" ="LU_Code"
        dbBinary "GUID" = Begin
            0xc93655d15a49d14b81c034e3066ebd51
        End
    End
    Begin
        dbText "Name" ="Utah_Species"
        dbBinary "GUID" = Begin
            0x30567f4faa91334ca82f381ecade6fb7
        End
    End
End
