dbMemo "SQL" ="SELECT tlu_LP_Soil_Surface.Surface_Code AS Master_Plant_Code,  tlu_LP_Soil_Surfa"
    "ce.Surface_Code AS Utah_Plant_Code,  (tlu_LP_Soil_Surface.Surface_Code & \" - \""
    " & tlu_LP_Soil_Surface.Surface_Description) AS Utah_Species, 1 as Sort_Seq\015\012"
    "FROM tlu_LP_Soil_Surface WHERE  tlu_LP_Soil_Surface.LC_Code = 1\015\012UNION SEL"
    "ECT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Utah_PLANT_Code,  tlu_NCP"
    "N_Plants.Utah_Species, 2 as Sort_Seq\015\012FROM tlu_NCPN_Plants WHERE tlu_NCPN_"
    "Plants.Utah_PLANT_Code is not null\015\012UNION SELECT tbl_Unknown_Species.Unkno"
    "wn_Code AS Master_Plant_Code,  tbl_Unknown_Species.Unknown_Code AS Utah_Plant_Co"
    "de, (tbl_Unknown_Species.Unknown_Code & \"   \" & tbl_Unknown_Species.Plant_Desc"
    "ription) AS Utah_Species, 2 AS Sort_Seq\015\012FROM tbl_Unknown_Species\015\012O"
    "RDER BY Sort_Seq, Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Utah_Species"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0x9e0b6af8c83bb240ae123d6b5ede9676
        End
    End
    Begin
        dbText "Name" ="Master_Plant_Code"
        dbBinary "GUID" = Begin
            0x24012dfa53733a408f214b9b6f21abeb
        End
    End
    Begin
        dbText "Name" ="Utah_Plant_Code"
        dbBinary "GUID" = Begin
            0x73cccd4ffd2bf9468fb29c044258f606
        End
    End
    Begin
        dbText "Name" ="Sort_Seq"
        dbBinary "GUID" = Begin
            0x67120e199fb13942a940862d8a20e80b
        End
    End
End
