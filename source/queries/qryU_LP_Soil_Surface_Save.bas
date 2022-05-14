dbMemo "SQL" ="SELECT tlu_LP_Soil_Surface.Surface_Code AS Master_Plant_Code,  tlu_LP_Soil_Surfa"
    "ce.Surface_Code AS Utah_Plant_Code, tlu_LP_Soil_Surface.Surface_Description AS U"
    "tah_Species, 1 as Sort_Seq\015\012FROM tlu_LP_Soil_Surface\015\012UNION SELECT \""
    "NVR\"  AS Master_Plant_Code,  \"NVR\" AS Utah_Plant_Code, \"No Value Recorded\" "
    "AS Utah_Species, 3 as Sort_Seq FROM tlu_LP_Soil_Surface \015\012UNION SELECT tlu"
    "_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Utah_Species as Utah_Plant_Code,"
    "  tlu_NCPN_Plants.Utah_Species, 2 as Sort_Seq \015\012FROM tlu_NCPN_Plants WHERE"
    " tlu_NCPN_Plants.Utah_PLANT_Code is not null \015\012UNION SELECT tbl_Unknown_Sp"
    "ecies.Unknown_Code AS Master_Plant_Code,  tbl_Unknown_Species.Unknown_Code AS Ut"
    "ah_Plant_Code, (tbl_Unknown_Species.Unknown_Code & \"   \" & tbl_Unknown_Species"
    ".Plant_Description) AS Utah_Species, 2 AS Sort_Seq\015\012FROM tbl_Unknown_Speci"
    "es\015\012ORDER BY Sort_Seq, Utah_PLANT_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x254b7aab86f00243b1c3af4e9aadabcc
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
        dbText "Name" ="Utah_Plant_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x73cccd4ffd2bf9468fb29c044258f606
        End
    End
    Begin
        dbText "Name" ="Sort_Seq"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53b39c1e9006aa42a97ece4dcb0268f3
        End
    End
End
