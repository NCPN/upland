dbMemo "SQL" ="SELECT tlu_LP_Soil_Surface.Surface_Code AS Master_Plant_Code,  tlu_LP_Soil_Surfa"
    "ce.Surface_Code AS LU_Code, tlu_LP_Soil_Surface.Surface_Description AS Utah_Spec"
    "ies, 1 as Sort_Seq\015\012FROM tlu_LP_Soil_Surface\015\012UNION SELECT tlu_NCPN_"
    "Plants.Master_PLANT_Code, tlu_NCPN_Plants.LU_Code,  tlu_NCPN_Plants.Utah_Species"
    ", 2 as Sort_Seq \015\012FROM tlu_NCPN_Plants WHERE tlu_NCPN_Plants.Utah_PLANT_Co"
    "de is not null \015\012UNION SELECT tbl_Unknown_Species.Unknown_Code AS Master_P"
    "lant_Code,  tbl_Unknown_Species.Unknown_Code AS LU_Code, (tbl_Unknown_Species.Un"
    "known_Code & \"   \" & tbl_Unknown_Species.Plant_Description) AS Utah_Species, 2"
    " AS Sort_Seq\015\012FROM tbl_Unknown_Species\015\012ORDER BY Sort_Seq, LU_Code;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdfedc937c1570348893449004583f38b
End
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
        dbText "Name" ="Sort_Seq"
        dbBinary "GUID" = Begin
            0x53b39c1e9006aa42a97ece4dcb0268f3
        End
    End
    Begin
        dbText "Name" ="LU_Code"
        dbBinary "GUID" = Begin
            0x97253bd6a7e3f14590108ebb49a22ab9
        End
    End
End
