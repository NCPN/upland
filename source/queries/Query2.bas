dbMemo "SQL" ="SELECT \015\012qryU_Top_Canopy.Master_PLANT_Code, \015\012qryU_Top_Canopy.LU_Cod"
    "e, \015\012qryU_Top_Canopy.Utah_Species,   \015\012qryU_Top_Canopy.Lifeform \015"
    "\012FROM qryU_Top_Canopy \015\012WHERE (((qryU_Top_Canopy.Utah_Species) Is Not N"
    "ull) AND ((qryU_Top_Canopy.[Lifeform])='Tree')) \015\012ORDER BY qryU_Top_Canopy"
    ".LU_Code  \015\012UNION ALL (SELECT \015\012tbl_Unknown_Species.Unknown_Code, \015"
    "\012tbl_Unknown_Species.Unknown_Code,   \015\012tbl_Unknown_Species.Plant_Type, "
    "\015\012tbl_Unknown_Species.Plant_Type AS Lifeform \015\012FROM tbl_Unknown_Spec"
    "ies \015\012WHERE tbl_Unknown_Species.Plant_Type IN ('Tree','Other') OR tbl_Unkn"
    "own_Species.Plant_Type IS NULL ORDER BY tbl_Unknown_Species.Unknown_Code);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "OrderBy" ="[Query2].[Master_PLANT_Code]"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf184d025b387bb4680a3084c828a64e5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryU_Top_Canopy.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryU_Top_Canopy.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryU_Top_Canopy.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryU_Top_Canopy.Lifeform"
        dbLong "AggregateType" ="-1"
    End
End
