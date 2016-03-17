dbMemo "SQL" ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
    "py.Utah_Species, qryU_Top_Canopy.Lifeform\015\012FROM qryU_Top_Canopy\015\012WHE"
    "RE (((qryU_Top_Canopy.Utah_Species) Is Not Null) AND ((qryU_Top_Canopy.[Lifeform"
    "])='Tree'))\015\012ORDER BY qryU_Top_Canopy.LU_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd3c582acbe24cf438a0ee5b9b9a1ca4f
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
