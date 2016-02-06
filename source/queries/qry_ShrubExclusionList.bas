dbMemo "SQL" ="SELECT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Family, tlu_NCP"
    "N_Plants.Master_Species, tlu_NCPN_Plants.Lifeform, tlu_NCPN_Plants.Nativity, tlu"
    "_NCPN_Plants.LU_Code, tlu_NCPN_Plants.UT_Family, tlu_NCPN_Plants.Utah_Species\015"
    "\012FROM tlu_NCPN_Plants\015\012WHERE Lifeform IN ('Shrub', 'DwarfShrub')\015\012"
    "AND\015\012Master_PLANT_Code IN (\015\012'YUAN2','YUANK','YUANT','YUBA','YUBA2',"
    "'YUBAV','YUCCA','YUEL',\015\012'YUELU','YUHA','YUHAN','GUMI','GUPO2','GUSA2','GU"
    "TIE','MARE11',\015\012'SYLO','SYMPH','SYOC','SYOR','SYOR2','SYRO','ARPA6','ARPR'"
    ",\015\012'ARPU5','ARUV','QUERC','QUGA','QUHA3','QUPA4','QUTU2',\015\012'RIMO2','"
    "AMUT','PRVI','ROACS','RONU','ROOD','ROSA5','ROWO');\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9d15e212236416428d10d51afe912d1b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "OrderBy" ="[qry_ShrubExclusionList].[Master_Family], [qry_ShrubExclusionList].[Master_PLANT"
    "_Code]"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc6faa48017dfbb4395ec85bad1a34144
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Lifeform"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xae8548a4593d244bb727f6e2b83d10eb
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Nativity"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x778dddd70ab1604092c3c4ea706e5448
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.LU_Code"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x424b40d773e2514a955b584fac8eee73
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x50ad7adc418ae94c8734ba20bf7f87da
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.UT_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67d3392766f79145903ce5eb6669cf08
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb02e2ef3a8ca14b951303da133e22ac
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3439c510fba4ce4cb1367b845423a9e1
        End
    End
End
