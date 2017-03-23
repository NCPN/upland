dbMemo "SQL" ="SELECT Unit_Code, Plot_ID, Start_Date, \"Exotic Frequency\" AS Method, Transect\015"
    "\012FROM QC_NDC_NotRecorded_ExoticFreq\015\012UNION\015\012SELECT Unit_Code, Plo"
    "t_ID, Start_Date, \"Shrubs\" AS Method, Transect\015\012FROM QC_NDC_NotRecorded_"
    "Shrubs\015\012UNION\015\012SELECT Unit_Code, Plot_ID, Start_Date, \"Seedlings\" "
    "AS Method, Transect\015\012FROM QC_NDC_NotRecorded_Seedlings\015\012UNION\015\012"
    "SELECT Unit_Code, Plot_ID, Start_Date, \"Saplings\" AS Method, \"n/a\" AS Transe"
    "ct\015\012FROM QC_NDC_NotRecorded_Saplings\015\012UNION\015\012SELECT Unit_Code,"
    " Plot_ID, Start_Date, \"Census\" AS Method, \"n/a\" AS Transect\015\012FROM QC_N"
    "DC_NotRecorded_Census\015\012UNION\015\012SELECT Unit_Code, Plot_ID, Start_Date,"
    " \"SI Exotics\" AS Method, \"n/a\" AS Transect\015\012FROM [QC_NDC_NotRecorded_S"
    "I-Exotics]\015\012UNION\015\012SELECT Unit_Code, Plot_ID, Start_Date, \"SI Distu"
    "rbance\" AS Method, \"n/a\" AS Transect\015\012FROM [QC_NDC_NotRecorded_SI-Distu"
    "rbance]\015\012UNION\015\012SELECT Unit_Code, Plot_ID, Start_Date, \"Fuels 1000 "
    "Hr\" AS Method, \"All\" AS Transect\015\012FROM QC_NDC_NotRecorded_Fuels1000Hr\015"
    "\012UNION\015\012SELECT Unit_Code, Plot_ID, Start_Date, \"Fuels 1000 Hr\" AS Met"
    "hod, \"A\" AS Transect\015\012FROM QC_NDC_NotRecorded_Fuels1000HrTrA\015\012UNIO"
    "N\015\012SELECT Unit_Code, Plot_ID, Start_Date, \"Fuels 1000 Hr\" AS Method, \"B"
    "\" AS Transect\015\012FROM QC_NDC_NotRecorded_Fuels1000HrTrB\015\012UNION\015\012"
    "SELECT Unit_Code, Plot_ID, Start_Date, \"Fuels 1000 Hr\" AS Method, \"C\" AS Tra"
    "nsect\015\012FROM QC_NDC_NotRecorded_Fuels1000HrTrC\015\012UNION SELECT Unit_Cod"
    "e, Plot_ID, Start_Date, \"Fuels 1000 Hr\" AS Method, \"D\" AS Transect\015\012FR"
    "OM QC_NDC_NotRecorded_Fuels1000HrTrD;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc5ddf0433f7c3241855692fb8cc0343d
        End
    End
    Begin
        dbText "Name" ="Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b85623c2940564490e04de82a30c082
        End
    End
    Begin
        dbText "Name" ="Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00fc5bbb6335bb4caeca1ad67598aba3
        End
    End
    Begin
        dbText "Name" ="Method"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd46f68acaa149a4ca9f4cd66d3d951f2
        End
    End
    Begin
        dbText "Name" ="Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x419d5148f7687449ac4caad8032e933b
        End
    End
End
