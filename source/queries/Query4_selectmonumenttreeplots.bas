dbMemo "SQL" ="SELECT DISTINCT [Unit_Code] & [Plot_ID] AS ParkPlot\015\012FROM tbl_Locations IN"
    "NER JOIN tbl_Monument ON tbl_Locations.Location_ID = tbl_Monument.Location_ID;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x24e5ac7a8578454ba793d78b856aa977
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ParkPlot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x76bca0a1258187439357da73ecb4860d
        End
    End
End
