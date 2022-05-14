dbMemo "SQL" ="SELECT tbl_Revisit_List.PARK, tbl_Revisit_List.Plot\015\012FROM tbl_Revisit_List"
    "\015\012ORDER BY tbl_Revisit_List.PARK, tbl_Revisit_List.Plot;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd4108c6660271c489e91a77801e9350c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Revisit_List.PARK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Revisit_List.Plot"
        dbLong "AggregateType" ="-1"
    End
End
