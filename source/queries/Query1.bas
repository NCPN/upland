dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), pid Long, vdate DateTime;\015\012SELECT DISTINCT q"
    "ry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Year, qry_Sp_Rpt_All.Plot_ID, qry_Sp_Rpt"
    "_All.Master_Family, qry_Sp_Rpt_All.Utah_Species\015\012FROM qry_Sp_Rpt_All\015\012"
    "WHERE Unit_Code = 'ARCH' \015\012AND Plot_ID = 1 \015\012AND \015\012qry_Sp_Rpt_"
    "All.Year = Year(#4/10/2017#)\015\012ORDER BY qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rp"
    "t_All.Plot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All.Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x527bad5a8cfbf248af01563c6fbc134e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
