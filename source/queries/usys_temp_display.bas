dbMemo "SQL" ="SELECT DISTINCT qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Year, qry_Sp_Rpt_All.Pl"
    "ot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All.Utah_Species\015\012FROM qry"
    "_Sp_Rpt_All\015\012WHERE Unit_Code = 'BLCA' \015\012AND Plot_ID = 41 \015\012AND"
    " \015\012qry_Sp_Rpt_All.Year = Year(#9/23/2008#)\015\012ORDER BY qry_Sp_Rpt_All."
    "Unit_Code, qry_Sp_Rpt_All.Plot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All."
    "Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
