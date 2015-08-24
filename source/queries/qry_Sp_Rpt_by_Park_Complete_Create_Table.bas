dbMemo "SQL" ="SELECT DISTINCT qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Year, qry_Sp_Rpt_All.Pl"
    "ot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All.Utah_Species, (qry_Sp_Rpt_Al"
    "l.Utah_Species+\"-\"+CStr(qry_Sp_Rpt_All.Year)) AS SpeciesYear, (qry_Sp_Rpt_All."
    "Unit_Code+\"-\"+CStr(qry_Sp_Rpt_All.Plot_ID)+\"-\"+CStr(qry_Sp_Rpt_All.Utah_Spec"
    "ies)) AS ParkPlotSpecies, (qry_Sp_Rpt_All.Unit_Code+\"-\"+CStr(qry_Sp_Rpt_All.Ut"
    "ah_Species)) AS ParkSpecies, (qry_Sp_Rpt_All.Unit_Code+\"-\"+CStr(qry_Sp_Rpt_All"
    ".Plot_ID)) AS ParkPlot INTO temp_Sp_Rpt_by_Park_Complete\015\012FROM qry_Sp_Rpt_"
    "All\015\012ORDER BY qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Plot_ID, qry_Sp_Rpt"
    "_All.Master_Family, qry_Sp_Rpt_All.Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x47e18387b7831f45bb791ed40985d23b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
