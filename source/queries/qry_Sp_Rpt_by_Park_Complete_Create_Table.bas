dbMemo "SQL" ="SELECT DISTINCT qry_Sp_Rpt_All_Revisits.Unit_Code, qry_Sp_Rpt_All_Revisits.Year,"
    " qry_Sp_Rpt_All_Revisits.Plot_ID, qry_Sp_Rpt_All_Revisits.Master_Family, qry_Sp_"
    "Rpt_All_Revisits.Utah_Species, (qry_Sp_Rpt_All_Revisits.Utah_Species+\"-\"+CStr("
    "qry_Sp_Rpt_All_Revisits.Year)) AS SpeciesYear, (qry_Sp_Rpt_All_Revisits.Unit_Cod"
    "e+\"-\"+CStr(qry_Sp_Rpt_All_Revisits.Plot_ID)+\"-\"+CStr(qry_Sp_Rpt_All_Revisits"
    ".Utah_Species)) AS ParkPlotSpecies, (qry_Sp_Rpt_All_Revisits.Unit_Code+\"-\"+CSt"
    "r(qry_Sp_Rpt_All_Revisits.Utah_Species)) AS ParkSpecies, (qry_Sp_Rpt_All_Revisit"
    "s.Unit_Code+\"-\"+CStr(qry_Sp_Rpt_All_Revisits.Plot_ID)) AS ParkPlot INTO temp_S"
    "p_Rpt_by_Park_Complete\015\012FROM qry_Sp_Rpt_All_Revisits\015\012ORDER BY qry_S"
    "p_Rpt_All_Revisits.Unit_Code, qry_Sp_Rpt_All_Revisits.Plot_ID, qry_Sp_Rpt_All_Re"
    "visits.Master_Family, qry_Sp_Rpt_All_Revisits.Utah_Species;\015\012"
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
    Begin
        dbText "Name" ="SpeciesYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a9b0a021021df4c8b1c09a2ad4dae31
        End
    End
    Begin
        dbText "Name" ="ParkPlotSpecies"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7434316a49f5514cb3368a0414fcc04e
        End
    End
    Begin
        dbText "Name" ="ParkSpecies"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04e5d7f9710bf54e9f5582e5ba298468
        End
    End
    Begin
        dbText "Name" ="ParkPlot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xac805cea229f214b8a22b979bd358072
        End
    End
End
