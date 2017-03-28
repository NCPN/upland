dbMemo "SQL" ="SELECT QC_NDC_NoData_Fuels1000HrTrB.Unit_Code, QC_NDC_NoData_Fuels1000HrTrB.Plot"
    "_ID, QC_NDC_NoData_Fuels1000HrTrB.Start_Date, QC_NDC_NoData_Fuels1000HrTrB.Fuels"
    "Transect\015\012FROM QC_NDC_NoData_Fuels1000HrTrB LEFT JOIN QC_NDC_NoDataCollect"
    "ed_Fuels1000HrTrB ON (QC_NDC_NoData_Fuels1000HrTrB.Unit_Code = QC_NDC_NoDataColl"
    "ected_Fuels1000HrTrB.Unit_Code) AND (QC_NDC_NoData_Fuels1000HrTrB.Plot_ID = QC_N"
    "DC_NoDataCollected_Fuels1000HrTrB.Plot_ID) AND (QC_NDC_NoData_Fuels1000HrTrB.Sta"
    "rt_Date = QC_NDC_NoDataCollected_Fuels1000HrTrB.Start_Date)\015\012WHERE (((QC_N"
    "DC_NoDataCollected_Fuels1000HrTrB.Unit_Code) Is Null))\015\012ORDER BY QC_NDC_No"
    "Data_Fuels1000HrTrB.Unit_Code, QC_NDC_NoData_Fuels1000HrTrB.Plot_ID, QC_NDC_NoDa"
    "ta_Fuels1000HrTrB.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1466f4c4a01c784c8bc5faec13116d97
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
