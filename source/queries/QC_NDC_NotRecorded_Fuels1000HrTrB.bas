Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Fuels1000HrTrB.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Fuels1000HrTrB"
    Name ="QC_NDC_NoDataCollected_Fuels1000HrTrB"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Unit_Code"
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Plot_ID"
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Start_Date"
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.FuelsTransect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrB"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrB"
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Unit_Code = QC_NDC_NoDataCollected_Fuels1000HrTrB.U"
        "nit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrB"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrB"
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Plot_ID = QC_NDC_NoDataCollected_Fuels1000HrTrB.Plo"
        "t_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrB"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrB"
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Start_Date = QC_NDC_NoDataCollected_Fuels1000HrTrB."
        "Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrB.Start_Date"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrB.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe2c1288944a14642b7178fe2f5442fdf
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrB.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf39134995126a84ba67225582576afd5
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrB.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5848006b82d6d54f88bd88e073ed3720
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrB.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53d06223ed38aa44b70ff65c52841609
        End
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1005
    Bottom =533
    Left =-1
    Top =-1
    Right =989
    Bottom =141
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =258
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Fuels1000HrTrB"
        Name =""
    End
    Begin
        Left =324
        Top =12
        Right =606
        Bottom =156
        Top =0
        Name ="QC_NDC_NoDataCollected_Fuels1000HrTrB"
        Name =""
    End
End
