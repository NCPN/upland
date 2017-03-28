Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Fuels1000HrTrA.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Fuels1000HrTrA"
    Name ="QC_NDC_NoDataCollected_Fuels1000HrTrA"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Unit_Code"
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Plot_ID"
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Start_Date"
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.FuelsTransect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrA"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrA"
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Unit_Code = QC_NDC_NoDataCollected_Fuels1000HrTrA.U"
        "nit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrA"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrA"
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Plot_ID = QC_NDC_NoDataCollected_Fuels1000HrTrA.Plo"
        "t_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrA"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrA"
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Start_Date = QC_NDC_NoDataCollected_Fuels1000HrTrA."
        "Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrA.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrA.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x65fa01e7a1d87744af51a6cbe7481eb9
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrA.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf3b81da33b059748b6f65e24196603c0
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrA.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x159952e55a4d08438d449e7c89e0a13f
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrA.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8f6fa2d2ff05b9499e681dded2ac5d7c
        End
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1044
    Bottom =533
    Left =-1
    Top =-1
    Right =1028
    Bottom =119
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =266
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Fuels1000HrTrA"
        Name =""
    End
    Begin
        Left =303
        Top =7
        Right =568
        Bottom =151
        Top =0
        Name ="QC_NDC_NoDataCollected_Fuels1000HrTrA"
        Name =""
    End
End
