Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Fuels1000HrTrD.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Fuels1000HrTrD"
    Name ="QC_NDC_NoDataCollected_Fuels1000HrTrD"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Unit_Code"
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Plot_ID"
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Start_Date"
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.FuelsTransect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrD"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrD"
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Unit_Code = QC_NDC_NoDataCollected_Fuels1000HrTrD.U"
        "nit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrD"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrD"
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Plot_ID = QC_NDC_NoDataCollected_Fuels1000HrTrD.Plo"
        "t_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrD"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrD"
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Start_Date = QC_NDC_NoDataCollected_Fuels1000HrTrD."
        "Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrD.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrD.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x54308bda5d7b5d4dbe685e11a5017829
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrD.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d37bf49f6f8b54fa166279eff2dd686
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrD.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x63e4ddcf487b124aa347bf9dd30f1b69
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrD.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6bf161adbcda9d4ab8a94100e81d8da4
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
    Bottom =137
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =263
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Fuels1000HrTrD"
        Name =""
    End
    Begin
        Left =315
        Top =8
        Right =582
        Bottom =152
        Top =0
        Name ="QC_NDC_NoDataCollected_Fuels1000HrTrD"
        Name =""
    End
End
