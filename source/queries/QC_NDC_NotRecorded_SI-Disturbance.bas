Operation =1
Option =0
Where ="((([QC_NDC_NoDataCollected_SI-Disturbance].Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_SI-Disturbance"
    Name ="QC_NDC_NoDataCollected_SI-Disturbance"
End
Begin OutputColumns
    Expression ="[QC_NDC_NoData_SI-Disturbance].Unit_Code"
    Expression ="[QC_NDC_NoData_SI-Disturbance].Plot_ID"
    Expression ="[QC_NDC_NoData_SI-Disturbance].Start_Date"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_SI-Disturbance"
    RightTable ="QC_NDC_NoDataCollected_SI-Disturbance"
    Expression ="[QC_NDC_NoData_SI-Disturbance].Unit_Code = [QC_NDC_NoDataCollected_SI-Disturbanc"
        "e].Unit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_SI-Disturbance"
    RightTable ="QC_NDC_NoDataCollected_SI-Disturbance"
    Expression ="[QC_NDC_NoData_SI-Disturbance].Plot_ID = [QC_NDC_NoDataCollected_SI-Disturbance]"
        ".Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_SI-Disturbance"
    RightTable ="QC_NDC_NoDataCollected_SI-Disturbance"
    Expression ="[QC_NDC_NoData_SI-Disturbance].Start_Date = [QC_NDC_NoDataCollected_SI-Disturban"
        "ce].Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="[QC_NDC_NoData_SI-Disturbance].Unit_Code"
    Flag =0
    Expression ="[QC_NDC_NoData_SI-Disturbance].Plot_ID"
    Flag =0
    Expression ="[QC_NDC_NoData_SI-Disturbance].Start_Date"
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
        dbText "Name" ="[QC_NDC_NoData_SI-Disturbance].Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5fea9b98d09cfb4ab9185122a68f7e56
        End
    End
    Begin
        dbText "Name" ="[QC_NDC_NoData_SI-Disturbance].Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x54af1b63c2a02346a4f1f7d7eba4782c
        End
    End
    Begin
        dbText "Name" ="[QC_NDC_NoData_SI-Disturbance].Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb3a3732de42b644a6dd236d9098b2b6
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
    Bottom =112
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =264
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_SI-Disturbance"
        Name =""
    End
    Begin
        Left =326
        Top =12
        Right =587
        Bottom =156
        Top =0
        Name ="QC_NDC_NoDataCollected_SI-Disturbance"
        Name =""
    End
End
