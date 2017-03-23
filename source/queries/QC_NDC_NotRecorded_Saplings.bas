Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Saplings.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Saplings"
    Name ="QC_NDC_NoDataCollected_Saplings"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Saplings.Unit_Code"
    Expression ="QC_NDC_NoData_Saplings.Plot_ID"
    Expression ="QC_NDC_NoData_Saplings.Start_Date"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Saplings"
    RightTable ="QC_NDC_NoDataCollected_Saplings"
    Expression ="QC_NDC_NoData_Saplings.Unit_Code = QC_NDC_NoDataCollected_Saplings.Unit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Saplings"
    RightTable ="QC_NDC_NoDataCollected_Saplings"
    Expression ="QC_NDC_NoData_Saplings.Plot_ID = QC_NDC_NoDataCollected_Saplings.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Saplings"
    RightTable ="QC_NDC_NoDataCollected_Saplings"
    Expression ="QC_NDC_NoData_Saplings.Start_Date = QC_NDC_NoDataCollected_Saplings.Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Saplings.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Saplings.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Saplings.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Saplings.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x267aafbc5b807f4ea5f4393099c85197
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Saplings.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x304cf9deeaf98240a9965f1f1d0d6463
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Saplings.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x11c95ad78b0bc64aa5c78702dd878306
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
    Bottom =139
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =229
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Saplings"
        Name =""
    End
    Begin
        Left =288
        Top =12
        Right =528
        Bottom =156
        Top =0
        Name ="QC_NDC_NoDataCollected_Saplings"
        Name =""
    End
End
