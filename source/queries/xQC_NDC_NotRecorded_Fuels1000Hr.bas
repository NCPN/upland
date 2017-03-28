Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Fuels1000Hr.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Fuels1000Hr"
    Name ="QC_NDC_NoDataCollected_Fuels1000Hr"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Fuels1000Hr.Unit_Code"
    Expression ="QC_NDC_NoData_Fuels1000Hr.Plot_ID"
    Expression ="QC_NDC_NoData_Fuels1000Hr.Start_Date"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Fuels1000Hr"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000Hr"
    Expression ="QC_NDC_NoData_Fuels1000Hr.Unit_Code = QC_NDC_NoDataCollected_Fuels1000Hr.Unit_Co"
        "de"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000Hr"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000Hr"
    Expression ="QC_NDC_NoData_Fuels1000Hr.Plot_ID = QC_NDC_NoDataCollected_Fuels1000Hr.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000Hr"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000Hr"
    Expression ="QC_NDC_NoData_Fuels1000Hr.Start_Date = QC_NDC_NoDataCollected_Fuels1000Hr.Start_"
        "Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Fuels1000Hr.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000Hr.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000Hr.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Fuels1000Hr.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x456184d2c1bd254bbb4d0371f7b9d4bd
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000Hr.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc5ae5be13953464ca9997d90e39b89e1
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000Hr.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc5761cbc8250684ba994d0720951e344
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
    Bottom =138
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =234
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Fuels1000Hr"
        Name =""
    End
    Begin
        Left =275
        Top =13
        Right =511
        Bottom =157
        Top =0
        Name ="QC_NDC_NoDataCollected_Fuels1000Hr"
        Name =""
    End
End
