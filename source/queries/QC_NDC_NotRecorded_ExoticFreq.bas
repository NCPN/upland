Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_ExoticFreq.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_ExoticFreq"
    Name ="QC_NDC_NoDataCollected_ExoticFreq"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_ExoticFreq.Unit_Code"
    Expression ="QC_NDC_NoData_ExoticFreq.Plot_ID"
    Expression ="QC_NDC_NoData_ExoticFreq.Start_Date"
    Expression ="QC_NDC_NoData_ExoticFreq.Transect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_ExoticFreq"
    RightTable ="QC_NDC_NoDataCollected_ExoticFreq"
    Expression ="QC_NDC_NoData_ExoticFreq.Unit_Code = QC_NDC_NoDataCollected_ExoticFreq.Unit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_ExoticFreq"
    RightTable ="QC_NDC_NoDataCollected_ExoticFreq"
    Expression ="QC_NDC_NoData_ExoticFreq.Plot_ID = QC_NDC_NoDataCollected_ExoticFreq.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_ExoticFreq"
    RightTable ="QC_NDC_NoDataCollected_ExoticFreq"
    Expression ="QC_NDC_NoData_ExoticFreq.Start_Date = QC_NDC_NoDataCollected_ExoticFreq.Start_Da"
        "te"
    Flag =2
    LeftTable ="QC_NDC_NoData_ExoticFreq"
    RightTable ="QC_NDC_NoDataCollected_ExoticFreq"
    Expression ="QC_NDC_NoData_ExoticFreq.Transect = QC_NDC_NoDataCollected_ExoticFreq.Transect"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_ExoticFreq.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_ExoticFreq.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_ExoticFreq.Start_Date"
    Flag =0
    Expression ="QC_NDC_NoData_ExoticFreq.Transect"
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
        dbText "Name" ="QC_NDC_NoData_ExoticFreq.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5a701bcd574c6145be810ca4994d4198
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_ExoticFreq.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x02d961f05543d943a585e64c1cd65eda
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_ExoticFreq.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc79b105d40a54e428a6d2b4b32f48151
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_ExoticFreq.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc6b4b3011c0624428dac33c6e90fca90
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
    Bottom =125
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =223
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_ExoticFreq"
        Name =""
    End
    Begin
        Left =265
        Top =11
        Right =498
        Bottom =155
        Top =0
        Name ="QC_NDC_NoDataCollected_ExoticFreq"
        Name =""
    End
End
