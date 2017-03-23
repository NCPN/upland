Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Shrubs.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Shrubs"
    Name ="QC_NDC_NoDataCollected_Shrubs"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Shrubs.Unit_Code"
    Expression ="QC_NDC_NoData_Shrubs.Plot_ID"
    Expression ="QC_NDC_NoData_Shrubs.Start_Date"
    Expression ="QC_NDC_NoData_Shrubs.Transect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Shrubs"
    RightTable ="QC_NDC_NoDataCollected_Shrubs"
    Expression ="QC_NDC_NoData_Shrubs.Unit_Code = QC_NDC_NoDataCollected_Shrubs.Unit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Shrubs"
    RightTable ="QC_NDC_NoDataCollected_Shrubs"
    Expression ="QC_NDC_NoData_Shrubs.Plot_ID = QC_NDC_NoDataCollected_Shrubs.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Shrubs"
    RightTable ="QC_NDC_NoDataCollected_Shrubs"
    Expression ="QC_NDC_NoData_Shrubs.Start_Date = QC_NDC_NoDataCollected_Shrubs.Start_Date"
    Flag =2
    LeftTable ="QC_NDC_NoData_Shrubs"
    RightTable ="QC_NDC_NoDataCollected_Shrubs"
    Expression ="QC_NDC_NoData_Shrubs.Transect = QC_NDC_NoDataCollected_Shrubs.Transect"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Shrubs.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Shrubs.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Shrubs.Start_Date"
    Flag =0
    Expression ="QC_NDC_NoData_Shrubs.Transect"
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
        dbText "Name" ="QC_NDC_NoData_Shrubs.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x49114d4ad97f1b4eb985638bc560c9ec
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Shrubs.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40e009adb05c5441ad6cf72480b07ddd
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Shrubs.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa53f4ce45131849ba402821ed3b4ee8
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Shrubs.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1e0d94966753a041a9a3035d562bac5a
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
    Bottom =136
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =210
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Shrubs"
        Name =""
    End
    Begin
        Left =268
        Top =11
        Right =487
        Bottom =155
        Top =0
        Name ="QC_NDC_NoDataCollected_Shrubs"
        Name =""
    End
End
