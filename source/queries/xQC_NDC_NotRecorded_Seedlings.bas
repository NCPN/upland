Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Seedlings.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Seedlings"
    Name ="QC_NDC_NoDataCollected_Seedlings"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Seedlings.Unit_Code"
    Expression ="QC_NDC_NoData_Seedlings.Plot_ID"
    Expression ="QC_NDC_NoData_Seedlings.Start_Date"
    Expression ="QC_NDC_NoData_Seedlings.Transect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Seedlings"
    RightTable ="QC_NDC_NoDataCollected_Seedlings"
    Expression ="QC_NDC_NoData_Seedlings.Unit_Code = QC_NDC_NoDataCollected_Seedlings.Unit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Seedlings"
    RightTable ="QC_NDC_NoDataCollected_Seedlings"
    Expression ="QC_NDC_NoData_Seedlings.Plot_ID = QC_NDC_NoDataCollected_Seedlings.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Seedlings"
    RightTable ="QC_NDC_NoDataCollected_Seedlings"
    Expression ="QC_NDC_NoData_Seedlings.Start_Date = QC_NDC_NoDataCollected_Seedlings.Start_Date"
    Flag =2
    LeftTable ="QC_NDC_NoData_Seedlings"
    RightTable ="QC_NDC_NoDataCollected_Seedlings"
    Expression ="QC_NDC_NoData_Seedlings.Transect = QC_NDC_NoDataCollected_Seedlings.Transect"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Seedlings.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Seedlings.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Seedlings.Start_Date"
    Flag =0
    Expression ="QC_NDC_NoData_Seedlings.Transect"
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
        dbText "Name" ="QC_NDC_NoData_Seedlings.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4fd090f719700640a9fd702f09a8fccb
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Seedlings.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd4689968eda5634fa147147967e8eb49
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Seedlings.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e105aa8653964438b894e11c88d8847
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Seedlings.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf3e5aa30396eaf45a82b8b65b27d9358
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
        Right =225
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Seedlings"
        Name =""
    End
    Begin
        Left =265
        Top =16
        Right =492
        Bottom =160
        Top =0
        Name ="QC_NDC_NoDataCollected_Seedlings"
        Name =""
    End
End
