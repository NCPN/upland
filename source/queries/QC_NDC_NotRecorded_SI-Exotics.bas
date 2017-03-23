Operation =1
Option =0
Where ="((([QC_NDC_NoDataCollected_SI-Exotics].Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_SI-Exotics"
    Name ="QC_NDC_NoDataCollected_SI-Exotics"
End
Begin OutputColumns
    Expression ="[QC_NDC_NoData_SI-Exotics].Unit_Code"
    Expression ="[QC_NDC_NoData_SI-Exotics].Plot_ID"
    Expression ="[QC_NDC_NoData_SI-Exotics].Start_Date"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_SI-Exotics"
    RightTable ="QC_NDC_NoDataCollected_SI-Exotics"
    Expression ="[QC_NDC_NoData_SI-Exotics].Unit_Code = [QC_NDC_NoDataCollected_SI-Exotics].Unit_"
        "Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_SI-Exotics"
    RightTable ="QC_NDC_NoDataCollected_SI-Exotics"
    Expression ="[QC_NDC_NoData_SI-Exotics].Plot_ID = [QC_NDC_NoDataCollected_SI-Exotics].Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_SI-Exotics"
    RightTable ="QC_NDC_NoDataCollected_SI-Exotics"
    Expression ="[QC_NDC_NoData_SI-Exotics].Start_Date = [QC_NDC_NoDataCollected_SI-Exotics].Star"
        "t_Date"
    Flag =2
End
Begin OrderBy
    Expression ="[QC_NDC_NoData_SI-Exotics].Unit_Code"
    Flag =0
    Expression ="[QC_NDC_NoData_SI-Exotics].Plot_ID"
    Flag =0
    Expression ="[QC_NDC_NoData_SI-Exotics].Start_Date"
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
        dbText "Name" ="[QC_NDC_NoData_SI-Exotics].Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d2710eab33f984ca3ca89fbc8aafdf1
        End
    End
    Begin
        dbText "Name" ="[QC_NDC_NoData_SI-Exotics].Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6c1dc8960a256046aed97e5c2a27a4ae
        End
    End
    Begin
        dbText "Name" ="[QC_NDC_NoData_SI-Exotics].Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06c97b9fbd316848892479df203d9347
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
    Bottom =129
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =233
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_SI-Exotics"
        Name =""
    End
    Begin
        Left =287
        Top =13
        Right =525
        Bottom =150
        Top =0
        Name ="QC_NDC_NoDataCollected_SI-Exotics"
        Name =""
    End
End
