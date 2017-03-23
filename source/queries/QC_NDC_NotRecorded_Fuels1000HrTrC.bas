Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Fuels1000HrTrC.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Fuels1000HrTrC"
    Name ="QC_NDC_NoDataCollected_Fuels1000HrTrC"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Unit_Code"
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Plot_ID"
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Start_Date"
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.FuelsTransect"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrC"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrC"
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Unit_Code = QC_NDC_NoDataCollected_Fuels1000HrTrC.U"
        "nit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrC"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrC"
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Plot_ID = QC_NDC_NoDataCollected_Fuels1000HrTrC.Plo"
        "t_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Fuels1000HrTrC"
    RightTable ="QC_NDC_NoDataCollected_Fuels1000HrTrC"
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Start_Date = QC_NDC_NoDataCollected_Fuels1000HrTrC."
        "Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTrC.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrC.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe92fa949c4940f4580bacfe34baf9d3c
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrC.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x55c78c10bc0fea46aa996fa664a011ef
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrC.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xffe8225f770fe047bb5903b23bbab106
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTrC.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x717112d23b4b424abf4515781b802296
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
    Bottom =146
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =258
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Fuels1000HrTrC"
        Name =""
    End
    Begin
        Left =343
        Top =13
        Right =607
        Bottom =157
        Top =0
        Name ="QC_NDC_NoDataCollected_Fuels1000HrTrC"
        Name =""
    End
End
