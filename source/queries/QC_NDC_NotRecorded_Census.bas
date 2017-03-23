Operation =1
Option =0
Where ="(((QC_NDC_NoDataCollected_Census.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_NoData_Census"
    Name ="QC_NDC_NoDataCollected_Census"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Census.Unit_Code"
    Expression ="QC_NDC_NoData_Census.Plot_ID"
    Expression ="QC_NDC_NoData_Census.Start_Date"
End
Begin Joins
    LeftTable ="QC_NDC_NoData_Census"
    RightTable ="QC_NDC_NoDataCollected_Census"
    Expression ="QC_NDC_NoData_Census.Unit_Code = QC_NDC_NoDataCollected_Census.Unit_Code"
    Flag =2
    LeftTable ="QC_NDC_NoData_Census"
    RightTable ="QC_NDC_NoDataCollected_Census"
    Expression ="QC_NDC_NoData_Census.Plot_ID = QC_NDC_NoDataCollected_Census.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_NoData_Census"
    RightTable ="QC_NDC_NoDataCollected_Census"
    Expression ="QC_NDC_NoData_Census.Start_Date = QC_NDC_NoDataCollected_Census.Start_Date"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Census.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Census.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Census.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Census.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b18829eee0fed47b5b4e7ce817abe15
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Census.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6542e75e76a7540841b9920d6b6eab8
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Census.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6ad1dab585510641803f8819cd11c2a9
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
    Bottom =140
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =212
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Census"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =464
        Bottom =156
        Top =0
        Name ="QC_NDC_NoDataCollected_Census"
        Name =""
    End
End
