Operation =1
Option =0
Where ="(((QC_NDC_NoData_Fuels1000HrTransects.FuelsTransect)=\"D\"))"
Begin InputTables
    Name ="QC_NDC_NoData_Fuels1000HrTransects"
End
Begin OutputColumns
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.Unit_Code"
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.Plot_ID"
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.Start_Date"
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.FuelsTransect"
End
Begin OrderBy
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.Unit_Code"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.Plot_ID"
    Flag =0
    Expression ="QC_NDC_NoData_Fuels1000HrTransects.Start_Date"
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
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTransects.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc509b259e0158741b466b4dca6d10ef9
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTransects.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xacb5eda12a7ae743b4a933f6f2c49345
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTransects.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xef9c1931de4e5d4f989f4ab36dc42745
        End
    End
    Begin
        dbText "Name" ="QC_NDC_NoData_Fuels1000HrTransects.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x63751ad599afb24fbe91859433987fb5
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
    Bottom =84
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =318
        Bottom =156
        Top =0
        Name ="QC_NDC_NoData_Fuels1000HrTransects"
        Name =""
    End
End
