Operation =1
Option =0
Where ="(((QC_NDC_Fuels1000Hr_Transects.Unit_Code) Is Null))"
Begin InputTables
    Name ="QC_NDC_Fuels1000Hr_Transects_All"
    Name ="QC_NDC_Fuels1000Hr_Transects"
End
Begin OutputColumns
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Unit_Code"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Plot_ID"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Start_Date"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.FuelsTransect"
End
Begin Joins
    LeftTable ="QC_NDC_Fuels1000Hr_Transects_All"
    RightTable ="QC_NDC_Fuels1000Hr_Transects"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Unit_Code = QC_NDC_Fuels1000Hr_Transects.Unit_C"
        "ode"
    Flag =2
    LeftTable ="QC_NDC_Fuels1000Hr_Transects_All"
    RightTable ="QC_NDC_Fuels1000Hr_Transects"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Plot_ID = QC_NDC_Fuels1000Hr_Transects.Plot_ID"
    Flag =2
    LeftTable ="QC_NDC_Fuels1000Hr_Transects_All"
    RightTable ="QC_NDC_Fuels1000Hr_Transects"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Start_Date = QC_NDC_Fuels1000Hr_Transects.Start"
        "_Date"
    Flag =2
    LeftTable ="QC_NDC_Fuels1000Hr_Transects_All"
    RightTable ="QC_NDC_Fuels1000Hr_Transects"
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.FuelsTransect = QC_NDC_Fuels1000Hr_Transects.Tr"
        "ansect"
    Flag =2
End
Begin OrderBy
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Unit_Code"
    Flag =0
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Plot_ID"
    Flag =0
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.Start_Date"
    Flag =0
    Expression ="QC_NDC_Fuels1000Hr_Transects_All.FuelsTransect"
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
        dbText "Name" ="QC_NDC_Fuels1000Hr_Transects_All.FuelsTransect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4bdd601dc192e14dbf8e572f1786c42a
        End
    End
    Begin
        dbText "Name" ="QC_NDC_Fuels1000Hr_Transects_All.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb22970eb6346324fa1a3836a76f15a46
        End
    End
    Begin
        dbText "Name" ="QC_NDC_Fuels1000Hr_Transects_All.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c7e5511d7552e43912eee85f3cec2fd
        End
    End
    Begin
        dbText "Name" ="QC_NDC_Fuels1000Hr_Transects_All.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7d5b016400dbda499c620049c963dca0
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
        Right =192
        Bottom =156
        Top =0
        Name ="QC_NDC_Fuels1000Hr_Transects_All"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="QC_NDC_Fuels1000Hr_Transects"
        Name =""
    End
End
