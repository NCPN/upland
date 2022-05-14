Operation =1
Option =0
Begin InputTables
    Name ="qry_Sp_Rpt_All"
End
Begin OutputColumns
    Expression ="qry_Sp_Rpt_All.Unit_Code"
    Expression ="qry_Sp_Rpt_All.Plot_ID"
    Expression ="qry_Sp_Rpt_All.Master_Family"
    Expression ="qry_Sp_Rpt_All.Utah_Species"
    Expression ="qry_Sp_Rpt_All.Year"
    Alias ="ParkPlot"
    Expression ="[Unit_Code] & [Plot_ID]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9a1fa43996676544910816580f299a92
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ParkPlot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x14b144b6206cbd4b9665d7e263b25cd0
        End
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =108
    Top =15
    Right =1220
    Bottom =811
    Left =-1
    Top =-1
    Right =1540
    Bottom =357
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qry_Sp_Rpt_All"
        Name =""
    End
End
