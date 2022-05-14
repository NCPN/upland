Operation =1
Option =0
Begin InputTables
    Name ="qry_present_by_year"
End
Begin OutputColumns
    Expression ="qry_present_by_year.Unit_Code"
    Expression ="qry_present_by_year.Plot_ID"
    Expression ="qry_present_by_year.Master_Family"
    Expression ="qry_present_by_year.Utah_Species"
    Alias ="Present in 2006"
    Expression ="Max(qry_present_by_year.[2006])"
    Alias ="Present in 2007"
    Expression ="Max(qry_present_by_year.[2007])"
End
Begin Groups
    Expression ="qry_present_by_year.Unit_Code"
    GroupLevel =0
    Expression ="qry_present_by_year.Plot_ID"
    GroupLevel =0
    Expression ="qry_present_by_year.Master_Family"
    GroupLevel =0
    Expression ="qry_present_by_year.Utah_Species"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5099d573f5c38f4fa56edd21717207dc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_present_by_year.Master_Family"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Present in 2006"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdb5ea07a9eea6843a5199b580864218b
        End
    End
    Begin
        dbText "Name" ="Present in 2007"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0e26a4e1e52e45479f524ac885f42177
        End
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1002
    Bottom =327
    Left =-1
    Top =-1
    Right =960
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="qry_present_by_year"
        Name =""
    End
End
