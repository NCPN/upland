Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Master_Stratification"
    Expression ="tbl_Locations.Primary_Eco_Site"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Elevation"
    Expression ="tbl_Locations.Plot_Slope"
    Expression ="tbl_Locations.Azimuth"
    Expression ="tbl_Locations.E_Coord"
    Expression ="tbl_Locations.N_Coord"
    Expression ="tbl_Events.Start_Date"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query1].[Start_Date] DESC, [Query1].[Unit_Code], [Query1].[Master_Stratificatio"
    "n], [Query1].[Primary_Eco_Site], [Query1].[Plot_ID]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x60386c9d04579d45aeb069a67014fb18
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Elevation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e7bf809b98c874c83bff711b01fe6d1
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Slope"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5c7e7a7d4365ce4291b2704e8a99c217
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbc9af4204f59804199119ad9be88e5d8
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Master_Stratification"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xecf5a7aaf59afe4a9e6433a812280415
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Primary_Eco_Site"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6e034669d04e97408b1a52d3de95b65d
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb26842f8281b6c4b97206eccd11a42bc
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Azimuth"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x10046c7a83037a468c3583008f3cd73f
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6ea40b52f2de494bb91de97de9bb72c9
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc762cfee79f4964b96c1a3c7ba72e356
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4f2545cdbe2e7d46893acf954c976cc9
        End
    End
End
Begin
    State =2
    Left =-4
    Top =-30
    Right =1065
    Bottom =762
    Left =-1
    Top =-1
    Right =1041
    Bottom =313
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
