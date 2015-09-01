dbMemo "SQL" ="SELECT *\015\012FROM qry_Sp_Rpt_Quadrat\015\012WHERE Unit_Code IS NULL \015\012O"
    "R\015\012Plot_ID IS NULL\015\012OR\015\012Master_Family IS NULL\015\012OR \015\012"
    "Utah_Species IS NULL\015\012OR\015\012Year IS NULL;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5be7aaae1da5df4a8e3cc296797c96e3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qry_Sp_Rpt_Quadrat.Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd4a8f383cdd8d046b88cd322501cd84d
        End
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_Quadrat.tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb55a6b454ba0ee47a1a2d3942e0a203c
        End
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_Quadrat.tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x118a48d27f60e14da5f8ebf5bbf425f1
        End
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_Quadrat.tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x82165d2f1de6124393c1312b97661be8
        End
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_Quadrat.tlu_NCPN_Plants.Master_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xba7d2ce77dae584fa840b6d1878ccf16
        End
    End
End
