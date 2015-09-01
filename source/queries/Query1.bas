Operation =1
Option =1
Where ="(((qry_Sp_Rpt_All.[Unit_Code]) Is Null)) OR (((qry_Sp_Rpt_All.[Plot_ID]) Is Null"
    ")) OR (((qry_Sp_Rpt_All.[Master_Family]) Is Null)) OR (((qry_Sp_Rpt_All.[Utah_Sp"
    "ecies]) Is Null)) OR (((qry_Sp_Rpt_All.[Year]) Is Null))"
Begin InputTables
    Name ="qry_Sp_Rpt_All"
End
Begin OutputColumns
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x40672db85e465c4d8e443a0bb72c66fb
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.qry_Sp_Rpt_BT_Add_Sp.tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.qry_Sp_Rpt_BT_Add_Sp.tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.qry_Sp_Rpt_BT_Add_Sp.tlu_NCPN_Plants.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.qry_Sp_Rpt_BT_Add_Sp.tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Sp_Rpt_All.qry_Sp_Rpt_BT_Add_Sp.Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =83
    Top =15
    Right =533
    Bottom =280
    Left =-1
    Top =-1
    Right =412
    Bottom =184
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="qry_Sp_Rpt_All"
        Name =""
    End
End
