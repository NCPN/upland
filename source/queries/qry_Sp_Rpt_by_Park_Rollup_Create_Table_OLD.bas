dbMemo "SQL" ="SELECT DISTINCT temp_Sp_Rpt_by_Park_Complete.Unit_Code, temp_Sp_Rpt_by_Park_Comp"
    "lete.Plot_ID, temp_Sp_Rpt_by_Park_Complete.Master_Family, temp_Sp_Rpt_by_Park_Co"
    "mplete.Utah_Species, ConcatRelated(\"Year\",\"temp_Sp_Rpt_by_Park_Complete\",\"P"
    "arkSpecies='\"+ParkSpecies+\"' and \015ParkPlot='\"+ParkPlot+\"'\",'',\"|\") AS "
    "SpeciesYears, temp_Sp_Rpt_by_Park_Complete.ParkPlotSpecies, temp_Sp_Rpt_by_Park_"
    "Complete.ParkPlot INTO temp_Sp_Rpt_by_Park_Rollup\015\012FROM temp_Sp_Rpt_by_Par"
    "k_Complete\015\012ORDER BY temp_Sp_Rpt_by_Park_Complete.Unit_Code, temp_Sp_Rpt_b"
    "y_Park_Complete.Plot_ID, temp_Sp_Rpt_by_Park_Complete.Master_Family, temp_Sp_Rpt"
    "_by_Park_Complete.Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5704e070d13aa546877b0c3c6736fcfe
        End
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeccdf5df65103840b0a0cc8b97459ab8
        End
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Master_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcac19f201441ad428ec4328ef0a97878
        End
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x87bf73fbe29b4342975f008ce4035247
        End
    End
    Begin
        dbText "Name" ="SpeciesYears"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x417a0b8935cece499a84514111e92d8e
        End
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.ParkPlotSpecies"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8fb81d054aaac40b5128eaed85879ee
        End
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.ParkPlot"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x389b914c9b5a734a9811510d4651bf50
        End
    End
End
