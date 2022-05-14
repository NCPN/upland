dbMemo "SQL" ="SELECT DISTINCT temp_Sp_Rpt_by_Park_Complete.Unit_Code, temp_Sp_Rpt_by_Park_Comp"
    "lete.Plot_ID, temp_Sp_Rpt_by_Park_Complete.Master_Family, temp_Sp_Rpt_by_Park_Co"
    "mplete.Utah_Species, ConcatRelated(\"Year\",\"temp_Sp_Rpt_by_Park_Complete\",\"P"
    "arkSpecies='\"+ParkSpecies+\"' and ParkPlot='\"+ParkPlot+\"'\",'',\"|\") AS Spec"
    "iesYears\015\012FROM temp_Sp_Rpt_by_Park_Complete\015\012ORDER BY temp_Sp_Rpt_by"
    "_Park_Complete.Unit_Code, temp_Sp_Rpt_by_Park_Complete.Plot_ID, temp_Sp_Rpt_by_P"
    "ark_Complete.Master_Family, temp_Sp_Rpt_by_Park_Complete.Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3ec149db1ab5f74294b955dc5bc3518d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([qry_Sp_Rpt_by_Park_Rollup].[Master_Family]=\"Chenopodiaceae\"))) AND ([qry_S"
    "p_Rpt_by_Park_Rollup].[Unit_Code]=\"DINO\")"
dbMemo "OrderBy" ="[qry_Sp_Rpt_by_Park_Rollup].[Plot_ID], [qry_Sp_Rpt_by_Park_Rollup].[Utah_Species"
    "], [qry_Sp_Rpt_by_Park_Rollup].[Unit_Code], [qry_Sp_Rpt_by_Park_Rollup].[Master_"
    "Family]"
Begin
    Begin
        dbText "Name" ="SpeciesYears"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3924"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0xd7749f0bd0d9ad4a9dfeecf5e32f1f7c
        End
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Sp_Rpt_by_Park_Complete.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2424"
        dbBoolean "ColumnHidden" ="0"
    End
End
