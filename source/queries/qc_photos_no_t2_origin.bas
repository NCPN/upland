dbMemo "SQL" ="SELECT qc_plot_visits.Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Da"
    "te, qc_plot_visits.Vegetation_Type\015\012FROM qc_plot_visits LEFT JOIN qc_photo"
    "s_t2_origin ON (qc_plot_visits.Unit_Code = qc_photos_t2_origin.Unit_Code) AND (q"
    "c_plot_visits.Plot_ID = qc_photos_t2_origin.Plot_ID) AND (qc_plot_visits.Start_D"
    "ate = qc_photos_t2_origin.Start_Date) AND (qc_plot_visits.Vegetation_Type = qc_p"
    "hotos_t2_origin.Vegetation_Type)\015\012WHERE (((qc_photos_t2_origin.Unit_Code) "
    "Is Null))\015\012ORDER BY qc_plot_visits.Unit_Code, qc_plot_visits.Plot_ID, qc_p"
    "lot_visits.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
