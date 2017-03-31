dbMemo "SQL" ="SELECT qc_plot_visits.Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Da"
    "te, qc_plot_visits.Vegetation_Type\015\012FROM qc_plot_visits LEFT JOIN qc_photo"
    "s_t1_end ON (qc_plot_visits.Unit_Code = qc_photos_t1_end.Unit_Code) AND (qc_plot"
    "_visits.Plot_ID = qc_photos_t1_end.Plot_ID) AND (qc_plot_visits.Start_Date = qc_"
    "photos_t1_end.Start_Date) AND (qc_plot_visits.Vegetation_Type = qc_photos_t1_end"
    ".Vegetation_Type)\015\012WHERE (((qc_plot_visits.Vegetation_Type)=\"forest\" Or "
    "(qc_plot_visits.Vegetation_Type)=\"woodland\" Or (qc_plot_visits.Vegetation_Type"
    ")=\"oak scrub\") AND ((qc_photos_t1_end.Unit_Code) Is Null))\015\012ORDER BY qc_"
    "plot_visits.Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
