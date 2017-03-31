dbMemo "SQL" ="SELECT qc_plot_visits.Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Da"
    "te, qc_plot_visits.Vegetation_Type\015\012FROM qc_plot_visits LEFT JOIN qc_photo"
    "s_t1_crust ON (qc_plot_visits.Vegetation_Type = qc_photos_t1_crust.Vegetation_Ty"
    "pe) AND (qc_plot_visits.Start_Date = qc_photos_t1_crust.Start_Date) AND (qc_plot"
    "_visits.Plot_ID = qc_photos_t1_crust.Plot_ID) AND (qc_plot_visits.Unit_Code = qc"
    "_photos_t1_crust.Unit_Code)\015\012WHERE (((qc_plot_visits.Vegetation_Type)=\"gr"
    "assland/shrubland\" \015\012OR (qc_plot_visits.Vegetation_Type)=\"woodland\") \015"
    "\012AND ((qc_photos_t1_crust.Unit_Code) Is Null))\015\012ORDER BY qc_plot_visits"
    ".Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
