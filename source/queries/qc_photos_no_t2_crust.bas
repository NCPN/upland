﻿dbMemo "SQL" ="SELECT qc_plot_visits.Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Da"
    "te, qc_plot_visits.Vegetation_Type\015\012FROM qc_plot_visits LEFT JOIN qc_photo"
    "s_t2_crust ON (qc_plot_visits.Unit_Code = qc_photos_t2_crust.Unit_Code) AND (qc_"
    "plot_visits.Plot_ID = qc_photos_t2_crust.Plot_ID) AND (qc_plot_visits.Start_Date"
    " = qc_photos_t2_crust.Start_Date) AND (qc_plot_visits.Vegetation_Type = qc_photo"
    "s_t2_crust.Vegetation_Type)\015\012WHERE (((qc_plot_visits.Vegetation_Type)=\"gr"
    "assland/shrubland\" \015\012OR (qc_plot_visits.Vegetation_Type)=\"woodland\") \015"
    "\012AND ((qc_photos_t2_crust.Unit_Code) Is Null))\015\012ORDER BY qc_plot_visits"
    ".Unit_Code, qc_plot_visits.Plot_ID, qc_plot_visits.Start_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
