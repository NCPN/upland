dbMemo "SQL" ="SELECT *\015\012FROM (((tbl_LP_Intercept INNER JOIN tbl_LP_Transect ON tbl_LP_Tr"
    "ansect.Transect_ID = tbl_LP_Intercept.Transect_ID) INNER JOIN tbl_Events ON tbl_"
    "Events.Event_ID = tbl_LP_Transect.Event_ID) INNER JOIN tbl_Locations ON tbl_Loca"
    "tions.Location_ID = tbl_Events.Location_ID) INNER JOIN tlu_NCPN_Plants ON tlu_NC"
    "PN_Plants.Master_PLANT_Code = tbl_LP_Intercept.Top\015\012WHERE ( tbl_Locations."
    "Plot_ID = 248\015\012OR\015\012 tbl_Locations.Plot_ID = 271\015\012OR\015\012 tb"
    "l_Locations.Plot_ID = 282)\015\012AND\015\012 tbl_Locations.Unit_Code = 'CURE'\015"
    "\012AND\015\012tlu_NCPN_Plants.Utah_species IS NULL;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x965ca860f3209244b36bf685e5480737
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.ZION"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb8ed4ca8731d80449b6f25cfc49c502b
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Co_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x17388b5e7f333c40a68c61e023079fcb
        End
    End
    Begin
        dbText "Name" ="l.T1E_Rebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.T1O_UTME"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a336af55a62a043a58d34939e464030
        End
    End
    Begin
        dbText "Name" ="pi.LCA6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Soil_Texture"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Grazed_field"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x898231c3c238a743bd4ffcaeb68fd450
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Unique_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1b318747956cf249b7331ef85b0d0de1
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Veg_Comments"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1ddbff5f4fd174438e9d33902c0e2909
        End
    End
    Begin
        dbText "Name" ="pi.LCS9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf7f0519f98353b478ecf13f89882d7b4
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T3_Elevation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x665d72e6a5ff0c4baa9d6bb56250b591
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeCUD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xce115c46206a6944b73a85c9e66fbfe0
        End
    End
    Begin
        dbText "Name" ="l.T3E_UTME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.10cm_Samples"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCA1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.CARE"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdca071df8e062848b3927c48716f51b0
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T2O_Rebar"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x557553050c85264595fa0b95292294bc
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Bearing_C"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5950077e07a7274aa3938a56990c6b61
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wy_PLANT_code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaa645c712684e24e96e896362a850600
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.ProtectedStatusNotes"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1936f6fd2a41b643913dec863f5af5d4
        End
    End
    Begin
        dbText "Name" ="l.Slope_D"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Site_Selection"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1f8bad05929122428f090343eb353536
        End
    End
    Begin
        dbText "Name" ="l.Rcvr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbefac24339363b4480706bc3d3a04c6b
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Surface"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xffad78e94252b14e9df1aefb7fdfaedc
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1e3bfc0e7e9cca45bfe5d58ac0c79b5c
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6bb416bdf4741439da6a8f3a7b0791c
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x70eed8bcb470124eb42395c014c89abe
        End
    End
    Begin
        dbText "Name" ="l.T1E_UTMN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd899af7f4a47314e888911a181e0fe77
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Other_Eco_Site"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf4f62d9c2f1ce24ab9fe1571a2d4996a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Azimuth"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6812133a4f65e14fa26e1b937a01a280
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS10"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa8cc73e48e50994399a2fb5294a3a2b1
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.D3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1489373509c40f49a723a496d97c4d81
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Transect.Event_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67398ba2cef9f84f9fc3ba3a14ab5be2
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Transect.Recorder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x13922ba2015c0848a560bd680e47af07
        End
    End
    Begin
        dbText "Name" ="tbl_Events.version_key_number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1224a74203062042ae2d3361cf32d2cd
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Fuels_Observer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x51bd098a0b78fd4591950a8a3b39d245
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Census_Recorder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb764ef3e5c102242a54493c992f62e60
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Sapling_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x301381eac2fe2d4b9e6cfd036fb9b1d6
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd92f57cf7e803548ab422707c951d77b
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Soil_Survey_Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x167f3e693ac50f4ea89f0b1f7a845c31
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Elevation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5e81089c926b63479f371eaaf7e7f81a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Coord_System"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeaa380b6fcc5714787dde65e9055a01a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Max_HDOP"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x601f4ace67c92e4f9c15bfff40ad4112
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Hillslope_Position"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x02bdd660da1efa4b874a8473e6355398
        End
    End
    Begin
        dbText "Name" ="pi.D4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope_Shape_Across"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.NABR"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x75a57ec6db266d43be5011fc090c0ddf
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Directions"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6c3f0998aafbf148b6cd3d086b218002
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Trinomial"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb3cda34a4868b40a259016929094bb8
        End
    End
    Begin
        dbText "Name" ="l.Soil_Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Sand_Modifier"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3621de0ef2a1074cb5d7ba3d2665665b
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T3E_Rebar"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x267adb05be3a2748b115e21cfc88f913
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeC"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd3ce0de91104ca4fb0e02df0d4dffd28
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8f40c2888999e44b6a7c53c924af3fd
        End
    End
    Begin
        dbText "Name" ="l.T3O_Rebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeAUD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Max_HDOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Site_Selection_Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Soil_Profile_Samples"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCA7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.FOBU"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xddc89b3ba9797842a753146cefe30a29
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T2O_UTMN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x94d52eaacfb0f64082f51f4d9312e814
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Bearing_B"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x156459a76a3d4345b658fb28258cec8e
        End
    End
    Begin
        dbText "Name" ="l.T2_Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Parent_Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope_C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Other_Percent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Sand_Modifier"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Dominant_Vegetation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf02ff771d8a2424b9e50f5331c575b08
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T1_Elevation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9a7629c09f1c484390746b1843401673
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Vegetation_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xee2eed37f4e34a43833251c825b1a252
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_PLANT_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f8040f58a15f74d9116cf14f3648fb2
        End
    End
    Begin
        dbText "Name" ="l.T1E_UTME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Primary_Eco_Site"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x821a1bcd8cdd364c87f914ec39fb22ec
        End
    End
    Begin
        dbText "Name" ="l.Veg_Assessment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.Top"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCA2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope_Complexity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Rock_Fragment_Qty"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7af98714ad7bdf4b942b2da168289f8f
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.COLM"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x78ba3e9c72d86e46b656db4fd6dffccd
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.HOVE (CO)"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x68cb513a0f168046aa3b9d0e7d6a86f6
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Nativity"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40f7b70d31656c408d0739c6bb694b29
        End
    End
    Begin
        dbText "Name" ="l.Dominant_Vegetation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Soil_Assessment"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcf23bb14e8d8f5438a01dcc93fad7525
        End
    End
    Begin
        dbText "Name" ="pi.LCS5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Geologic_Setting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.T3E_UTMN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb6c2579b3b0aa74d842d8a25be253b2a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeBUD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xceb9f7f44d61c14987af55b898d9e4e9
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wy_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2932bb4cec40c7499ed9732e369e0445
        End
    End
    Begin
        dbText "Name" ="l.T3O_UTMN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Master_Stratification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.2cm_Samples"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaece6e4ad109814992cff2c0142f6d0e
        End
    End
    Begin
        dbText "Name" ="l.Max_PDOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Transect_Length"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.Surface_Alive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Rock_Fragment_Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.DINO (CO)"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xea1215ec74481f49a03580cb188bde07
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.TICA"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfcb4c13da376a74ea6953513f4344391
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T2O_UTME"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67fb859f7596c24ba857dcf771fcdade
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Bearing_A"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf9e58e39e332b45add22032710e269c
        End
    End
    Begin
        dbText "Name" ="l.T2E_Rebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope_B"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Primary_Percent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Soil_Survey_Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.BLCA"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc45f94e60ca2ca4681e6c93189032a5c
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T1E_Rebar"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa908e4f9989abf4a977bccb6c3d6230c
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.WY_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b586e9d29fb594194f070fd38ff8544
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Intercept_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x37c9f1bf645df34582221964698a67f9
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Alive"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1f5c318b82cf2f45b9694a8cf2d96b01
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9c45fc310c4c04ea109b9820a881eb1
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf5fe8aefc509fc4f89223764edf7f78b
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x709d1338f7b84a468aba3abad2a6b2b1
        End
    End
    Begin
        dbText "Name" ="l.T1O_Rebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe08afe3423014a418a4aa6233f5a3fb5
        End
    End
    Begin
        dbText "Name" ="l.Probable_Component"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA9"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb3adfbc807f1f149bc16c883fd4ac08e
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.D2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0b7ac60dc29a764f8cf8190446a2b55b
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Transect.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d24ed6973f57a40b340341ea4522e14
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Transect.Observer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5ab410686fb01d41bbeab7ded02bdff0
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Protocol_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3799305875dded4787ed18587764cca3
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Observer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4e9c3966a35e2a45b84099abc543e041
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Census_Observer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xde5f3d414401004f88b2a30e04540395
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Sapling_Recorder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x51a10fe8772c1345a1141209ecf4a21b
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Meta_MID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd68d0428b6cbd5438e041e72b7585479
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Recorder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4185db901a2680468e806d1758d5be7a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Rcvr_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb868548bb31d2f44ad1ef584d6b3cea5
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Coord_Units"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9a5179595e28f441821f7d065c4427e8
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Max_PDOP"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b0adc0f0766a440b379127becd0bd0c
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Landform"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x55b284c284df8840a1bcea49fb4092d4
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_Shape_Across"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8a95d86e9918d498aef791d5844f12b
        End
    End
    Begin
        dbText "Name" ="pi.LCA8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Add_Synonyms"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x20ae62f5b5f0d14aaf9024ea0ee222f7
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.HOVE (UT)"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdc5a2408b2fca34483b3bbbb88326167
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Duration"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x208ab83edd91e0419c8445a2e055269b
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.10cm_Samples"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x36c17a9b8956924cb75522efa93489dd
        End
    End
    Begin
        dbText "Name" ="pi.Intercept_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.D1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.T3E_UTME"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1a9d7bd2d661e4408ce4efba37472674
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeB"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc445e29848166a47b9217a24687d5a7f
        End
    End
    Begin
        dbText "Name" ="l.Plot_Slope"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.T3O_UTME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeDUD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.Surface"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCA3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.DINO (UT)"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd59358d27d7aec40bba3997bea417ae9
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_D"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3805579893e756468e036e3084a4e173
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.New_Record"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe3ee5650324d404c816679247a373ee2
        End
    End
    Begin
        dbText "Name" ="l.T2E_UTMN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope_A"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Recorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Effervescence"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb6f26d92f4f68d479f6b1ebc94d5603e
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T1E_UTMN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1fcc70c1f9f77747afd2f4cdf6ec6147
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.CO_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5f2398b6e8967048acb9d759c4d45a9e
        End
    End
    Begin
        dbText "Name" ="l.T1O_UTMN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Soil_Comments"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b8bbb4fb47e5c4ab3d958d4495b8fa5
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe0b6167f5b5419479aac428d9d30618f
        End
    End
    Begin
        dbText "Name" ="pi.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x580809816fba1b41bb6f98a67092a6df
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Lifeform"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xed4aee1ed7151044a9e84acbcfcb29e1
        End
    End
    Begin
        dbText "Name" ="l.Plot_Aspect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Site_Selection_Comments"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf44da29d7a705e449d5c2700be69179c
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Soil_Profile_Samples"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x27c1f8d843638c44b8c0ce8066912f7b
        End
    End
    Begin
        dbText "Name" ="l.Effervescence"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.GIS_Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.CANY"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x517116a2eee7f4468a3974f7525259a3
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T3O_Rebar"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x222b16c5d219304fbc0edd2546a16d77
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeAUD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa936365d19e0e542950eb480af7dbebb
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wetland_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8fb8c09ab8dbed4b8204c77925aa9537
        End
    End
    Begin
        dbText "Name" ="l.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Other_Percent"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9c381ce8a96b846818c75750b8c204c
        End
    End
    Begin
        dbText "Name" ="pi.LCA9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Depth"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe5e23d97d4965b419512a6f055a09bf4
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T2_Elevation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0baadf664b78d74dae609e1c77e27cff
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Parent_Material"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9ee3e04b9729714187704ba21fc14dd0
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_C"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xae2b08cd38270b49a6f33dd24ed96bfa
        End
    End
    Begin
        dbText "Name" ="l.T2E_UTME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Bearing_D"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Veg_Assessment"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0bb48c3ee027da41a6adc01723b8cafa
        End
    End
    Begin
        dbText "Name" ="l.Coord_Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x231ecc79a9bf1a4f8842b50cf642c9c5
        End
    End
    Begin
        dbText "Name" ="pi.D3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SiteDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Hillslope_Position"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Taxonomic_Notes"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9d9d73ac4629324cb07134e8b28019d9
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T1E_UTME"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0806c28a397bc24a967a3f3ecda82da6
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.UT_Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfcf619b272d7fe4088c2844bb328908e
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Top"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8ba759ee54cbd64fa98050eec450353f
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd7921da4d70abf44adfdeae9ab4162c1
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x80274ce0a24d1a4fa0cc533d21989f44
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8b635aeb39649c4f85a1dd8eaf48300b
        End
    End
    Begin
        dbText "Name" ="l.T1O_UTME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x07c1c5bf5b4af140845bb851569fbb33
        End
    End
    Begin
        dbText "Name" ="l.Depth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCS9"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4032b08cab56a644bdb35bf6084ee080
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.D1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf3c14b1c01212e4eba55e9da8f29908a
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.D5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x34a1d39db650744aa881330abb6c5a56
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Transect.Visit_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe604037f29a91644b0f3e3c76f01134f
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Location_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6c560cc72e98bd43851ac987570d9512
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Comments"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x131148ebdd86434fa4fec092acbf517a
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Fuels_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x168e1d0379617c4dbb9fe043e60d6e17
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Sapling_Observer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6ca8bdc5c3f2449978dafdae68191b2
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.GIS_Location_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x919857b3e1e2f142941d7ef5a698ad9b
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SiteDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x287d8b513dbf51429ad28fc69414d447
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.GPS_File_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ba7e24da069df43a1e9511d95268715
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x480ee56ddb3a2e4a80ac2f1525f40003
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Datum"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ab06ad0aff4164f9c1790eb7a60331c
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Geologic_Setting"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x65c8be92128ccd4ba2c302fe86c4260d
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_Shape_Down"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8241ad813002ff4895d599266b429b94
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x45636b8eb3142045af7eeb3aea899fdf
        End
    End
    Begin
        dbText "Name" ="pi.LCA4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wetland_Code_Info_Source"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x356f1866806baf41b8d2273a6d1b83d4
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Photosynthetic_Pathway"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x34dbfe47a5a338439d465b0b5bb5e938
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.ProtectedStatusID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4b9252df0eceab44a892e18152bc5cb3
        End
    End
    Begin
        dbText "Name" ="l.Grazed_field"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Transect_Length"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x58a5859e485aba42bb77a9395e67620d
        End
    End
    Begin
        dbText "Name" ="pi.LCS7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.T3O_UTMN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8edeb4562f76141aaee3d460e35c478
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeA"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x944d65da728744428fc949c445e4e89a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Master_Stratification"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6aa73d9478ead448702623f010c17ed
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x500a62ad577124408fb143eb6607135a
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.LU_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb729e652ae4d5a458ef6d51c95f1e3e1
        End
    End
    Begin
        dbText "Name" ="l.T3_Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeCUD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Primary_Percent"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x45f13b49b852e843b7d8393d243457fe
        End
    End
    Begin
        dbText "Name" ="l.Coord_System"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Veg_Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope_Shape_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Rock_Fragment_Size"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc1927af58dc81a498dd69d572139216c
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.ARCH"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0b90459fc0d6714ca1af8543ddec2119
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T2E_Rebar"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x08b8659edc8af8419ea4980817e4f932
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_B"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x62b54e8acb495147bef5b0b1a091759c
        End
    End
    Begin
        dbText "Name" ="l.T2O_Rebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Bearing_C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Probable_Component"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04d6349783cb304db07f545d2555db62
        End
    End
    Begin
        dbText "Name" ="pi.Point"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Landform"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.CEBR"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd66a28e77660e40a85e5d0475ce32f7
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x551f2b03649bd046b9f0b57a8ad323b5
        End
    End
    Begin
        dbText "Name" ="l.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.T1O_Rebar"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa886a9ecab86054ea5ff843fd7e7b3cb
        End
    End
    Begin
        dbText "Name" ="l.Rock_Fragment_Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Site_Selection"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCA10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Co_PLANT_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x124da57ca3f9194b89382e438e08c7d6
        End
    End
    Begin
        dbText "Name" ="l.Plot_Directions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Datum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Other_Eco_Site"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.D5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Soil_Map_Unit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.PISP"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb860120776d6394c8df894480ac84c0c
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T3O_UTME"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3acabb2e00b27a47b4ba6c0689e4bbff
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeDUD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbfebed68707e97459d790f33518a6122
        End
    End
    Begin
        dbText "Name" ="l.T3E_Rebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Observer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9596a6b7430f149a338a41d383ab828
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Slope"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbefd8f8af0cb5f4498852e4c6c3d0eb6
        End
    End
    Begin
        dbText "Name" ="pi.LCA5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.T2E_UTMN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd83269503b34ef488a98e1c502e84764
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_A"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x10a8f08708762440a46da26823ef573b
        End
    End
    Begin
        dbText "Name" ="l.T2O_UTMN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Bearing_B"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Percent_Slope"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2956799807c0344d8d12edab9e9e4687
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.GOSP"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03e6133318ec6741a4f95021262439b3
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Point"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3b9967e648bd3d4d881a50065945ad3d
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.Surface_Alive"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x779dafbabeaf154390a8c30a5ed2fa29
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc7ac3ef9f40104e872f7f79e227625e
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xea878ee349d7c147826a075de8becd80
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaaff449ad62e3440824383b5f406d6b1
        End
    End
    Begin
        dbText "Name" ="l.T1_Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Vegetation_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xce4e86708755014e8b9dd1374a58509a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T1O_UTMN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x51063b33e57188438e39f02a19c49eb6
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.LCA10"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe93ee751a4b29d48a70c11e8138508a2
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Intercept.D4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53d94f64b671174b81cc7e603646e355
        End
    End
    Begin
        dbText "Name" ="tbl_LP_Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x66450f83a9599a43b119a570c287298e
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2835a400adda7e4d9fe541d5c76c415b
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Start_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbaceaa22b0a92c40b0fd9a253e611583
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Fuels_Recorder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbcea6025191b154fa572fe49f55b270e
        End
    End
    Begin
        dbText "Name" ="tbl_Events.Census_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x001b0c0704dd6948aece38c252be4cf9
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdb1cb1df35473b4789f227263f799171
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x82ab911d2351354f99d6844143eeb271
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Soil_Map_Unit"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e60f9db3eb30043b5fa8cbf70ad640e
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4e767030ff404b4da9ebb7ffe5b9d1a7
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.UTM_Zone"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe01b0754325e9b428b3eb575ab23955a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Updated_Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf5a31a2a394edb42a4fddd493cbf5f60
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Slope_Complexity"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x46ee5ed319b0344a9c89eb4d45fd0d3f
        End
    End
    Begin
        dbText "Name" ="l.GPS_File_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.BRCA"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x25b95567ef531643be630767053e1573
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Aspect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xde47f8ef6a0f194f84dd57e39e8f3919
        End
    End
    Begin
        dbText "Name" ="l.UTM_Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Primary_Eco_Site"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.Alive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.LCS3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Percent_Slope"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.CURE"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8461a73604a6474188d100b69add30a7
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.SlopeD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x12a7137cefe5224da07a455aace5d19e
        End
    End
    Begin
        dbText "Name" ="l.T3E_UTMN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SlopeBUD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Soil_Assessment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pi.D2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Soil_Texture"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5be718c7a4d5d54892df51e861a6e21e
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.T2E_UTME"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x525afb9e9e61774fbeeb8aad64635019
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Bearing_D"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x072163b9c715d04aa489d0fea6af3414
        End
    End
    Begin
        dbText "Name" ="l.T2O_UTME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Bearing_A"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.2cm_Samples"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Meta_MID"
        dbLong "AggregateType" ="-1"
    End
End
