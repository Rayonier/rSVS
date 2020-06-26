# example2.r

# Postex_016.csv
#Plot,Plot_Radius,Nr,Tree_Spc,Tree_Dia,Tree_Hgt,Tree_PosTex1,Tree_PosTex2,Tree_PosTex3,Tree_Local_x,Tree_Local_y,Tree_Local_Dist,Tree_Local_Angle,Tree_Angle_ToPlotCenter,Latitude,Longitude,Tree_Nr
#16,45,1,1,89,65,1259,1200,1392,3.76,-12.19,12.76,162.84,-3,0.0001102,0.0000338,1
#16,45,2,1,71,60,1146,1112,1296,2.04,-11.61,11.79,170.06,-4,0.0001049,0.0000183,2
ptf = system.file( "extdata", "Postex_016.csv", package = "rSVS" )
svs(ptf)

# other
