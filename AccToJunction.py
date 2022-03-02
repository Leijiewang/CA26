# -*- coding: utf-8 -*-
"""
Created on Tue Mar  1 11:22:13 2022

@author: ljwang
"""

from shapely.geometry import Point
import geopandas as gpd
import pandas as pd
process = pd.DataFrame()
#ACC
accident = pd.read_excel(r".\inputdata\109年+110年事故資料(整合).xlsx", dtype = {" GPS經度" : float, " GPS緯度" : float })
accident_geom = [Point(xy) for xy in zip(accident[" GPS經度"],accident[" GPS緯度"])]
crs = {'init': 'epsg:3826'}
gdf_accident = gpd.GeoDataFrame(accident, crs=crs, geometry=accident_geom)
print("肇事資料OK")
#Jnction
jn = pd.read_excel(r".\inputdata\路口定位.xlsx", dtype = {"PositionLon" : float, "PositionLat" : float })
jn_geom = [Point(xy) for xy in zip(jn["PositionLon"],jn["PositionLat"])]
crs = {'init': 'epsg:3826'}
gdf_jn = gpd.GeoDataFrame(jn, crs=crs, geometry=jn_geom)
print("路口資料OK")
for i in range(0,len(gdf_jn)):
    print(gdf_jn["路口地點"][i])
    base = gdf_jn["geometry"][i].buffer(0.0005)
    new = gdf_accident.loc[gdf_accident['geometry'].within(base),:].reset_index(drop=True) #挑出路口半徑500公尺內所有事故
    new["路口地點"] = gdf_jn["路口地點"][i]
    new["路口經度"] = gdf_jn["PositionLon"][i]
    new["路口緯度"] = gdf_jn["PositionLat"][i]
    new = pd.DataFrame(new)
    process = pd.concat([process, new])



# def buffer(data):
#     print(data)
#     base = data["geometry"].buffer(0.0005)  #路口周圍500公尺
#     new = gdf_accident.loc[gdf_accident['geometry'].within(base),:].reset_index(drop=True) #挑出路口半徑500公尺內所有事故
#     new = new.concat([new, new])
#     return new

# junctionac = gdf_jn.apply(lambda x : buffer(x))
process.index = range(len(process))
process = process.loc[:,["總編號(案件編號)", "路口地點", "路口經度", "路口緯度"]]
res = pd.merge(accident, process, on ="總編號(案件編號)", how = "left")

writer = pd.ExcelWriter(r".\outputdata\109年+110年事故資料(整合)+路口.xlsx", engine='xlsxwriter')
res.to_excel(writer, index = False)

writer.book.use_zip64()

writer.save()

#HeatMAP
