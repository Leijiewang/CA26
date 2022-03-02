# -*- coding: utf-8 -*-
"""
Created on Tue Mar  1 18:06:42 2022

@author: ljwang
"""

import pandas as pd
import folium
import pandas as pd
import numpy as np
import webbrowser
from folium.plugins import HeatMap
import xlrd
res = pd.read_excel(r".\outputdata\109年+110年事故資料(整合)+路口.xlsx", usecols = ["總編號(案件編號)"," GPS經度", " GPS緯度", "路口地點"])
res = res.loc[res["路口地點"].notnull()]
res = res.drop_duplicates(subset=['總編號(案件編號)'])
res.index = range(len(res))

Lat = res[" GPS緯度"]
Lng = res[" GPS經度"]
LOC=[]
for lng, lat in zip(list(Lng), list(Lat)):
    LOC.append([lat,lng])
    
Center = [np.mean(np.array(Lat, dtype="float32")), np.mean(np.array(Lng, dtype="float32"))]
m = folium.Map(location = Center, zoom_start=15)
HeatMap(LOC).add_to(m)

name = r".\outputdata\map.html"
m.save(name)

data2 = pd.read_excel(r".\inputdata\路口定位.xlsx")
for i in range(len(data2)):
    print(data2["PositionLat"][i])
    m.add_child(folium.Marker(location=[data2["PositionLat"][i],data2["PositionLon"][i]],
                             popup=data2["路口地點"][i]))
name = r".\outputdata\map.html"
m.save(name)