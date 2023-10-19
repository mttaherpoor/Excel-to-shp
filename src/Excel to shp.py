#!/usr/bin/env python
# coding: utf-8

# # In this project, we are converting data from Excel to shapefiles using Python in 6 different mode:
# - Point
# - MultiPoint
# - LineString
# - MultiLineString
# - Polygon
# - MultiPolygon

# ## Import the following libraries.
# Make sure to install the necessary ones :
# 
# - **geopandas** install required.
# 
# - **pandas,shapely,ast,os** already installed!

# In[6]:


import geopandas as gpd
import pandas as pd
import shapely
import ast
import os


# ## Import Excel
# 
# - Then, it reads the address from an Excel file where the addresses are stored in a column named "Address."
# 
# - It selects one of the six modes from the "Mode" column 
# - assigns it to the "coor" column.
#   - If the coordinates are in the format of latitude and longitude, it assigns the value **4326**.
#   - If the coordinates are in the format of UTM Zone, it **adds 23600 to the zone number** and assigns the result to the "coor" column.for example UTM 39 N is 23639

# In[65]:


address=input("insert your excel urls : ")
df=pd.read_excel(address)


# ## Select the directory address.

# In[62]:


input_dir = input("Enter the directory path: ")
if not os.path.exists(input_dir):
    os.makedirs(input_dir)


# ## Excel_to_gdf function
# ### input:
# - DataFrame :Excel file
# - Mode : Choose from this :
#     - Point
#     - MultiPoint
#     - LineString
#     - MultiLineString
#     - Polygon
#     - MultiPolygon
# - Column : The field that contains geometric information
# 
# ### output:
# - GeoDataFrame : Excel file with *geomtry field*

# In[71]:


def Excel_to_gdf(df, mode, column):
    # Convert string representations of lists to actual lists
    df[column] = df[column].apply(ast.literal_eval)
    
    if mode == "MultiPolygon":
        # Create MultiPolygon geometries
        df['geometry'] = df[column].apply(lambda x: shapely.geometry.MultiPolygon([shapely.geometry.Polygon(poly[0]) for poly in x]))
    elif mode == "Polygon":
        # Create Polygon geometries
        df['geometry'] = df[column].apply(lambda x: shapely.geometry.Polygon(x[0]))
    elif mode == "Point":
        # Create Point geometries
        df['geometry'] = df[column].apply(lambda x: shapely.geometry.Point(x))
    elif mode == "MultiPoint":
        # Create MultiPoint geometries
        df['geometry'] = df[column].apply(lambda x: shapely.geometry.MultiPoint(x))
    elif mode == "LineString":
        # Create LineString geometries
        df['geometry'] = df[column].apply(lambda x: shapely.geometry.LineString(x))
    elif mode == "MultiLineString":
        # Create MultiLineString geometries
        df['geometry'] = df[column].apply(lambda x: shapely.geometry.MultiLineString(x))
        
    else:
        raise ValueError("Invalid mode. Supported modes are: MultiPolygon, Polygon, Point, MultiPoint, LineString and MultiLineString.")
    
    df.drop(columns=column, inplace=True)
    gdf = gpd.GeoDataFrame(df, geometry='geometry')
    return gdf


# ## Export Shapefile
# 
# - dfi : Each Data in Excel
# - outputNamei : name of each data
# - crsi : Coordinate System of each data

# In[70]:


for i,row in df.iterrows():
    dfi=pd.read_excel(row["address"])
    mode=row["mode"]
    column=row["column"]
    gdfi=Excel_to_gdf(dfi,mode,column)
    name=row["address"]
    outputNamei=name.split("\\")[-1].split(".")[0]
    crsi = row["Coor"]
    # Save the GeoDataFrame to a shapefile with UTF-8 encoding and UTM 39N CRS
    gdfi.to_file(os.path.join(input_dir, f"{outputNamei}.shp"), encoding='utf-8', driver='ESRI Shapefile', crs=crs)
    


# # Thansks
