"""
Wind load calculation in accordance with DS/EN 1991-1-4

@author: Phealy Thourk
"""

from Analysis_Result import Run_Analysis_Result


"""-----------Input------------------------"""

# Building data
z = 6.8  # m
Vb0 = 36  # m/s
Terrain_category = "III"  # 0, I, II, III, IV

Building_width = 14.4 # m
Building_length = 28.8 # m
Building_height = 6.8 # m (from the ground to the apex)
Roof_slope = 20 # degree

# Code title
Code_title = "DS/EN 1991-1-4"


"""----------------Run analysis and Write results to excel----------------------------------------------"""
Run_Analysis_Result(z, Vb0, Terrain_category, Building_width, Building_height, Building_length, Roof_slope, Code_title)