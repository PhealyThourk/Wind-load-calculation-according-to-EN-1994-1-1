import numpy as np
from matplotlib import pyplot as plt
import xlsxwriter
import math
from factor import DirFactor, Duopitch_Cpe1_0, Duopitch_Cpe_90, SeasonFactor, OrographyFactor, TurbulenceFactor, ShapeFactor, PcParapet
from factor import Cal_Cpe, Duopitch_Cpe_90, Cal_Rougnesslength, Cal_Cpe_roof, Duopitch_Cpe10_0, Duopitch_Cpe1_0

dir_factor = DirFactor()
season_factor = SeasonFactor()
orography_factor = OrographyFactor()
turbulence_factor = TurbulenceFactor()
pressure_coefficient = PcParapet()
C_dir = dir_factor.c_dir
C_season = season_factor.c_season
C_o = orography_factor.c_o
K_I = turbulence_factor.K_I

pc = pressure_coefficient.p_c

def Run_Analysis_Result(z, Vb0, Terrain_category, Building_width, Building_height, Building_length, Roof_slope, Code_title):
    z_0, z_min = Cal_Rougnesslength(Terrain_category) #Depend on terrain category
    z_0_II, z_min_II = Cal_Rougnesslength("II")  # Example for another terrain category
    z_max = 200  # m
    Density_of_air = 1.25  # kg/m3

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('Wind load results.xlsx')
    worksheet = workbook.add_worksheet()

    # Set header
    worksheet.set_header('&LCalculated by Phealy Thourk&CDate &D&R&G', {'image_right': 'Windload_tallbuilding_logo_small.png'})

    # Set page numbe in footer
    worksheet.set_footer('&CPage &P of &N')

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({"bold": True})



    title_formart = workbook.add_format({
        'font_size': 16,        # Set font size
        'font_color': 'blue',   # Set font color
        'bold': True            # Optional: make text bold
    })

    Chapter_formart = workbook.add_format({
        'font_size': 13,        # Set font size
        'font_color': 'black',   # Set font color
        'bold': True,            # Optional: make text bold
        'left': 5,          # Adds thin border around all sides
        'left_color': 'black',  # Optional: set border color
    })

    Chapter_formart_noborder = workbook.add_format({
        'font_size': 13,        # Set font size
        'font_color': 'black',   # Set font color
        'bold': True,            # Optional: make text bold
    })

    Normal_formart = workbook.add_format({
        'font_size': 11,        # Set font size
        'font_color': 'black',   # Set font color
        'bold': False           # Optional: make text bold
    })

    Ref_formart = workbook.add_format({
        'font_size': 9,        # Set font size
        'font_color': 'black',   # Set font color
        'bold': False ,          # Optional: make text bold
        'left': 5,          # Adds thin border around all sides
        'left_color': 'black',  # Optional: set border color
    })

    border_left_format = workbook.add_format({
        'left': 5,          # Adds thin border around all sides
        'left_color': 'black',  # Optional: set border color
        'font_size': 11,        # Set font size
        'font_color': 'black',   # Set font color
    })

    border_right_format = workbook.add_format({
        'left': 5,          # Adds thin border around all sides
        'left_color': 'black',  # Optional: set border color
        'font_size': 11,        # Set font size
        'font_color': 'black',   # Set font color
    })

    border_lefttop_format = workbook.add_format({
        'left': 5,          # Adds thin border around all sides
        'left_color': 'black',  # Optional: set border color
        'top': 5,          # Adds thin border around all sides
        'top_color': 'black',  # Optional: set border color
        'font_size': 11,        # Set font size
        'font_color': 'black',   # Set font color
    })

    # Apply format to each cell at the left end
    for row in range(5, 300):          # Row indices (0-based, so row 3 = A4)
        worksheet.write(row, 0, '', border_left_format)
        worksheet.write(row-1, 7, '', border_right_format)
        worksheet.write(row-1, 9, '', border_right_format)
        
        
    # Write title
    worksheet.write("A1", "Wind load calculation in accordance with DS/EN 1991-1-4", title_formart)
    worksheet.repeat_rows(0, 0)  # Repeat first row (row 1 in Excel) on every page
    worksheet.repeat_rows(1, 0)  # Repeat first row (row 2 in Excel) on every page
    worksheet.repeat_rows(2, 0)  # Repeat first row (row 2 in Excel) on every page
    worksheet.repeat_rows(3, 0)  # Repeat first row (row 2 in Excel) on every page



    worksheet.write("A5", "Building geometry", Chapter_formart)
    worksheet.write("A7", "Width of the building:", border_left_format)
    worksheet.write("A8", "Height of the building:", border_left_format)
    worksheet.write("A9", "Length of the building:", border_left_format)
    worksheet.write("A10", "Roof slope:", border_left_format)

    worksheet.write("G7", f"{Building_width} m", Normal_formart)
    worksheet.write("G8", f"{Building_height} m", Normal_formart)
    worksheet.write("G9", f"{Building_length} m", Normal_formart)
    worksheet.write("G10", f"{Roof_slope} deg", Normal_formart)

    worksheet.write("H5", "Reference", Chapter_formart )

    # Plot building layout
    Gable_height = Building_height - (0.5*Building_width*np.tan(Roof_slope*np.pi/180))
    X_coord = np.array([0, 0, 0.5*Building_width, Building_width, Building_width, 0])
    Y_coord = np.array([0, Gable_height, Building_height, Gable_height, 0, 0])

    Y2_coord = np.array([0, Gable_height, Building_height, Building_height, Gable_height, 0, 0])
    Z_coord = np.array([0, 0, 0, Building_length, Building_length, Building_length,0])

    plt.figure()
    plt.plot(X_coord, Y_coord, "b-o")
    plt.xlabel("Width [m]")
    plt.ylabel("Height [m]")
    plt.title("Cross section")
    plt.savefig("building layout_Crosssection.png")

    plt.show()

    plt.figure()
    plt.plot(Z_coord, Y2_coord, "b-o")
    plt.xlabel("Length[m]")
    plt.ylabel("Height [m]")
    plt.title("Elevation")
    plt.savefig("building layout_Elevation.png")

    plt.show()

    # Insert figure of building layout
    worksheet.write("A12", "Plot of building layout", border_left_format)
    worksheet.insert_image("B13", "building layout_Crosssection.png", {'x_scale': 0.6, 'y_scale': 0.6})
    worksheet.insert_image("B28", "building layout_Elevation.png", {'x_scale': 0.6, 'y_scale': 0.6})


    # Insert Wind velocity and velocity pressure
    worksheet.write("A43", "Wind velocity and velocity pressure", Chapter_formart)
    worksheet.write("A44", "Fundamental value of the basic wind velocity", border_left_format)
    worksheet.write("F44", f"vb,0 = ", Normal_formart)
    worksheet.write("G44", f"{Vb0} m/s", Normal_formart)

    worksheet.write("A45", "Directional factor", border_left_format)
    worksheet.write("F45", f"c_dir =", Normal_formart)
    worksheet.write("G45", f"{C_dir} ", Normal_formart)

    worksheet.write("A46", "Season factor", border_left_format)
    worksheet.write("F46", f"c_season =", Normal_formart)
    worksheet.write("G46", f"{C_season} ", Normal_formart)

    worksheet.set_column('I:I', 8)   # Column I width = 15
    worksheet.write("H45", Code_title + ", 4.2 (2)P", Ref_formart)
    worksheet.write("H46", Code_title + ", 4.2 (2)P", Ref_formart)

    worksheet.write("A47", "The basic wind velocity", border_left_format)
    worksheet.write("D47", "vb = c_dir * c_season * vb,0 = ", Normal_formart)
    worksheet.write("G47", f"{C_dir*C_season*Vb0} m/s", Normal_formart)
    worksheet.write("H47", Code_title + ", 4.2 (2)P", Ref_formart)

    worksheet.write("A48", "Orography factor", border_left_format)
    worksheet.write("F48", f"co(z)=", Normal_formart)
    worksheet.write("G48", f"{C_o}", Normal_formart)
    worksheet.write("H48", Code_title + ", 4.3.1", Ref_formart)


    worksheet.write("A49", "Terrain category", border_left_format)
    worksheet.write("G49", Terrain_category, Normal_formart)

    worksheet.write("A50", "Reference height ", border_left_format)
    worksheet.write("F50", f"ze = ", Normal_formart)
    worksheet.write("G50", f"{z} m", Normal_formart)

    worksheet.write("A51", "Roughness length", border_left_format)
    worksheet.write("F51", f"z0 = {z_0} m", Normal_formart)
    worksheet.write("H51", Code_title + ", table 4.1", Ref_formart)

    worksheet.write("A52", "Minimum height", border_left_format)
    worksheet.write("F52", f"zmin = {z_min} m", Normal_formart)

    worksheet.write("A53", "Roughness length for terrain category II", border_left_format)
    worksheet.write("F53", f"z0_II = {z_0_II} m", Normal_formart)


    # Calculation process
    Vb = Vb0 * C_dir * C_season
    K_r = 0.19 * pow(z_0 / z_0_II, 0.07)

    worksheet.write("A54", "Terrain factor", border_left_format)
    worksheet.write("C54", f"kr = 0.19 * (z0/z0_II)^0.07 = {round(K_r,2)}", Normal_formart)

    worksheet.write("A55", "Maximum height", border_left_format)
    worksheet.write("C55", f"zmax = {z_max} m", Normal_formart)

    def cr(self):
        if (z_min < self or z_min == self) and (self < z_max or self == z_max):
            return K_r * math.log(self / z_0)
        elif self < z_min:
            return K_r * math.log(z_min / z_0)


    def vm(self):
        return cr(self) * C_o * Vb

    worksheet.write("A56", "Terrain roughness", border_left_format)
    if (z_min < z < z_max) or (z == z_min) or (z == z_max):
        worksheet.write("C56", f"cr(ze) = kr * ln(z/z0) = {round(cr(z),2)} for zmin ≤ ze ≤ zmax", Normal_formart)
    elif (z < z_min) or (z == z_min):
        worksheet.write("C56", f"cr(ze) = cr(zmin) = {round(cr(z),2)} for ze ≤ zmin", Normal_formart)
    worksheet.write("H56", Code_title + ", 4.3.2", Ref_formart)
    worksheet.write("A57", "Mean wind velocity", border_left_format)
    worksheet.write("C57", f"vm = cr(ze) * co(ze) * vb = {round(cr(z),2)} * {round(C_o,2)} * {round(Vb,2)} = {round(vm(z),2)} m/s", Normal_formart)
    worksheet.write("H57", Code_title + ", 4.3.1", Ref_formart)
    worksheet.write("A58", "Turbulence factor", border_left_format)
    worksheet.write("F58", f"KI = {K_I}", Normal_formart)
    worksheet.write("H58", Code_title + ", 4.4", Ref_formart)
    Standard_deviation = K_I * Vb * K_r   # m/s

    worksheet.write("A59", "Standard deviation of wind velocity fluctuations", border_left_format)
    worksheet.write("C60", f"σv = KI * vb * kr = {K_I} * {round(Vb,2)} * {round(K_r,2)} = {round(Standard_deviation,2)} m/s", Normal_formart)
    worksheet.write("H60", Code_title + ", 4.4", Ref_formart)   


    def iv(self):
        if (z_min < self or z_min == self) and (self < z_max or self == z_max):
            return Standard_deviation / vm(self)
        elif self < z_min:
            return Standard_deviation / vm(z_min)

    worksheet.write("A61", "Turbulence intensity", border_left_format)
    if (z_min < z < z_max) or (z == z_min) or (z == z_max):
        worksheet.write("C61", f"Iv = σv/vm = {round(Standard_deviation,2)} / {round(vm(z),2)} = {round(iv(z),2)} for zmin ≤ ze ≤ zmax", Normal_formart)
    elif (z < z_min) or (z == z_min):
        worksheet.write("C61", f"Iv = σv/vm(zmin) = {round(Standard_deviation,2)} / {round(vm(z_min),2)} = {round(iv(z),2)} for ze ≤ zmin", Normal_formart)
    worksheet.write("H61", Code_title + ", 4.4", Ref_formart)

    """Wind peak velocity pressure"""


    def qp(self):
        return (1 + 7 * iv(self)) * 0.5 * Density_of_air * pow(vm(self), 2) * 0.001  # kN/m2




    Vertical_space = 12 # cells
    worksheet.write(f"A{51+Vertical_space}", "Wind peak velocity pressure diagram ", border_left_format)

    """Wind load on free standing wall/parapet"""

    #pressure_on_parapet = max(pc) * qp(z)  # kN/m2

    ref_height = np.linspace(0,100,50)

    qpz_func = np.zeros(50)

    for i in range(0,50):
        qpz_func[i] = qp(ref_height[i])
    # %%

    worksheet.write(f"A{68+Vertical_space}", "Wind peak velocity pressure at ze ", border_left_format)
    worksheet.write(f"E{68+Vertical_space}", f"qp(ze) =", Normal_formart)
    worksheet.write(f"F{68+Vertical_space}", f"{round(qp(z),3)} kN/m2", Normal_formart)
    worksheet.write(f"H{68+Vertical_space}", Code_title + ", 4.5", Ref_formart)

    plt.figure()
    plt.plot(qpz_func, ref_height, "b",  label = "Terrain category " + Terrain_category)
    plt.title("Wind peak velocity pressure")
    plt.ylabel("Reference height z [m]")
    plt.xlabel(r"Wind peak velocity pressure qp(z) [$kN/m^2$]")
    plt.legend()
    plt.grid()

    plt.savefig("Wind peak velocity pressure.png")

    plt.show()

    # Insert figure 
    worksheet.insert_image(f"B{53+Vertical_space}", "Wind peak velocity pressure.png", {'x_scale': 0.6, 'y_scale': 0.6})
    worksheet.write(f"H{53+Vertical_space}", Code_title + ", 4.5", Ref_formart)


    # Wind in longitudinal direction
    d = Building_length
    b = Building_width
    h = Building_height
    worksheet.write(f"A{70+Vertical_space}", "Wind load in longitudinal direction", Chapter_formart)
    worksheet.write(f"A{72+Vertical_space}", "External pressure coefficients on vertical walls", border_left_format)
    worksheet.write(f"A{74+Vertical_space}", f"h = {Building_height} m", border_left_format)
    worksheet.write(f"C{74+Vertical_space}", f"d = {d} m", Normal_formart)
    worksheet.write(f"E{74+Vertical_space}", f"b = {b} m", Normal_formart)

    e_val = min(b, 2*h)

    worksheet.write(f"A{75+Vertical_space}", f"e = b or 2h whichever is smaller = {e_val} m", border_left_format)
    worksheet.write(f"E{75+Vertical_space}", f"h/d = {round(h/d,2)}", Normal_formart)
    worksheet.write(f"A{77+Vertical_space}", "Take into account the lack of correlation of wind pressure", border_left_format)
    worksheet.write(f"A{78+Vertical_space}", "between the windward and leeward side", border_left_format)

    if (h/d > 5) or (h/d == 5):
        rho = 1
        worksheet.write(f"A{79+Vertical_space}", f"h/d >= 5 and rho = {rho}", border_left_format)
    elif (h/d < 1) or (h/d == 1):
        rho = 0.85
        worksheet.write(f"A{79+Vertical_space}", f"h/d <= 1 and rho = {rho}", border_left_format)
    else:
        rho = ((1-0.85)/(5-1))*((h/d)-1) + 0.85
        worksheet.write(f"A{79+Vertical_space}", f"1 < h/d < 5 and rho = {rho}", border_left_format)

    [Cpe10, Cpe1] = ShapeFactor(h,d)

        
    if e_val < d:
        worksheet.write(f"D{79+Vertical_space}", f"e < d", Normal_formart)
        worksheet.write(f"A{82+Vertical_space}", "A", border_left_format)
        worksheet.write(f"A{83+Vertical_space}", "B", border_left_format)
        worksheet.write(f"A{84+Vertical_space}", "C", border_left_format)
        worksheet.write(f"A{85+Vertical_space}", "D", border_left_format)
        worksheet.write(f"A{86+Vertical_space}", "E", border_left_format)
        Area_A = (e_val/5)*h
        Area_B = (4*e_val/5)*h
        Area_C = (d-e_val)*h
        Area_D = b*h - (0.5*b*(0.5*b*np.tan(Roof_slope*np.pi/180)))
        Area_E = Area_D
        
        Area_zone = np.array([Area_A,Area_B,Area_C,Area_D,Area_E])
        
        worksheet.write(f"B{82+Vertical_space}", f"{round(Area_A,2)}", Normal_formart)
        worksheet.write(f"B{83+Vertical_space}", f"{round(Area_B,2)}", Normal_formart)
        worksheet.write(f"B{84+Vertical_space}", f"{round(Area_C,2)}", Normal_formart)
        worksheet.write(f"B{85+Vertical_space}", f"{round(Area_D,2)}", Normal_formart)
        worksheet.write(f"B{86+Vertical_space}", f"{round(Area_E,2)}", Normal_formart)
        
        Cpe = Cal_Cpe(Cpe1,Cpe10,Area_zone)
        
        # External pressure coefficient
        qpzeCpe = Cpe*qp(z)
        qpzeCpe[3] = rho * qpzeCpe[3]
        qpzeCpe[4] = rho * qpzeCpe[4]
        
        for i in range(0,5):
            j = 82+i+Vertical_space
            worksheet.write(f"C{j}", f"{round(Cpe10[i+1][0],2)}",Normal_formart)
            worksheet.write(f"D{j}", f"{round(Cpe1[i],2)}",Normal_formart)
            worksheet.write(f"E{j}", f"{round(Cpe[i],2)}",Normal_formart)
            worksheet.write(f"F{j}", f"{round(qpzeCpe[i],2)}",Normal_formart)
        
        # Insert figure 
        worksheet.insert_image(f"B{96+Vertical_space-8}", "Elevation_vertwall.png", {'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image(f"B{112+Vertical_space-8}", "esmallerthand.png", {'x_scale': 1, 'y_scale': 1})
        
        
    elif (e_val > 5*d) or (e_val == 5*d):
        worksheet.write(f"D{79+Vertical_space}", f"e >= 5d", Normal_formart)

        Area_A = d*h
        Area_D = b*h - (0.5*b*(0.5*b*np.tan(Roof_slope*np.pi/180)))
        Area_E = Area_D

        Area_zone = np.array([Area_A,1,1,Area_D,Area_E])

        Cpe = Cal_Cpe(Cpe1,Cpe10,Area_zone)
        qpzeCpe = Cpe * qp(z)
        qpzeCpe[1] = rho * qpzeCpe[1]
        qpzeCpe[2] = rho * qpzeCpe[2]

        worksheet.write(f"A{82+Vertical_space}", "A", border_left_format)
        worksheet.write(f"A{83+Vertical_space}", "D", border_left_format)
        worksheet.write(f"A{84+Vertical_space}", "E", border_left_format)
        
        
        worksheet.write(f"C{82+Vertical_space}", f"{round(Cpe10[1][0],2)}",Normal_formart)
        worksheet.write(f"D{82+Vertical_space}", f"{round(Cpe1[0],2)}",Normal_formart)
        worksheet.write(f"E{82+Vertical_space}", f"{round(Cpe[0],2)}",Normal_formart)
        
        worksheet.write(f"C{83+Vertical_space}", f"{round(Cpe10[4][0],2)}",Normal_formart)
        worksheet.write(f"D{83+Vertical_space}", f"{round(Cpe1[3],2)}",Normal_formart)
        worksheet.write(f"E{83+Vertical_space}", f"{round(Cpe[3],2)}",Normal_formart)
        
        worksheet.write(f"C{84+Vertical_space}", f"{round(Cpe10[5][0],2)}",Normal_formart)
        worksheet.write(f"D{84+Vertical_space}", f"{round(Cpe1[4],2)}",Normal_formart)
        worksheet.write(f"E{84+Vertical_space}", f"{round(Cpe[4],2)}",Normal_formart)
        
        worksheet.write(f"F{82+Vertical_space}", f"{round(qpzeCpe[0],2)}",Normal_formart)
        worksheet.write(f"F{83+Vertical_space}", f"{round(qpzeCpe[3],2)}",Normal_formart)
        worksheet.write(f"F{84+Vertical_space}", f"{round(qpzeCpe[4],2)}",Normal_formart)
        
        worksheet.insert_image(f"B{96+Vertical_space-8}", "Elevation_vertwall.png", {'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image(f"B{112+Vertical_space-8}", "elargerthand.png", {'x_scale': 1, 'y_scale': 1})
        
    elif (e_val > d) or (e_val == d):
        worksheet.write(f"D{79+Vertical_space}", f"e >= d", Normal_formart)

        
        
        worksheet.write(f"A{82+Vertical_space}", "A", border_left_format)
        worksheet.write(f"A{83+Vertical_space }", "B", border_left_format)
        worksheet.write(f"A{84+Vertical_space}", "D", border_left_format)
        worksheet.write(f"A{85+Vertical_space}", "E", border_left_format)
        Area_A = (e_val/5)*h
        Area_B = (d-e_val*0.2)*h
        Area_D = b*h - (0.5*b*(0.5*b*np.tan(Roof_slope*np.pi/180)))
        Area_E = Area_D

        Area_zone = np.array([Area_A,Area_B,1,Area_D,Area_E])

        Cpe = Cal_Cpe(Cpe1,Cpe10,Area_zone)
        qpzeCpe = Cpe * qp(z)
        qpzeCpe[2] = rho * qpzeCpe[2]   
        qpzeCpe[3] = rho * qpzeCpe[3]
        
        worksheet.write(f"C{82+Vertical_space}", f"{round(Cpe10[1][0],2)}",Normal_formart)
        worksheet.write(f"D{82+Vertical_space}", f"{round(Cpe1[0],2)}",Normal_formart)
        worksheet.write(f"E{82+Vertical_space}", f"{round(Cpe[0],2)}",Normal_formart)
        
        worksheet.write(f"C{83+Vertical_space}", f"{round(Cpe10[2][0],2)}",Normal_formart)
        worksheet.write(f"D{83+Vertical_space}", f"{round(Cpe1[1],2)}",Normal_formart)
        worksheet.write(f"E{83+Vertical_space}", f"{round(Cpe[1],2)}",Normal_formart)
        
        
        worksheet.write(f"C{84+Vertical_space}", f"{round(Cpe10[4][0],2)}",Normal_formart)
        worksheet.write(f"D{84+Vertical_space}", f"{round(Cpe1[3],2)}",Normal_formart)
        worksheet.write(f"E{84+Vertical_space}", f"{round(Cpe[3],2)}",Normal_formart)
        
        worksheet.write(f"C{85+Vertical_space}", f"{round(Cpe10[5][0],2)}",Normal_formart)
        worksheet.write(f"D{85+Vertical_space}", f"{round(Cpe1[4],2)}",Normal_formart)
        worksheet.write(f"E{85+Vertical_space}", f"{round(Cpe[4],2)}",Normal_formart)
        
        worksheet.write(f"F{82+Vertical_space}", f"{round(qpzeCpe[0],2)}",Normal_formart)
        worksheet.write(f"F{83+Vertical_space}", f"{round(qpzeCpe[1],2)}",Normal_formart)
        worksheet.write(f"F{84+Vertical_space}", f"{round(qpzeCpe[3],2)}",Normal_formart)
        worksheet.write(f"F{85+Vertical_space}", f"{round(qpzeCpe[4],2)}",Normal_formart)
        
        worksheet.insert_image(f"B{96+Vertical_space-8}", "Elevation_vertwall.png", {'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image(f"B{112+Vertical_space-8}", "elargerthand.png", {'x_scale': 1, 'y_scale': 1})
        




    worksheet.write(f"A{81+Vertical_space}", "Zone", border_left_format)
    worksheet.write(f"H{74+Vertical_space}", Code_title + ", fig 7.5", Ref_formart)
    worksheet.write(f"H{81+Vertical_space}", Code_title + ", table", Ref_formart)
    worksheet.write(f"H{82+Vertical_space}", "7.1", Ref_formart)
    worksheet.write(f"H{84+Vertical_space}", Code_title + ", fig 7.2", Ref_formart)
    worksheet.write(f"H{85+Vertical_space}", "Lack of correlation is", Ref_formart)
    worksheet.write(f"H{86+Vertical_space}", "considered", Ref_formart)
    worksheet.write(f"H{87+Vertical_space}", "for zone D and E", Ref_formart)


    worksheet.write(f"B{81+Vertical_space}", "Area [m2]", Normal_formart)
    worksheet.write(f"C{81+Vertical_space}", "Cpe10",Normal_formart)
    worksheet.write(f"D{81+Vertical_space}", "Cpe1", Normal_formart)
    worksheet.write(f"E{81+Vertical_space}", "Cpe", Normal_formart)
    worksheet.write(f"F{81+Vertical_space}", "Ext.pressure [kPa]", Normal_formart)


    worksheet.write(f"A{125+Vertical_space-8}", "External pressure coefficients on the roof", border_left_format)
    worksheet.write(f"A{126+Vertical_space-8}", "Wind direction: 90 degree, pitch angle = 20 degree", border_left_format)

    worksheet.write(f"A{128+Vertical_space-8}", "Zone", border_left_format)
    worksheet.write(f"H{130+Vertical_space-8}", Code_title + ", fig 7.8", Ref_formart)
    worksheet.write(f"H{131+Vertical_space-8}", Code_title, Ref_formart)
    worksheet.write(f"H{132+Vertical_space-8}", "table 7.4b", Ref_formart)
    worksheet.write(f"H{145+Vertical_space-8}", Code_title + ", fig 7.8", Ref_formart)

    worksheet.write(f"B{128+Vertical_space-8}", "Area [m2]", Normal_formart)
    worksheet.write(f"C{128+Vertical_space-8}", "Cpe10",Normal_formart)
    worksheet.write(f"D{128+Vertical_space-8}", "Cpe1", Normal_formart)
    worksheet.write(f"E{128+Vertical_space-8}", "Cpe", Normal_formart)
    worksheet.write(f"F{128+Vertical_space-8}", "Ext.pressure [kPa]", Normal_formart)

    worksheet.write(f"A{129+Vertical_space-8}", "F", border_left_format)
    worksheet.write(f"A{130+Vertical_space-8}", "G", border_left_format)
    worksheet.write(f"A{131+Vertical_space-8}", "H", border_left_format)
    worksheet.write(f"A{132+Vertical_space-8}", "I", border_left_format)

    # Calculate area for roof zones
    Area_F = (e_val/10) * (e_val/4) / np.cos(Roof_slope*np.pi/180)
    Area_G = (e_val/10) * (0.5*b-(e_val/4)) / np.cos(Roof_slope*np.pi/180)
    Area_H = ((e_val/2)-(e_val/10)) * (b/2) / np.cos(Roof_slope*np.pi/180)
    Area_I = (d - (e_val/2)) * (b/2) / np.cos(Roof_slope*np.pi/180)

    Area_FGHI = np.array([Area_F, Area_G, Area_H, Area_I])

    worksheet.write(f"B{129+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
    worksheet.write(f"B{130+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)       
    worksheet.write(f"B{131+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
    worksheet.write(f"B{132+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)

    Cpe10_roof_long, Cpe1_roof_long = Duopitch_Cpe_90(Roof_slope)

    worksheet.write(f"C{129+Vertical_space-8}", f"{round(Cpe10_roof_long[0],2)}", Normal_formart)
    worksheet.write(f"C{130+Vertical_space-8}", f"{round(Cpe10_roof_long[1],2)}", Normal_formart)
    worksheet.write(f"C{131+Vertical_space-8}", f"{round(Cpe10_roof_long[2],2)}", Normal_formart)
    worksheet.write(f"C{132+Vertical_space-8}", f"{round(Cpe10_roof_long[3],2)}", Normal_formart)

    worksheet.write(f"D{129+Vertical_space-8}", f"{round(Cpe1_roof_long[0],2)}", Normal_formart)
    worksheet.write(f"D{130+Vertical_space-8}", f"{round(Cpe1_roof_long[1],2)}", Normal_formart)
    worksheet.write(f"D{131+Vertical_space-8}", f"{round(Cpe1_roof_long[2],2)}", Normal_formart)
    worksheet.write(f"D{132+Vertical_space-8}", f"{round(Cpe1_roof_long[3],2)}", Normal_formart)

    Cpe_roof_long = Cal_Cpe_roof(Cpe1_roof_long,Cpe10_roof_long,Area_FGHI)

    worksheet.write(f"E{129+Vertical_space-8}", f"{round(Cpe_roof_long[0],2)}", Normal_formart)
    worksheet.write(f"E{130+Vertical_space-8}", f"{round(Cpe_roof_long[1],2)}", Normal_formart)
    worksheet.write(f"E{131+Vertical_space-8}", f"{round(Cpe_roof_long[2],2)}", Normal_formart)
    worksheet.write(f"E{132+Vertical_space-8}", f"{round(Cpe_roof_long[3],2)}", Normal_formart)

    qpez_roof_long = Cpe_roof_long * qp(z)

    worksheet.write(f"F{129+Vertical_space-8}", f"{round(qpez_roof_long[0],2)}", Normal_formart)
    worksheet.write(f"F{130+Vertical_space-8}", f"{round(qpez_roof_long[1],2)}", Normal_formart)
    worksheet.write(f"F{131+Vertical_space-8}", f"{round(qpez_roof_long[2],2)}", Normal_formart)
    worksheet.write(f"F{132+Vertical_space-8}", f"{round(qpez_roof_long[3],2)}", Normal_formart)

    worksheet.insert_image(f"B{139+Vertical_space-8}", "duopitch_wind90.png", {'x_scale': 0.8, 'y_scale': 0.8})


    worksheet.write(f"A{151+Vertical_space-8}", f"Internal wind pressure coefficient:   cpi = {0.2} and {-0.3}", border_left_format)
    worksheet.write(f"H{151+Vertical_space-8}", Code_title + ", 7.2.9", Ref_formart)
    c_sc_d = 1.0
    worksheet.write(f"A{152+Vertical_space-8}", f"Structural factor cscd = {c_sc_d}", border_left_format)
    worksheet.write(f"H{152+Vertical_space-8}", Code_title + ", 6.2", Ref_formart)

    worksheet.write(f"A{154+Vertical_space-8}", f"Combine external and internal pressures to get net pressure", border_left_format)
    worksheet.write(f"A{156+Vertical_space-8}", f"Case 1: combine with cpi = 0.2", border_left_format)
    worksheet.write(f"E{156+Vertical_space-8}", f"Case 2: combine with cpi = -0.3", Normal_formart)

    worksheet.write(f"A{158+Vertical_space-8}", f"Zone", border_left_format)
    worksheet.write(f"E{158+Vertical_space-8}", f"Zone", Normal_formart)
    worksheet.write(f"B{158+Vertical_space-8}", f"Net pressure [kPa]", Normal_formart)
    worksheet.write(f"F{158+Vertical_space-8}", f"Net pressure [kPa]", Normal_formart)
    if e_val < d:
        worksheet.write(f"A{159+Vertical_space-8}", "A", border_left_format)
        worksheet.write(f"A{160+Vertical_space-8}", "B", border_left_format)
        worksheet.write(f"A{161+Vertical_space-8}", "C", border_left_format)
        worksheet.write(f"A{162+Vertical_space-8}", "D", border_left_format)
        worksheet.write(f"A{163+Vertical_space-8}", "E", border_left_format)
        worksheet.write(f"A{164+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{165+Vertical_space-8}", "G", border_left_format)        
        worksheet.write(f"A{166+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{167+Vertical_space-8}", "I", border_left_format)

        worksheet.write(f"E{159+Vertical_space-8}", "A", Normal_formart)
        worksheet.write(f"E{160+Vertical_space-8}", "B", Normal_formart)
        worksheet.write(f"E{161+Vertical_space-8}", "C", Normal_formart)
        worksheet.write(f"E{162+Vertical_space-8}", "D", Normal_formart)
        worksheet.write(f"E{163+Vertical_space-8}", "E", Normal_formart)
        worksheet.write(f"E{164+Vertical_space-8}", "F", Normal_formart)
        worksheet.write(f"E{165+Vertical_space-8}", "G", Normal_formart)
        worksheet.write(f"E{166+Vertical_space-8}", "H", Normal_formart)
        worksheet.write(f"E{167+Vertical_space-8}", "I", Normal_formart)

        Netpressure_facade_case1 = c_sc_d*qpzeCpe - 0.2 * qp(z)
        Netpressure_roof_case1 = c_sc_d*qpez_roof_long - 0.2 * qp(z)
        Netpressure_facade_case2 = c_sc_d*qpzeCpe - (- 0.3 * qp(z))
        Netpressure_roof_case2 = c_sc_d*qpez_roof_long - (- 0.3 * qp(z))
        worksheet.write(f"B{159+Vertical_space-8}", f"{round(Netpressure_facade_case1[0],2)}", Normal_formart)
        worksheet.write(f"B{160+Vertical_space-8}", f"{round(Netpressure_facade_case1[1],2)}", Normal_formart)
        worksheet.write(f"B{161+Vertical_space-8}", f"{round(Netpressure_facade_case1[2],2)}", Normal_formart)
        worksheet.write(f"B{162+Vertical_space-8}", f"{round(Netpressure_facade_case1[3],2)}", Normal_formart)
        worksheet.write(f"B{163+Vertical_space-8}", f"{round(Netpressure_facade_case1[4],2)}", Normal_formart)
        worksheet.write(f"B{164+Vertical_space-8}", f"{round(Netpressure_roof_case1[0],2)}", Normal_formart)
        worksheet.write(f"B{165+Vertical_space-8}", f"{round(Netpressure_roof_case1[1],2)}", Normal_formart)
        worksheet.write(f"B{166+Vertical_space-8}", f"{round(Netpressure_roof_case1[2],2)}", Normal_formart)
        worksheet.write(f"B{167+Vertical_space-8}", f"{round(Netpressure_roof_case1[3],2)}", Normal_formart)

        worksheet.write(f"F{159+Vertical_space-8}", f"{round(Netpressure_facade_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{160+Vertical_space-8}", f"{round(Netpressure_facade_case2[1],2)}", Normal_formart)
        worksheet.write(f"F{161+Vertical_space-8}", f"{round(Netpressure_facade_case2[2],2)}", Normal_formart)
        worksheet.write(f"F{162+Vertical_space-8}", f"{round(Netpressure_facade_case2[3],2)}", Normal_formart)
        worksheet.write(f"F{163+Vertical_space-8}", f"{round(Netpressure_facade_case2[4],2)}", Normal_formart)  
        worksheet.write(f"F{164+Vertical_space-8}", f"{round(Netpressure_roof_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{165+Vertical_space-8}", f"{round(Netpressure_roof_case2[1],2)}", Normal_formart)
        worksheet.write(f"F{166+Vertical_space-8}", f"{round(Netpressure_roof_case2[2],2)}", Normal_formart)
        worksheet.write(f"F{167+Vertical_space-8}", f"{round(Netpressure_roof_case2[3],2)}", Normal_formart)

    elif e_val >= d and e_val < 5*d:
        worksheet.write(f"A{159+Vertical_space-8}", "A", border_left_format)
        worksheet.write(f"A{160+Vertical_space-8}", "B", border_left_format)
        worksheet.write(f"A{161+Vertical_space-8}", "D", border_left_format)
        worksheet.write(f"A{162+Vertical_space-8}", "E", border_left_format)
        worksheet.write(f"A{163+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{164+Vertical_space-8}", "G", border_left_format)        
        worksheet.write(f"A{165+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{166+Vertical_space-8}", "I", border_left_format)   

        worksheet.write(f"E{159+Vertical_space-8}", "A", Normal_formart)
        worksheet.write(f"E{160+Vertical_space-8}", "B", Normal_formart)
        worksheet.write(f"E{161+Vertical_space-8}", "D", Normal_formart)
        worksheet.write(f"E{162+Vertical_space-8}", "E", Normal_formart)
        worksheet.write(f"E{163+Vertical_space-8}", "F", Normal_formart)
        worksheet.write(f"E{164+Vertical_space-8}", "G", Normal_formart)    
        worksheet.write(f"E{165+Vertical_space-8}", "H", Normal_formart)
        worksheet.write(f"E{166+Vertical_space-8}", "I", Normal_formart)

        Netpressure_facade_case1 = c_sc_d*qpzeCpe - 0.2 * qp(z)
        Netpressure_roof_case1 = c_sc_d*qpez_roof_long - 0.2 * qp(z)
        Netpressure_facade_case2 = c_sc_d*qpzeCpe - (- 0.3 * qp(z))
        Netpressure_roof_case2 = c_sc_d*qpez_roof_long - (- 0.3 * qp(z))
        worksheet.write(f"B{159+Vertical_space-8}", f"{round(Netpressure_facade_case1[0],2)}", Normal_formart)
        worksheet.write(f"B{160+Vertical_space-8}", f"{round(Netpressure_facade_case1[1],2)}", Normal_formart)
        worksheet.write(f"B{161+Vertical_space-8}", f"{round(Netpressure_facade_case1[3],2)}", Normal_formart)
        worksheet.write(f"B{162+Vertical_space-8}", f"{round(Netpressure_facade_case1[4],2)}", Normal_formart)
        worksheet.write(f"B{163+Vertical_space-8}", f"{round(Netpressure_roof_case1[0],2)}", Normal_formart)
        worksheet.write(f"B{164+Vertical_space-8}", f"{round(Netpressure_roof_case1[1],2)}", Normal_formart)
        worksheet.write(f"B{165+Vertical_space-8}", f"{round(Netpressure_roof_case1[2],2)}", Normal_formart)
        worksheet.write(f"B{166+Vertical_space-8}", f"{round(Netpressure_roof_case1[3],2)}", Normal_formart)

        worksheet.write(f"F{159+Vertical_space-8}", f"{round(Netpressure_facade_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{160+Vertical_space-8}", f"{round(Netpressure_facade_case2[1],2)}", Normal_formart)
        worksheet.write(f"F{161+Vertical_space-8}", f"{round(Netpressure_facade_case2[3],2)}", Normal_formart)
        worksheet.write(f"F{162+Vertical_space-8}", f"{round(Netpressure_facade_case2[4],2)}", Normal_formart)
        worksheet.write(f"F{163+Vertical_space-8}", f"{round(Netpressure_roof_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{164+Vertical_space-8}", f"{round(Netpressure_roof_case2[1],2)}", Normal_formart)
        worksheet.write(f"F{165+Vertical_space-8}", f"{round(Netpressure_roof_case2[2],2)}", Normal_formart)
        worksheet.write(f"F{166+Vertical_space-8}", f"{round(Netpressure_roof_case2[3],2)}", Normal_formart)


    elif e_val >= 5*d:  
        worksheet.write(f"A{159+Vertical_space-8}", "A", border_left_format)
        worksheet.write(f"A{160+Vertical_space-8}", "D", border_left_format)
        worksheet.write(f"A{161+Vertical_space-8}", "E", border_left_format)
        worksheet.write(f"A{162+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{163+Vertical_space-8}", "G", border_left_format)        
        worksheet.write(f"A{164+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{165+Vertical_space-8}", "I", border_left_format)

        worksheet.write(f"E{159+Vertical_space-8}", "A", Normal_formart)
        worksheet.write(f"E{160+Vertical_space-8}", "D", Normal_formart)
        worksheet.write(f"E{161+Vertical_space-8}", "E", Normal_formart)
        worksheet.write(f"E{162+Vertical_space-8}", "F", Normal_formart)    
        worksheet.write(f"E{163+Vertical_space-8}", "G", Normal_formart)
        worksheet.write(f"E{164+Vertical_space-8}", "H", Normal_formart)
        worksheet.write(f"E{165+Vertical_space-8}", "I", Normal_formart)

        Netpressure_facade_case1 = c_sc_d*qpzeCpe - 0.2 * qp(z)
        Netpressure_roof_case1 = c_sc_d*qpez_roof_long - 0.2 * qp(z)
        Netpressure_facade_case2 = c_sc_d*qpzeCpe - (- 0.3 * qp(z))
        Netpressure_roof_case2 = c_sc_d*qpez_roof_long - (- 0.3 * qp(z))
        worksheet.write(f"B{159+Vertical_space-8}", f"{round(Netpressure_facade_case1[0],2)}", Normal_formart)
        worksheet.write(f"B{160+Vertical_space-8}", f"{round(Netpressure_facade_case1[3],2)}", Normal_formart)
        worksheet.write(f"B{161+Vertical_space-8}", f"{round(Netpressure_facade_case1[4],2)}", Normal_formart)
        worksheet.write(f"B{162+Vertical_space-8}", f"{round(Netpressure_roof_case1[0],2)}", Normal_formart)
        worksheet.write(f"B{163+Vertical_space-8}", f"{round(Netpressure_roof_case1[1],2)}", Normal_formart)
        worksheet.write(f"B{164+Vertical_space-8}", f"{round(Netpressure_roof_case1[2],2)}", Normal_formart)
        worksheet.write(f"B{165+Vertical_space-8}", f"{round(Netpressure_roof_case1[3],2)}", Normal_formart)

        worksheet.write(f"F{159+Vertical_space-8}", f"{round(Netpressure_facade_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{160+Vertical_space-8}", f"{round(Netpressure_facade_case2[3],2)}", Normal_formart)
        worksheet.write(f"F{161+Vertical_space-8}", f"{round(Netpressure_facade_case2[4],2)}", Normal_formart)
        worksheet.write(f"F{162+Vertical_space-8}", f"{round(Netpressure_roof_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{163+Vertical_space-8}", f"{round(Netpressure_roof_case2[1],2)}", Normal_formart)
        worksheet.write(f"F{164+Vertical_space-8}", f"{round(Netpressure_roof_case2[2],2)}", Normal_formart)
        worksheet.write(f"F{165+Vertical_space-8}", f"{round(Netpressure_roof_case2[3],2)}", Normal_formart)


    # Wind in transverse direction
    d = Building_width
    b = Building_length
    h = Building_height
    worksheet.write(f"A{170+Vertical_space-8}", "Wind load in transverse direction", Chapter_formart)
    worksheet.write(f"A{172+Vertical_space-8}", "External pressure coefficients on vertical walls", border_left_format)
    worksheet.write(f"A{174+Vertical_space-8}", f"h = {Building_height} m", border_left_format)
    worksheet.write(f"C{174+Vertical_space-8}", f"d = {d} m", Normal_formart)
    worksheet.write(f"E{174+Vertical_space-8}", f"b = {b} m", Normal_formart)

    e_val = min(b, 2*h)

    worksheet.write(f"A{175+Vertical_space-8}", f"e = b or 2h whichever is smaller = {e_val} m", border_left_format)
    worksheet.write(f"E{175+Vertical_space-8}", f"h/d = {round(h/d,2)}", Normal_formart)
    worksheet.write(f"A{177+Vertical_space-8}", "Take into account the lack of correlation of wind pressure", border_left_format)
    worksheet.write(f"A{178+Vertical_space-8}", "between the windward and leeward side", border_left_format)

    if (h/d > 5) or (h/d == 5):
        rho = 1
        worksheet.write(f"A{179+Vertical_space-8}", f"h/d >= 5 and rho = {rho}", border_left_format)
    elif (h/d < 1) or (h/d == 1):
        rho = 0.85
        worksheet.write(f"A{179+Vertical_space-8}", f"h/d <= 1 and rho = {rho}", border_left_format)
    else:
        rho = ((1-0.85)/(5-1))*((h/d)-1) + 0.85
        worksheet.write(f"A{179+Vertical_space-8}", f"1 < h/d < 5 and rho = {rho}", border_left_format)

    [Cpe10, Cpe1] = ShapeFactor(h,d)

    if e_val < d:
        worksheet.write(f"D{179+Vertical_space-8}", f"e < d", Normal_formart)
        worksheet.write(f"A{182+Vertical_space-8}", "A", border_left_format)
        worksheet.write(f"A{183+Vertical_space-8}", "B", border_left_format)
        worksheet.write(f"A{184+Vertical_space-8}", "C", border_left_format)
        worksheet.write(f"A{185+Vertical_space-8}", "D", border_left_format)
        worksheet.write(f"A{186+Vertical_space-8}", "E", border_left_format)
        Area_A = (e_val/5)*h
        Area_B = (4*e_val/5)*h
        Area_C = (d-e_val)*h
        Area_D = b*h - (0.5*b*(0.5*b*np.tan(Roof_slope*np.pi/180)))
        Area_E = Area_D
        
        Area_zone = np.array([Area_A,Area_B,Area_C,Area_D,Area_E])
        
        worksheet.write(f"B{182+Vertical_space-8}", f"{round(Area_A,2)}", Normal_formart)
        worksheet.write(f"B{183+Vertical_space-8}", f"{round(Area_B,2)}", Normal_formart)
        worksheet.write(f"B{184+Vertical_space-8}", f"{round(Area_C,2)}", Normal_formart)
        worksheet.write(f"B{185+Vertical_space-8}", f"{round(Area_D,2)}", Normal_formart)
        worksheet.write(f"B{186+Vertical_space-8}", f"{round(Area_E,2)}", Normal_formart)
        
        Cpe = Cal_Cpe(Cpe1,Cpe10,Area_zone)
        
        # External pressure coefficient
        qpzeCpe = Cpe*qp(z)
        qpzeCpe[3] = rho * qpzeCpe[3]
        qpzeCpe[4] = rho * qpzeCpe[4]
        
        for i in range(0,5):
            j = 182+i+Vertical_space-8
            worksheet.write(f"C{j}", f"{round(Cpe10[i+1][0],2)}",Normal_formart)
            worksheet.write(f"D{j}", f"{round(Cpe1[i],2)}",Normal_formart)
            worksheet.write(f"E{j}", f"{round(Cpe[i],2)}",Normal_formart)
            worksheet.write(f"F{j}", f"{round(qpzeCpe[i],2)}",Normal_formart)
        
        # Insert figure 
        worksheet.insert_image(f"B{188+Vertical_space-8}", "Elevation_vertwall.png", {'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image(f"B{204+Vertical_space-8}", "esmallerthand.png", {'x_scale': 1, 'y_scale': 1})
        
        
    elif (e_val > 5*d) or (e_val == 5*d):
        worksheet.write(f"D{179+Vertical_space-8}", f"e >= 5d", Normal_formart)

        Area_A = d*h
        Area_D = b*h - (0.5*b*(0.5*b*np.tan(Roof_slope*np.pi/180)))
        Area_E = Area_D

        Area_zone = np.array([Area_A,1,1,Area_D,Area_E])

        Cpe = Cal_Cpe(Cpe1,Cpe10,Area_zone)
        qpzeCpe = Cpe * qp(z)
        qpzeCpe[1] = rho * qpzeCpe[1]
        qpzeCpe[2] = rho * qpzeCpe[2]

        worksheet.write(f"A{182+Vertical_space-8}", "A", border_left_format)
        worksheet.write(f"A{183+Vertical_space-8}", "D", border_left_format)
        worksheet.write(f"A{184+Vertical_space-8}", "E", border_left_format)
        
        
        worksheet.write(f"C{182+Vertical_space-8}", f"{round(Cpe10[1][0],2)}",Normal_formart)
        worksheet.write(f"D{182+Vertical_space-8}", f"{round(Cpe1[0],2)}",Normal_formart)
        worksheet.write(f"E{182+Vertical_space-8}", f"{round(Cpe[0],2)}",Normal_formart)
        
        worksheet.write(f"C{183+Vertical_space-8}", f"{round(Cpe10[4][0],2)}",Normal_formart)
        worksheet.write(f"D{183+Vertical_space-8}", f"{round(Cpe1[3],2)}",Normal_formart)
        worksheet.write(f"E{183+Vertical_space-8}", f"{round(Cpe[3],2)}",Normal_formart)
        
        worksheet.write(f"C{184+Vertical_space-8}", f"{round(Cpe10[5][0],2)}",Normal_formart)
        worksheet.write(f"D{184+Vertical_space-8}", f"{round(Cpe1[4],2)}",Normal_formart)
        worksheet.write(f"E{184+Vertical_space-8}", f"{round(Cpe[4],2)}",Normal_formart)
        
        worksheet.write(f"F{182+Vertical_space-8}", f"{round(qpzeCpe[0],2)}",Normal_formart)
        worksheet.write(f"F{183+Vertical_space-8}", f"{round(qpzeCpe[3],2)}",Normal_formart)
        worksheet.write(f"F{184+Vertical_space-8}", f"{round(qpzeCpe[4],2)}",Normal_formart)
        
        worksheet.insert_image(f"B{188+Vertical_space-8}", "Elevation_vertwall.png", {'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image(f"B{204+Vertical_space-8}", "elargerthand.png", {'x_scale': 1, 'y_scale': 1})
        
    elif (e_val > d) or (e_val == d):
        worksheet.write(f"D{179+Vertical_space-8}", f"e >= d", Normal_formart)

        
        
        worksheet.write(f"A{182+Vertical_space-8}", "A", border_left_format)
        worksheet.write(f"A{183+Vertical_space-8}", "B", border_left_format)
        worksheet.write(f"A{184+Vertical_space-8}", "D", border_left_format)
        worksheet.write(f"A{185+Vertical_space-8}", "E", border_left_format)
        Area_A = (e_val/5)*h
        Area_B = (d-e_val*0.2)*h
        Area_D = b*h - (0.5*b*(0.5*b*np.tan(Roof_slope*np.pi/180)))
        Area_E = Area_D

        Area_zone = np.array([Area_A,Area_B,1,Area_D,Area_E])

        Cpe = Cal_Cpe(Cpe1,Cpe10,Area_zone)
        qpzeCpe = Cpe * qp(z)
        qpzeCpe[2] = rho * qpzeCpe[2]   
        qpzeCpe[3] = rho * qpzeCpe[3]
        
        worksheet.write(f"C{182+Vertical_space-8}", f"{round(Cpe10[1][0],2)}",Normal_formart)
        worksheet.write(f"D{182+Vertical_space-8}", f"{round(Cpe1[0],2)}",Normal_formart)
        worksheet.write(f"E{182+Vertical_space-8}", f"{round(Cpe[0],2)}",Normal_formart)
        
        worksheet.write(f"C{183+Vertical_space-8}", f"{round(Cpe10[2][0],2)}",Normal_formart)
        worksheet.write(f"D{183+Vertical_space-8}", f"{round(Cpe1[1],2)}",Normal_formart)
        worksheet.write(f"E{183+Vertical_space-8}", f"{round(Cpe[1],2)}",Normal_formart)
        
        
        worksheet.write(f"C{184+Vertical_space-8}", f"{round(Cpe10[4][0],2)}",Normal_formart)
        worksheet.write(f"D{184+Vertical_space-8}", f"{round(Cpe1[3],2)}",Normal_formart)
        worksheet.write(f"E{184+Vertical_space-8}", f"{round(Cpe[3],2)}",Normal_formart)
        
        worksheet.write(f"C{185+Vertical_space-8}", f"{round(Cpe10[5][0],2)}",Normal_formart)
        worksheet.write(f"D{185+Vertical_space-8}", f"{round(Cpe1[4],2)}",Normal_formart)
        worksheet.write(f"E{185+Vertical_space-8}", f"{round(Cpe[4],2)}",Normal_formart)
        
        worksheet.write(f"F{182+Vertical_space-8}", f"{round(qpzeCpe[0],2)}",Normal_formart)
        worksheet.write(f"F{183+Vertical_space-8}", f"{round(qpzeCpe[1],2)}",Normal_formart)
        worksheet.write(f"F{184+Vertical_space-8}", f"{round(qpzeCpe[3],2)}",Normal_formart)
        worksheet.write(f"F{185+Vertical_space-8}", f"{round(qpzeCpe[4],2)}",Normal_formart)
        
        worksheet.insert_image(f"B{188+Vertical_space-8}", "Elevation_vertwall.png", {'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image(f"B{204+Vertical_space-8}", "elargerthand.png", {'x_scale': 1, 'y_scale': 1})
        




    worksheet.write(f"A{181+Vertical_space-8}", "Zone", border_left_format)
    worksheet.write(f"H{174+Vertical_space-8}", Code_title + ", fig 7.5", Ref_formart)
    worksheet.write(f"H{181+Vertical_space-8}", Code_title + ", table", Ref_formart)
    worksheet.write(f"H{182+Vertical_space-8}", "7.1", Ref_formart)
    worksheet.write(f"H{184+Vertical_space-8}", Code_title + ", fig 7.2", Ref_formart)
    worksheet.write(f"H{185+Vertical_space-8}", "Lack of correlation is", Ref_formart)
    worksheet.write(f"H{186+Vertical_space-8}", "considered", Ref_formart)
    worksheet.write(f"H{187+Vertical_space-8}", "for zone D and E", Ref_formart)


    worksheet.write(f"B{181+Vertical_space-8}", "Area [m2]", Normal_formart)
    worksheet.write(f"C{181+Vertical_space-8}", "Cpe10",Normal_formart)
    worksheet.write(f"D{181+Vertical_space-8}", "Cpe1", Normal_formart)
    worksheet.write(f"E{181+Vertical_space-8}", "Cpe", Normal_formart)
    worksheet.write(f"F{181+Vertical_space-8}", "Ext.pressure [kPa]", Normal_formart)

    worksheet.write(f"A{218+Vertical_space-8}", "External pressure coefficients on the roof", border_left_format)
    worksheet.write(f"A{219+Vertical_space-8}", "Wind direction: 0 degree, pitch angle = 20 degree", border_left_format)

    worksheet.write(f"A{221+Vertical_space-8}", "Zone", border_left_format)
    worksheet.write(f"B{221+Vertical_space-8}", "Area [m2]", Normal_formart)
    worksheet.write(f"C{221+Vertical_space-8}", "Cpe10",Normal_formart)
    worksheet.write(f"D{221+Vertical_space-8}", "Cpe1", Normal_formart)
    worksheet.write(f"E{221+Vertical_space-8}", "Cpe", Normal_formart)
    worksheet.write(f"F{221+Vertical_space-8}", "Ext.pressure [kPa]", Normal_formart)

    Area_F = (e_val/4) * (e_val/10)/np.cos(Roof_slope*np.pi/180)
    Area_G = (b - (e_val/2))*(e_val/10)/np.cos(Roof_slope*np.pi/180)
    Area_H = b*(0.5*d - (e_val/10))/np.cos(Roof_slope*np.pi/180)
    Area_I = b*(0.5*d - (e_val/10))/np.cos(Roof_slope*np.pi/180)
    Area_J = 2*Area_F + Area_G  

    Area_FGHIJ = np.array([Area_F, Area_G, Area_H, Area_I, Area_J])

    Cpe10_0_real_case1, Cpe10_0_real_case2, Cpe10_0_real_case3, Cpe10_0_real_case4 = Duopitch_Cpe10_0(Roof_slope)
    Cpe1_0_real_case1, Cpe1_0_real_case2, Cpe1_0_real_case3, Cpe1_0_real_case4 = Duopitch_Cpe1_0(Roof_slope)

    Cpe_roof_short_case1 = Cal_Cpe_roof(Cpe1_0_real_case1,Cpe10_0_real_case1,Area_FGHIJ)
    Cpe_roof_short_case2 = Cal_Cpe_roof(Cpe1_0_real_case2,Cpe10_0_real_case2,Area_FGHIJ)
    Cpe_roof_short_case3 = Cal_Cpe_roof(Cpe1_0_real_case3,Cpe10_0_real_case3,Area_FGHIJ)
    Cpe_roof_short_case4 = Cal_Cpe_roof(Cpe1_0_real_case4,Cpe10_0_real_case4,Area_FGHIJ)

    qpez_roof_short_case1 = Cpe_roof_short_case1 * qp(z)
    qpez_roof_short_case2 = Cpe_roof_short_case2 * qp(z)
    qpez_roof_short_case3 = Cpe_roof_short_case3 * qp(z)
    qpez_roof_short_case4 = Cpe_roof_short_case4 * qp(z)

    if Roof_slope >= -45 and Roof_slope < -5:
        worksheet.write(f"A{222+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{223+Vertical_space-8}", "G", border_left_format)
        worksheet.write(f"A{224+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{225+Vertical_space-8}", "I", border_left_format)
        worksheet.write(f"A{226+Vertical_space-8}", "J", border_left_format)
        if Roof_slope > 0:
            worksheet.insert_image(f"B{228+Vertical_space-8}", "Duopitch_wind0_general_1.png", {'x_scale': 0.8, 'y_scale': 0.8})
        else:
            worksheet.insert_image(f"B{228+Vertical_space-8}", "Duopitch_wind0_general_2.png", {'x_scale': 0.8, 'y_scale': 0.8})
        worksheet.insert_image(f"B{232+6+Vertical_space-8}", "Duopitch_wind0_surface.png", {'x_scale': 0.8, 'y_scale': 0.8})

        worksheet.write(f"B{222+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
        worksheet.write(f"B{223+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)
        worksheet.write(f"B{224+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
        worksheet.write(f"B{225+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)
        worksheet.write(f"B{226+Vertical_space-8}", f"{round(Area_J,2)}", Normal_formart)

        worksheet.write(f"C{222+Vertical_space-8}", f"{round(Cpe10_0_real_case1[0],2)}", Normal_formart)
        worksheet.write(f"C{223+Vertical_space-8}", f"{round(Cpe10_0_real_case1[1],2)}", Normal_formart)
        worksheet.write(f"C{224+Vertical_space-8}", f"{round(Cpe10_0_real_case1[2],2)}", Normal_formart)
        worksheet.write(f"C{225+Vertical_space-8}", f"{round(Cpe10_0_real_case1[3],2)}", Normal_formart)
        worksheet.write(f"C{226+Vertical_space-8}", f"{round(Cpe10_0_real_case1[4],2)}", Normal_formart)

        worksheet.write(f"D{222+Vertical_space-8}", f"{round(Cpe1_0_real_case1[0],2)}", Normal_formart)
        worksheet.write(f"D{223+Vertical_space-8}", f"{round(Cpe1_0_real_case1[1],2)}", Normal_formart)
        worksheet.write(f"D{224+Vertical_space-8}", f"{round(Cpe1_0_real_case1[2],2)}", Normal_formart)
        worksheet.write(f"D{225+Vertical_space-8}", f"{round(Cpe1_0_real_case1[3],2)}", Normal_formart)
        worksheet.write(f"D{226+Vertical_space-8}", f"{round(Cpe1_0_real_case1[4],2)}", Normal_formart)

        worksheet.write(f"E{222+Vertical_space-8}", f"{round(Cpe_roof_short_case1[0],2)}", Normal_formart)
        worksheet.write(f"E{223+Vertical_space-8}", f"{round(Cpe_roof_short_case1[1],2)}", Normal_formart)
        worksheet.write(f"E{224+Vertical_space-8}", f"{round(Cpe_roof_short_case1[2],2)}", Normal_formart)
        worksheet.write(f"E{225+Vertical_space-8}", f"{round(Cpe_roof_short_case1[3],2)}", Normal_formart)
        worksheet.write(f"E{226+Vertical_space-8}", f"{round(Cpe_roof_short_case1[4],2)}", Normal_formart)

        worksheet.write(f"F{222+Vertical_space-8}", f"{round(qpez_roof_short_case1[0],2)}", Normal_formart)
        worksheet.write(f"F{223+Vertical_space-8}", f"{round(qpez_roof_short_case1[1],2)}", Normal_formart)
        worksheet.write(f"F{224+Vertical_space-8}", f"{round(qpez_roof_short_case1[2],2)}", Normal_formart)
        worksheet.write(f"F{225+Vertical_space-8}", f"{round(qpez_roof_short_case1[3],2)}", Normal_formart)
        worksheet.write(f"F{226+Vertical_space-8}", f"{round(qpez_roof_short_case1[4],2)}", Normal_formart)
    
    elif Roof_slope > 45 :
        worksheet.write(f"A{222+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{223+Vertical_space-8}", "G", border_left_format)
        worksheet.write(f"A{224+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{225+Vertical_space-8}", "I", border_left_format)
        worksheet.write(f"A{226+Vertical_space-8}", "J", border_left_format)
        if Roof_slope > 0:
            worksheet.insert_image(f"B{228+Vertical_space-8}", "Duopitch_wind0_general_1.png", {'x_scale': 0.8, 'y_scale': 0.8})
        else:
            worksheet.insert_image(f"B{228+Vertical_space-8}", "Duopitch_wind0_general_2.png", {'x_scale': 0.8, 'y_scale': 0.8})
        worksheet.insert_image(f"B{232+6+Vertical_space-8}", "Duopitch_wind0_surface.png", {'x_scale': 0.8, 'y_scale': 0.8})

        worksheet.write(f"B{222+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
        worksheet.write(f"B{223+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)
        worksheet.write(f"B{224+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
        worksheet.write(f"B{225+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)
        worksheet.write(f"B{226+Vertical_space-8}", f"{round(Area_J,2)}", Normal_formart)

        worksheet.write(f"C{222+Vertical_space-8}", f"{round(Cpe10_0_real_case1[0],2)}", Normal_formart)
        worksheet.write(f"C{223+Vertical_space-8}", f"{round(Cpe10_0_real_case1[1],2)}", Normal_formart)
        worksheet.write(f"C{224+Vertical_space-8}", f"{round(Cpe10_0_real_case1[2],2)}", Normal_formart)
        worksheet.write(f"C{225+Vertical_space-8}", f"{round(Cpe10_0_real_case1[3],2)}", Normal_formart)
        worksheet.write(f"C{226+Vertical_space-8}", f"{round(Cpe10_0_real_case1[4],2)}", Normal_formart)

        worksheet.write(f"D{222+Vertical_space-8}", f"{round(Cpe1_0_real_case1[0],2)}", Normal_formart)
        worksheet.write(f"D{223+Vertical_space-8}", f"{round(Cpe1_0_real_case1[1],2)}", Normal_formart)
        worksheet.write(f"D{224+Vertical_space-8}", f"{round(Cpe1_0_real_case1[2],2)}", Normal_formart)
        worksheet.write(f"D{225+Vertical_space-8}", f"{round(Cpe1_0_real_case1[3],2)}", Normal_formart)
        worksheet.write(f"D{226+Vertical_space-8}", f"{round(Cpe1_0_real_case1[4],2)}", Normal_formart)

        worksheet.write(f"E{222+Vertical_space-8}", f"{round(Cpe_roof_short_case1[0],2)}", Normal_formart)
        worksheet.write(f"E{223+Vertical_space-8}", f"{round(Cpe_roof_short_case1[1],2)}", Normal_formart)
        worksheet.write(f"E{224+Vertical_space-8}", f"{round(Cpe_roof_short_case1[2],2)}", Normal_formart)
        worksheet.write(f"E{225+Vertical_space-8}", f"{round(Cpe_roof_short_case1[3],2)}", Normal_formart)
        worksheet.write(f"E{226+Vertical_space-8}", f"{round(Cpe_roof_short_case1[4],2)}", Normal_formart)

        worksheet.write(f"F{222+Vertical_space-8}", f"{round(qpez_roof_short_case1[0],2)}", Normal_formart)
        worksheet.write(f"F{223+Vertical_space-8}", f"{round(qpez_roof_short_case1[1],2)}", Normal_formart)
        worksheet.write(f"F{224+Vertical_space-8}", f"{round(qpez_roof_short_case1[2],2)}", Normal_formart)
        worksheet.write(f"F{225+Vertical_space-8}", f"{round(qpez_roof_short_case1[3],2)}", Normal_formart)
        worksheet.write(f"F{226+Vertical_space-8}", f"{round(qpez_roof_short_case1[4],2)}", Normal_formart)

    else:
        worksheet.write(f"A{220 + Vertical_space-8}", "Case 1", border_left_format)
        worksheet.write(f"A{222+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{223+Vertical_space-8}", "G", border_left_format)
        worksheet.write(f"A{224+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{225+Vertical_space-8}", "I", border_left_format)
        worksheet.write(f"A{226+Vertical_space-8}", "J", border_left_format)

        worksheet.write(f"B{222+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
        worksheet.write(f"B{223+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)
        worksheet.write(f"B{224+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
        worksheet.write(f"B{225+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)
        worksheet.write(f"B{226+Vertical_space-8}", f"{round(Area_J,2)}", Normal_formart)

        worksheet.write(f"C{222+Vertical_space-8}", f"{round(Cpe10_0_real_case1[0],2)}", Normal_formart)
        worksheet.write(f"C{223+Vertical_space-8}", f"{round(Cpe10_0_real_case1[1],2)}", Normal_formart)
        worksheet.write(f"C{224+Vertical_space-8}", f"{round(Cpe10_0_real_case1[2],2)}", Normal_formart)
        worksheet.write(f"C{225+Vertical_space-8}", f"{round(Cpe10_0_real_case1[3],2)}", Normal_formart)
        worksheet.write(f"C{226+Vertical_space-8}", f"{round(Cpe10_0_real_case1[4],2)}", Normal_formart)

        worksheet.write(f"D{222+Vertical_space-8}", f"{round(Cpe1_0_real_case1[0],2)}", Normal_formart)
        worksheet.write(f"D{223+Vertical_space-8}", f"{round(Cpe1_0_real_case1[1],2)}", Normal_formart)
        worksheet.write(f"D{224+Vertical_space-8}", f"{round(Cpe1_0_real_case1[2],2)}", Normal_formart)
        worksheet.write(f"D{225+Vertical_space-8}", f"{round(Cpe1_0_real_case1[3],2)}", Normal_formart)
        worksheet.write(f"D{226+Vertical_space-8}", f"{round(Cpe1_0_real_case1[4],2)}", Normal_formart)

        worksheet.write(f"E{222+Vertical_space-8}", f"{round(Cpe_roof_short_case1[0],2)}", Normal_formart)
        worksheet.write(f"E{223+Vertical_space-8}", f"{round(Cpe_roof_short_case1[1],2)}", Normal_formart)
        worksheet.write(f"E{224+Vertical_space-8}", f"{round(Cpe_roof_short_case1[2],2)}", Normal_formart)
        worksheet.write(f"E{225+Vertical_space-8}", f"{round(Cpe_roof_short_case1[3],2)}", Normal_formart)
        worksheet.write(f"E{226+Vertical_space-8}", f"{round(Cpe_roof_short_case1[4],2)}", Normal_formart)

        worksheet.write(f"F{222+Vertical_space-8}", f"{round(qpez_roof_short_case1[0],2)}", Normal_formart)
        worksheet.write(f"F{223+Vertical_space-8}", f"{round(qpez_roof_short_case1[1],2)}", Normal_formart)
        worksheet.write(f"F{224+Vertical_space-8}", f"{round(qpez_roof_short_case1[2],2)}", Normal_formart)
        worksheet.write(f"F{225+Vertical_space-8}", f"{round(qpez_roof_short_case1[3],2)}", Normal_formart)
        worksheet.write(f"F{226+Vertical_space-8}", f"{round(qpez_roof_short_case1[4],2)}", Normal_formart)

        # Case 2
        worksheet.write(f"A{228 + Vertical_space-8}", "Case 2", border_left_format)
        worksheet.write(f"A{229+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{230+Vertical_space-8}", "G", border_left_format)
        worksheet.write(f"A{231+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{232+Vertical_space-8}", "I", border_left_format)
        worksheet.write(f"A{233+Vertical_space-8}", "J", border_left_format)

        worksheet.write(f"B{229+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
        worksheet.write(f"B{230+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)
        worksheet.write(f"B{231+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
        worksheet.write(f"B{232+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)
        worksheet.write(f"B{233+Vertical_space-8}", f"{round(Area_J,2)}", Normal_formart)

        worksheet.write(f"C{229+Vertical_space-8}", f"{round(Cpe10_0_real_case2[0],2)}", Normal_formart)
        worksheet.write(f"C{230+Vertical_space-8}", f"{round(Cpe10_0_real_case2[1],2)}", Normal_formart)
        worksheet.write(f"C{231+Vertical_space-8}", f"{round(Cpe10_0_real_case2[2],2)}", Normal_formart)
        worksheet.write(f"C{232+Vertical_space-8}", f"{round(Cpe10_0_real_case2[3],2)}", Normal_formart)
        worksheet.write(f"C{233+Vertical_space-8}", f"{round(Cpe10_0_real_case2[4],2)}", Normal_formart)

        worksheet.write(f"D{229+Vertical_space-8}", f"{round(Cpe1_0_real_case2[0],2)}", Normal_formart)
        worksheet.write(f"D{230+Vertical_space-8}", f"{round(Cpe1_0_real_case2[1],2)}", Normal_formart)
        worksheet.write(f"D{231+Vertical_space-8}", f"{round(Cpe1_0_real_case2[2],2)}", Normal_formart)
        worksheet.write(f"D{232+Vertical_space-8}", f"{round(Cpe1_0_real_case2[3],2)}", Normal_formart)
        worksheet.write(f"D{233+Vertical_space-8}", f"{round(Cpe1_0_real_case2[4],2)}", Normal_formart)

        worksheet.write(f"E{229+Vertical_space-8}", f"{round(Cpe_roof_short_case2[0],2)}", Normal_formart)
        worksheet.write(f"E{230+Vertical_space-8}", f"{round(Cpe_roof_short_case2[1],2)}", Normal_formart)
        worksheet.write(f"E{231+Vertical_space-8}", f"{round(Cpe_roof_short_case2[2],2)}", Normal_formart)
        worksheet.write(f"E{232+Vertical_space-8}", f"{round(Cpe_roof_short_case2[3],2)}", Normal_formart)
        worksheet.write(f"E{233+Vertical_space-8}", f"{round(Cpe_roof_short_case2[4],2)}", Normal_formart)
        
        worksheet.write(f"F{229+Vertical_space-8}", f"{round(qpez_roof_short_case2[0],2)}", Normal_formart)
        worksheet.write(f"F{230+Vertical_space-8}", f"{round(qpez_roof_short_case2[1],2)}", Normal_formart)
        worksheet.write(f"F{231+Vertical_space-8}", f"{round(qpez_roof_short_case2[2],2)}", Normal_formart)
        worksheet.write(f"F{232+Vertical_space-8}", f"{round(qpez_roof_short_case2[3],2)}", Normal_formart)
        worksheet.write(f"F{233+Vertical_space-8}", f"{round(qpez_roof_short_case2[4],2)}", Normal_formart)

        # Case 3

        worksheet.write(f"A{235 + Vertical_space-8}", "Case 3", border_left_format)
        worksheet.write(f"A{236+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{237+Vertical_space-8}", "G", border_left_format)
        worksheet.write(f"A{238+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{239+Vertical_space-8}", "I", border_left_format)
        worksheet.write(f"A{240+Vertical_space-8}", "J", border_left_format)

        worksheet.write(f"B{236+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
        worksheet.write(f"B{237+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)
        worksheet.write(f"B{238+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
        worksheet.write(f"B{239+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)
        worksheet.write(f"B{240+Vertical_space-8}", f"{round(Area_J,2)}", Normal_formart)

        worksheet.write(f"C{236+Vertical_space-8}", f"{round(Cpe10_0_real_case3[0],2)}", Normal_formart)
        worksheet.write(f"C{237+Vertical_space-8}", f"{round(Cpe10_0_real_case3[1],2)}", Normal_formart)
        worksheet.write(f"C{238+Vertical_space-8}", f"{round(Cpe10_0_real_case3[2],2)}", Normal_formart)
        worksheet.write(f"C{239+Vertical_space-8}", f"{round(Cpe10_0_real_case3[3],2)}", Normal_formart)
        worksheet.write(f"C{240+Vertical_space-8}", f"{round(Cpe10_0_real_case3[4],2)}", Normal_formart)

        worksheet.write(f"D{236+Vertical_space-8}", f"{round(Cpe1_0_real_case3[0],2)}", Normal_formart)
        worksheet.write(f"D{237+Vertical_space-8}", f"{round(Cpe1_0_real_case3[1],2)}", Normal_formart)
        worksheet.write(f"D{238+Vertical_space-8}", f"{round(Cpe1_0_real_case3[2],2)}", Normal_formart)
        worksheet.write(f"D{239+Vertical_space-8}", f"{round(Cpe1_0_real_case3[3],2)}", Normal_formart)
        worksheet.write(f"D{240+Vertical_space-8}", f"{round(Cpe1_0_real_case3[4],2)}", Normal_formart)

        worksheet.write(f"E{236+Vertical_space-8}", f"{round(Cpe_roof_short_case3[0],2)}", Normal_formart)
        worksheet.write(f"E{237+Vertical_space-8}", f"{round(Cpe_roof_short_case3[1],2)}", Normal_formart)
        worksheet.write(f"E{238+Vertical_space-8}", f"{round(Cpe_roof_short_case3[2],2)}", Normal_formart)
        worksheet.write(f"E{239+Vertical_space-8}", f"{round(Cpe_roof_short_case3[3],2)}", Normal_formart)
        worksheet.write(f"E{240+Vertical_space-8}", f"{round(Cpe_roof_short_case3[4],2)}", Normal_formart)

        worksheet.write(f"F{236+Vertical_space-8}", f"{round(qpez_roof_short_case3[0],2)}", Normal_formart)
        worksheet.write(f"F{237+Vertical_space-8}", f"{round(qpez_roof_short_case3[1],2)}", Normal_formart)
        worksheet.write(f"F{238+Vertical_space-8}", f"{round(qpez_roof_short_case3[2],2)}", Normal_formart)
        worksheet.write(f"F{239+Vertical_space-8}", f"{round(qpez_roof_short_case3[3],2)}", Normal_formart)
        worksheet.write(f"F{240+Vertical_space-8}", f"{round(qpez_roof_short_case3[4],2)}", Normal_formart)

        # Case 4

        worksheet.write(f"A{242 + Vertical_space-8}", "Case 4", border_left_format)
        worksheet.write(f"A{243+Vertical_space-8}", "F", border_left_format)
        worksheet.write(f"A{244+Vertical_space-8}", "G", border_left_format)
        worksheet.write(f"A{245+Vertical_space-8}", "H", border_left_format)
        worksheet.write(f"A{246+Vertical_space-8}", "I", border_left_format)
        worksheet.write(f"A{247+Vertical_space-8}", "J", border_left_format)

        worksheet.write(f"B{243+Vertical_space-8}", f"{round(Area_F,2)}", Normal_formart)
        worksheet.write(f"B{244+Vertical_space-8}", f"{round(Area_G,2)}", Normal_formart)
        worksheet.write(f"B{245+Vertical_space-8}", f"{round(Area_H,2)}", Normal_formart)
        worksheet.write(f"B{246+Vertical_space-8}", f"{round(Area_I,2)}", Normal_formart)
        worksheet.write(f"B{247+Vertical_space-8}", f"{round(Area_J,2)}", Normal_formart)

        worksheet.write(f"C{243+Vertical_space-8}", f"{round(Cpe10_0_real_case4[0],2)}", Normal_formart)
        worksheet.write(f"C{244+Vertical_space-8}", f"{round(Cpe10_0_real_case4[1],2)}", Normal_formart)
        worksheet.write(f"C{245+Vertical_space-8}", f"{round(Cpe10_0_real_case4[2],2)}", Normal_formart)
        worksheet.write(f"C{246+Vertical_space-8}", f"{round(Cpe10_0_real_case4[3],2)}", Normal_formart)
        worksheet.write(f"C{247+Vertical_space-8}", f"{round(Cpe10_0_real_case4[4],2)}", Normal_formart)

        worksheet.write(f"D{243+Vertical_space-8}", f"{round(Cpe1_0_real_case4[0],2)}", Normal_formart)
        worksheet.write(f"D{244+Vertical_space-8}", f"{round(Cpe1_0_real_case4[1],2)}", Normal_formart)
        worksheet.write(f"D{245+Vertical_space-8}", f"{round(Cpe1_0_real_case4[2],2)}", Normal_formart)
        worksheet.write(f"D{246+Vertical_space-8}", f"{round(Cpe1_0_real_case4[3],2)}", Normal_formart)
        worksheet.write(f"D{247+Vertical_space-8}", f"{round(Cpe1_0_real_case4[4],2)}", Normal_formart)

        worksheet.write(f"E{243+Vertical_space-8}", f"{round(Cpe_roof_short_case4[0],2)}", Normal_formart)
        worksheet.write(f"E{244+Vertical_space-8}", f"{round(Cpe_roof_short_case4[1],2)}", Normal_formart)
        worksheet.write(f"E{245+Vertical_space-8}", f"{round(Cpe_roof_short_case4[2],2)}", Normal_formart)
        worksheet.write(f"E{246+Vertical_space-8}", f"{round(Cpe_roof_short_case4[3],2)}", Normal_formart)
        worksheet.write(f"E{247+Vertical_space-8}", f"{round(Cpe_roof_short_case4[4],2)}", Normal_formart)

        worksheet.write(f"F{243+Vertical_space-8}", f"{round(qpez_roof_short_case4[0],2)}", Normal_formart)
        worksheet.write(f"F{244+Vertical_space-8}", f"{round(qpez_roof_short_case4[1],2)}", Normal_formart)
        worksheet.write(f"F{245+Vertical_space-8}", f"{round(qpez_roof_short_case4[2],2)}", Normal_formart)
        worksheet.write(f"F{246+Vertical_space-8}", f"{round(qpez_roof_short_case4[3],2)}", Normal_formart)
        worksheet.write(f"F{247+Vertical_space-8}", f"{round(qpez_roof_short_case4[4],2)}", Normal_formart)

        if Roof_slope > 0:
            worksheet.insert_image(f"B{249+Vertical_space-8}", "Duopitch_wind0_general_1.png", {'x_scale': 0.8, 'y_scale': 0.8})
        else:
            worksheet.insert_image(f"B{249+Vertical_space-8}", "Duopitch_wind0_general_2.png", {'x_scale': 0.8, 'y_scale': 0.8})
        worksheet.insert_image(f"B{253+6+Vertical_space-8}", "Duopitch_wind0_surface.png", {'x_scale': 0.8, 'y_scale': 0.8})





    #print(table_1)
    #print(table_2)
    #print(f"Wind pressure on parapet: {pressure_on_parapet} kN/m2")

    # Finally, close the Excel file
    # via the close() method.
    workbook.close()


    """print statement"""
    print("Wind load calculation in accordance with DS/EN 1991-1-4")
    print("------------------------------------------------------------------------------")
    print("Building geometry:")
    print(f"Width of the building: {Building_width} m")
    print(f"Height of the building: {Building_height} m")
    print(f"Length of the building: {Building_length} m")
    print(f"Roof slope: {Roof_slope} degree")
    print("Terrain category: " + Terrain_category )
    print(f"Fundamental value of the basic wind velocity: vb,0 = {Vb0} m/s")
    
    print("------------------------------------------------------------------------------")
    print("\nWind load calculation:")
    print("Wind velocity and velocity pressure:")
    print(f"Reference height: ze = {z} m")
    print(f"Terrain factor: kr = {round(K_r,2)}")
    print(f"Basic wind velocity: vb = {round(Vb,2)} m/s")
    print(f"Orography factor: co(z) = {C_o}")
    print(f"Mean wind velocity: vm = {round(vm(z),2)} m/s")
    print(f"Turbulence factor: KI = {K_I}")
    print(f"Wind peak velocity pressure: qpz = {round(qp(z),3)} kN/m2")


    return
