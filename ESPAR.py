import pyawr.mwoffice as mwo
import math

# Launch AWR Design Environment
awrde = mwo.CMWOffice()

# Create or access an EM Structure
if not awrde.Project.EM3DStructures.Exists("Full_System"):
    awrde.Project.EM3DStructures.Add("Full_System")

em = awrde.Project.EM3DStructures("Full_System")
external_editor = em.OpenExternalEditor(mwo.mwExternalEditorType.mwEET_OrionEditor)

script = ""
with open("C:/Users/muhaad/Desktop/Recording.py", 'r') as f:
    #initialization script (read setup from default em enclosure definition)
    script = f.read()
    external_editor.RunScript(script)
    script = ""

def push_script():
    global script
    global external_editor
    external_editor.RunScript(script)
    script = ""

num_parasitic_elements = 7
pcb_t = 1520
pcb_l = 35000
pcb_w = 15000
metal_t = 36

d1 = 500        #not directly d1 (d1+d2)
d2 = 250
via_r = d2 - 18
g1 = 150
fl = 20000     #feedline length
fw = 355.6

parasitic_d = 2000
radiator_d = 0
radiator_cap_r = 450    #lowers resonnant frequency, increases bandwidth, decreases ressonant peak (Lower Q, Lower ressonant freuency)
parasitic_cap_r = 400

##feedline parameters##
gnd_w = 3400
cpw_via_r = 250
via_s = 1016
via_b = 150
cpw_g = 100

##meta parameters##
antenna_r = parasitic_d + d2 + g1*2

n = 1        #current solid number (gives unique name to each created solid)
                                                                                                ##min max inc
#PyVariable_29 = PyVariableManager_1.NewVariable("v1", "0", "my new variable (v1)", "1, 1, 1, 0", "-1", "0", "1", "Global")

#meas_value("Measurement_1",freq_index,'DB|S(1,1)|','0')

script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"pcb_t\", \"({pcb_t}*1e-6)*m\", \"pcb thickness in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"pcb_l\", \"({pcb_l}*1e-6)*m\", \"pcb length in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"pcb_w\", \"({pcb_w}*1e-6)*m\", \"pcb width in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"

script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"parasitic_d\", \"({parasitic_d}*1e-6)*m\", \"parasitic element distance from center\", \"1, 1, 1, 0\", \"(2000*1e-6)*m\", \"(5000*1e-6)*m\", \"(100*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"radiator_d\", \"({radiator_d}*1e-6)*m\", \"radiator element distance from center\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"d1\", \"({d1}*1e-6)*m\", \"center element radius\", \"1, 1, 1, 0\", \"(150*1e-6)*m\", \"(750*1e-6)*m\", \"(50*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"d2\", \"({d2}*1e-6)*m\", \"parasitic element radius\", \"1, 1, 1, 0\", \"(100*1e-6)*m\", \"(250*1e-6)*m\", \"(25*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"via_r\", \"d2-(18*1e-6)*m\", \"via metal thickness\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"g1\", \"({g1}*1e-6)*m\", \"microstrip gap\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"metal_t\", \"({metal_t}*1e-6)*m\", \"metal_thickness\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"radiator_cap_r\", \"({radiator_cap_r}*1e-6)*m\", \"radiator cap radius\", \"1, 1, 1, 0\", \"(150*1e-6)*m\", \"(550*1e-6)*m\", \"(25*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"parasitic_cap_r\", \"({parasitic_cap_r}*1e-6)*m\", \"parasitic cap radius\", \"1, 1, 1, 0\", \"(150*1e-6)*m\", \"(400*1e-6)*m\", \"(25*1e-6)*m\", \"Global\")\n"

##feedline parameters##
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"fl\", \"({fl}*1e-6)*m\", \"feedline length\", \"1, 1, 1, 0\", \"(20000*1e-6)*m\", \"(25000*1e-6)*m\", \"(50*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"fw\", \"({fw}*1e-6)*m\", \"feedline width\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"cpw_via_r\", \"({cpw_via_r}*1e-6)*m\", \"cpw via radius\", \"1, 1, 1, 0\", \"(100*1e-6)*m\", \"(250*1e-6)*m\", \"(25*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"gnd_w\", \"({gnd_w}*1e-6)*m\", \"cpw ground width\", \"1, 1, 1, 0\", \"(3400*1e-6)*m\", \"(4000*1e-6)*m\", \"(50*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"via_s\", \"({via_s}*1e-6)*m\", \"via spacing\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"via_b\", \"({via_b}*1e-6)*m\", \"via horizontal spacing\", \"1, 1, 1, 0\", \"(150*1e-6)*m\", \"(500*1e-6)*m\", \"(25*1e-6)*m\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"cpw_g\", \"({cpw_g}*1e-6)*m\", \"CPW gap\", \"1, 1, 1, 0\", \"(100*1e-6)*m\", \"(300*1e-6)*m\", \"(10*1e-6)*m\", \"Global\")\n"

##meta parameters##
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"antenna_r\", \"parasitic_d+d2+2*g1\", \"total antenna radius\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"

##Create Substrate##
script = script + f"air_box = PyGeometry_1.Box(\"air_box\", \"World\", \"Air\", \"antenna_r+fl, -pcb_w/2, -pcb_t\", \"-pcb_l, pcb_w, pcb_t*3\", True)\n"
script = script + f"pcb = PyGeometry_1.Box(\"pcb\", \"World\", \"FR4\", \"antenna_r+fl, -pcb_w/2, 0\", \"-pcb_l, pcb_w, pcb_t\", True)\n"
script = script + f"metal = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"0, 0, pcb_t\", \"0, 0, 1\", \"metal_t\", \"parasitic_d+d2+2*g1\", True)\n"

##create Antenna Footprint##
script = script + f"metal_blank = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"0, 0, pcb_t\", \"0, 0, 1\", \"metal_t\", \"d1+d2+g1\", True)\n"
script = script + f"PyBoolean_metal = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (metal.Get_ISolid()), (metal_blank.Get_ISolid()), \"True\", True)\n"
script = script + f"base = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"0, 0, pcb_t\", \"0, 0, 1\", \"metal_t\", \"d2+d1\", True)\n"


for i in range(1, num_parasitic_elements + 1):
    curr_angle = (i * 2*math.pi / num_parasitic_elements) + 0.5*(1 * 2*math.pi / num_parasitic_elements)
    # name   #coord  #material  #center    #axil direction #height #radius
    script = script + f"parasitic_blank_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"pcb_t\", \"d2\", True)\n"
    script = script + f"parasitic_top_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, pcb_t\", \"0, 0, 1\", \"metal_t\", \"d2+g1\", True)\n"
    
    
    script = script + f"PyVariable_{i} = PyVariableManager_1.NewVariable(\"p{i}\", \"0\", \"p1\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
    script = script + f"PyMaterial_{i} = PyMaterialManager_1.NewMaterial(\"p{i}\")\n"
    script = script + f"PyMaterial_{i}.SetExElectricBulkConductivity(\"((1e+20)*p{i}+(1e-20))*(S/m)\")\n"
    script = script + f"PyMaterial_{i}.SetExType(\"Electric Conductor (Bulk Conductivity)\")\n"
    script = script + f"PyMaterial_{i}.Get_IAttribute().SetExColor(\"255, 230, 20, 50\")\n"

    ##Set pin or air gap
    #script = script + f"pin_diode_{i} = PyGeometry_1.Cylinder(\"pin_{i}\", \"World\", \"p{i}\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, pcb_t\", \"0, 0, 1\", \"metal_t\", \"d2+g1\", True)\n"
    script = script + f"pin_diode_{i} = PyGeometry_1.Cylinder(\"pin_{i}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, pcb_t\", \"0, 0, 1\", \"metal_t\", \"d2+g1\", True)\n"


    script = script + f"parasitic_bottom_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"
    script = script + f"parasitic_cap_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"d2+parasitic_cap_r\", True)\n"
    script = script + f"PyBoolean_bottom_{i} = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (parasitic_cap_{i}.Get_ISolid()), (parasitic_bottom_hole_{i}.Get_ISolid()), \"True\", True)\n"

    ##parasitic top via cap##
    #script = script + f"parasitic_top_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"
    #script = script + f"parasitic_cap_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"d2+parasitic_cap_r\", True)\n"

    script = script + f"parasitic_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"pcb_t\", \"d2\", True)\n"
    script = script + f"via_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n+1}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"pcb_t+metal_t\", \"via_r\", True)\n"
    script = script + f"PyBoolean_{i} = PyGeometry_1.Boolean(\"Subtraction_{i}\", \"World\", \"Perfect Conductor\", \"Subtraction\", (parasitic_{i}.Get_ISolid()), (via_{i}.Get_ISolid()), \"True\", True)\n"
    script = script + f"PyBoolean_metal.Get_IBoolean().AddTools((parasitic_top_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
   
    if(i == 1):
        script = script + f"PyBoolean_1 = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"FR4\", \"Subtraction\", (pcb.Get_ISolid()), (parasitic_blank_1.Get_ISolid()), \"True\", True)\n"
    else:
        script = script + f"PyBoolean_1.Get_IBoolean().AddTools((parasitic_blank_{i}.Get_ISolid()), PyGeometry_1)\n"
    n = n + 2


script = script + f"radiator_cap = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"0, 0, -metal_t\", \"0, 0, 1\", \"metal_t\", \"radiator_cap_r\", True)\n"
# script = script + f"PyBoolean_bottom_{i}r = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (radiator_cap.Get_ISolid()), (radiator_bottom_hole.Get_ISolid()), \"True\", True)\n"

push_script()

for i in range(1, 2):
    curr_angle = i * 2*math.pi / 3
    # name   #coord  #material  #center    #axil direction #height #radius
    script = script + f"radiator_blank_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"pcb_t\", \"d2\", True)\n"
    script = script + f"radiator_top_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, pcb_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"
    
    script = script + f"radiator_bottom_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"

    script = script + f"radiator_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"pcb_t\", \"d2\", True)\n"
    script = script + f"via_{i}r = PyGeometry_1.Cylinder(\"Cylinder_{n+1}\", \"World\", \"Air\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"pcb_t+metal_t\", \"via_r\", True)\n"
    script = script + f"PyBoolean_{i}r = PyGeometry_1.Boolean(\"Subtraction_{i}r\", \"World\", \"Perfect Conductor\", \"Subtraction\", (radiator_{i}.Get_ISolid()), (via_{i}r.Get_ISolid()), \"True\", True)\n"
    
    #script = script + f"PyBoolean_metal.Get_IBoolean().AddTools((radiator_top_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
    if(i == 1):
        script = script + f"PyBoolean_radiator_cap = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (radiator_cap.Get_ISolid()), (radiator_bottom_hole_1.Get_ISolid()), \"True\", True)\n"
        script = script + f"PyBoolean_base = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (base.Get_ISolid()), (radiator_top_hole_1.Get_ISolid()), \"True\", True)\n"
    else:
        script = script + f"PyBoolean_radiator_cap.Get_IBoolean().AddTools((radiator_bottom_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
        script = script + f"PyBoolean_base.Get_IBoolean().AddTools((radiator_top_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
   

    if(num_parasitic_elements == 0 and i == 1):
        script = script + f"PyBoolean_1 = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"FR4\", \"Subtraction\", (pcb.Get_ISolid()), (radiator_blank_{i}.Get_ISolid()), \"True\", True)\n"
    else:
        script = script + f"PyBoolean_1.Get_IBoolean().AddTools((radiator_blank_{i}.Get_ISolid()), PyGeometry_1)\n"
    n = n + 2

push_script()

##Create CPW##
script = script + f"PyAttributeSet_1 = PyStructure_1.GetAttributeSet(0)\n"
script = script + f"gnd = PyGeometry_1.Box(\"Box_gnd\", \"World\", \"Perfect Conductor\", \"0, -gnd_w/2, pcb_t\", \"antenna_r+fl, gnd_w, metal_t\", True)\n"
script = script + f"CPW = PyGeometry_1.Box(\"Box_125\", \"World\", \"Perfect Conductor\", \"antenna_r, -fw/2, pcb_t\", \"fl, fw, metal_t\", True)\n"
script = script + f"CPW_blank = PyGeometry_1.Box(\"Box_125\", \"World\", \"Air\", \"0, -(cpw_g+fw/2), pcb_t\", \"antenna_r+fl, fw+cpw_g*2, metal_t\", True)\n"
script = script + f"CPW_transistion = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"0, 0, pcb_t\", \"0, 0, 1\", \"metal_t\", \"antenna_r\", True)\n"
script = script + f"PyBoolean_gnd = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (gnd.Get_ISolid()), (CPW_blank.Get_ISolid()), \"True\", True)\n"
script = script + f"PyBoolean_gnd.Get_IBoolean().AddTools((CPW_transistion.Get_ISolid()), PyGeometry_1)\n"
script = script + f"bottom_gnd = PyGeometry_1.Box(\"Box_pcb\", \"World\", \"Perfect Conductor\", \"antenna_r, -pcb_w/2, 0\", \"fl, pcb_w, -metal_t\", True)\n"

##Create Feedline##
script = script + f"feedline = PyGeometry_1.Box(\"Box_125\", \"World\", \"Perfect Conductor\", \"d2-(9*1e-6)*m, -fw/2, pcb_t\", \"antenna_r, fw, metal_t\", True)\n"
script = script + f"feedline_blank = PyGeometry_1.Box(\"Box_125\", \"World\", \"Air\", \"d1, -(cpw_g+fw/2), pcb_t\", \"antenna_r, fw+cpw_g*2, metal_t\", True)\n"
script = script + f"PyBoolean_metal.Get_IBoolean().AddTools((feedline_blank.Get_ISolid()), PyGeometry_1)\n"


# PyRF_Port_1 = PyAttributeSet_1.NewExcitation("RF_Port", "Port_1")
# PyApplication_2 = PySolid_62.Get_ISolid().ApplyAttributeFaces(PyRF_Port_1, (0,), (5,), 0)
# PyRF_Port_2 = PyAttributeSet_1.NewExcitation("RF_Port", "Port_2")
# PyApplication_3 = PyBoolean_20.Get_ISolid().ApplyAttributeFaces(PyRF_Port_2, (0, 0), (9, 2), 0)
# PyRF_Port_2.Get_IRFPort().SetExNumber("-1")


##Create RF Port 1##
script = script + f"PyRF_Port_1 = PyAttributeSet_1.NewExcitation(\"RF_Port\", \"Port_1\")\n"
script = script + f"PyApplication_1 = CPW.Get_ISolid().ApplyAttributeFaces(PyRF_Port_1, (0,), (5,), 0)\n"
script = script + f"PyRF_Port_1.Get_IRFPort().SetExType(\"Wave\")\n"

# ##Create RF Port -1##
script = script + f"PyRF_Port_n1 = PyAttributeSet_1.NewExcitation(\"RF_Port\", \"Port_2\")\n"
script = script + f"PyApplication_1 = PyBoolean_gnd.Get_ISolid().ApplyAttributeFaces(PyRF_Port_n1, (0, 0), (9, 2), 0)\n"
script = script + f"PyRF_Port_n1.Get_IRFPort().SetExType(\"Wave\")\n"
script = script + f"PyRF_Port_n1.Get_IRFPort().SetExNumber(\"-1\")\n"


##create vias##
num_vias = int(fl/via_s)

for i in range(1, num_vias):
    script = script + f" cpw_via_{i}R = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"antenna_r+fl-fl/({num_vias}/{i}), -fw/2-cpw_g-cpw_via_r-via_b, 0\", \"0, 0, 1\", \"pcb_t\", \"cpw_via_r\", True)\n"
    script = script + f" cpw_via_{i}L = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"antenna_r+fl-fl/({num_vias}/{i}), fw/2+cpw_g+cpw_via_r+via_b, 0\", \"0, 0, 1\", \"pcb_t\", \"cpw_via_r\", True)\n"


script = script + f"PyFD_Open_1 = PyAttributeSet_1.NewBoundaryCondition(\"FD_Open\", \"Open_1\")\n"
script = script + f"PyApplication_2 = air_box.Get_ISolid().ApplyAttributeFaces(PyFD_Open_1, (0, 0, 0, 0), (0, 1, 2, 3), 0)\n"

script = script + f"PyAttributeSet_1.Get_IAttributeSet().SetVisibility((), (PyFD_Open_1.Get_IAttribute()))"


external_editor.RunScript(script)

external_editor.Save()
external_editor.Close()
    
em.Update()     #upadtes the 3d EM structure view in AWR

