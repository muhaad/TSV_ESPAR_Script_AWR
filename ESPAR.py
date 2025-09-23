import pyawr.mwoffice as mwo
import math

# Launch AWR Design Environment
awrde = mwo.CMWOffice()

# Create or access an EM Structure
if not awrde.Project.EM3DStructures.Exists("My3DStructure2"):
    awrde.Project.EM3DStructures.Add("My3DStructure2")

em = awrde.Project.EM3DStructures("My3DStructure2")
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
wafer_t = 500
wafer_r = 2000
metal_t = 10

d1 = 370/2
d2 = 120/2
via_r = 40
g1 = 40/2

parasitic_d = 1737/2
radiator_d = 100

radiator_cap_r = 1006/2
parasitic_cap_r = 120


n = 1        #current solid number (gives unique name to each created solid)

#PyVariable_2 = PyVariableManager_1.NewVariable("v1", "0", "my new variable (v1)", "0, 0, 0, 0", "0", "0", "0", "Global")

script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"wafer_t\", \"({wafer_t}*1e-6)*m\", \"wafer thickness in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"wafer_r\", \"({wafer_r}*1e-6)*m\", \"wafer thickness in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"parasitic_d\", \"({parasitic_d}*1e-6)*m\", \"parasitic element distance from center\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"radiator_d\", \"({radiator_d}*1e-6)*m\", \"radiator element distance from center\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"d1\", \"({d1}*1e-6)*m\", \"center element radius\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"d2\", \"({d2}*1e-6)*m\", \"parasitic element radius\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"via_r\", \"({via_r}*1e-6)*m\", \"via metal thickness\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"g1\", \"({g1}*1e-6)*m\", \"microstrip gap\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"metal_t\", \"({metal_t}*1e-6)*m\", \"metal_thickness\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"radiator_cap_r\", \"({radiator_cap_r}*1e-6)*m\", \"metal_thickness\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"parasitic_cap_r\", \"({parasitic_cap_r}*1e-6)*m\", \"metal_thickness\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"


script = script + f"wafer = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Diel_1\", \"0, 0, 0\", \"0, 0, 1\", \"wafer_t\", \"wafer_r\", True)\n"
script = script + f"metal = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"0, 0, wafer_t\", \"0, 0, 1\", \"metal_t\", \"parasitic_d+d2+2*g1\", True)\n"

script = script + f"metal_blank = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"0, 0, wafer_t\", \"0, 0, 1\", \"metal_t\", \"d1+g1\", True)\n"
script = script + f"PyBoolean_metal = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (metal.Get_ISolid()), (metal_blank.Get_ISolid()), \"True\", True)\n"

script = script + f"base = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"0, 0, wafer_t\", \"0, 0, 1\", \"metal_t\", \"d1\", True)\n"

script = script + f"air_box = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"0, 0, -0.5*wafer_t\", \"0, 0, 1\", \"wafer_t*2\", \"wafer_r*1.2\", True)\n"


for i in range(1, num_parasitic_elements + 1):
    curr_angle = i * 2*math.pi / num_parasitic_elements
    # name   #coord  #material  #center    #axil direction #height #radius
    script = script + f"parasitic_blank_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"wafer_t\", \"d2\", True)\n"
    script = script + f"parasitic_top_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, wafer_t\", \"0, 0, 1\", \"metal_t\", \"d2+g1\", True)\n"
    
    
    script = script + f"pin_diode_{i} = PyGeometry_1.Cylinder(\"pin_{i}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, wafer_t\", \"0, 0, 1\", \"metal_t\", \"d2+g1\", True)\n"
    
    script = script + f"PyVariable_{i} = PyVariableManager_1.NewVariable(\"p{i}\", \"0\", \"p1\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
    script = script + f"PyMaterial_{i} = PyMaterialManager_1.NewMaterial(\"p{i}\")\n"
    script = script + f"PyMaterial_{i}.SetExElectricBulkConductivity(\"((1e+20)*p1+(1e-20))*(S/m)\")\n"
    script = script + f"PyMaterial_{i}.SetExType(\"Electric Conductor (Bulk Conductivity)\")\n"
    script = script + f"PyMaterial_{i}.Get_IAttribute().SetExColor(\"255, 230, 20, 50\")\n"


    script = script + f"parasitic_bottom_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"
    script = script + f"parasitic_cap_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"parasitic_cap_r\", True)\n"
    script = script + f"PyBoolean_bottom_{i} = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (parasitic_cap_{i}.Get_ISolid()), (parasitic_bottom_hole_{i}.Get_ISolid()), \"True\", True)\n"

    script = script + f"parasitic_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"wafer_t\", \"d2\", True)\n"
    script = script + f"via_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n+1}\", \"World\", \"Air\", \"parasitic_d*{math.cos(curr_angle)}, parasitic_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"wafer_t+metal_t\", \"via_r\", True)\n"
    script = script + f"PyBoolean_{i} = PyGeometry_1.Boolean(\"Subtraction_{i}\", \"World\", \"Perfect Conductor\", \"Subtraction\", (parasitic_{i}.Get_ISolid()), (via_{i}.Get_ISolid()), \"True\", True)\n"
    script = script + f"PyBoolean_metal.Get_IBoolean().AddTools((parasitic_top_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
   
    if(i == 1):
        script = script + f"PyBoolean_1 = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Diel_1\", \"Subtraction\", (wafer.Get_ISolid()), (parasitic_blank_1.Get_ISolid()), \"True\", True)\n"
    else:
        script = script + f"PyBoolean_1.Get_IBoolean().AddTools((parasitic_blank_{i}.Get_ISolid()), PyGeometry_1)\n"
    n = n + 2


script = script + f"radiator_cap = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"0, 0, -metal_t\", \"0, 0, 1\", \"metal_t\", \"radiator_cap_r\", True)\n"
# script = script + f"PyBoolean_bottom_{i}r = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (radiator_cap.Get_ISolid()), (radiator_bottom_hole.Get_ISolid()), \"True\", True)\n"

for i in range(1, 4):
    curr_angle = i * 2*math.pi / 3
    # name   #coord  #material  #center    #axil direction #height #radius
    script = script + f"radiator_blank_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"wafer_t\", \"d2\", True)\n"
    script = script + f"radiator_top_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, wafer_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"
    
    script = script + f"radiator_bottom_hole_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Air\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, -metal_t\", \"0, 0, 1\", \"metal_t\", \"via_r\", True)\n"

    script = script + f"radiator_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect Conductor\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"wafer_t\", \"d2\", True)\n"
    script = script + f"via_{i}r = PyGeometry_1.Cylinder(\"Cylinder_{n+1}\", \"World\", \"Air\", \"radiator_d*{math.cos(curr_angle)}, radiator_d*{math.sin(curr_angle)}, 0\", \"0, 0, 1\", \"wafer_t+metal_t\", \"via_r\", True)\n"
    script = script + f"PyBoolean_{i}r = PyGeometry_1.Boolean(\"Subtraction_{i}r\", \"World\", \"Perfect Conductor\", \"Subtraction\", (radiator_{i}.Get_ISolid()), (via_{i}r.Get_ISolid()), \"True\", True)\n"
    
    #script = script + f"PyBoolean_metal.Get_IBoolean().AddTools((radiator_top_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
    if(i == 1):
        script = script + f"PyBoolean_radiator_cap = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (radiator_cap.Get_ISolid()), (radiator_bottom_hole_1.Get_ISolid()), \"True\", True)\n"
        script = script + f"PyBoolean_base = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Perfect Conductor\", \"Subtraction\", (base.Get_ISolid()), (radiator_top_hole_1.Get_ISolid()), \"True\", True)\n"
    else:
        script = script + f"PyBoolean_radiator_cap.Get_IBoolean().AddTools((radiator_bottom_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
        script = script + f"PyBoolean_base.Get_IBoolean().AddTools((radiator_top_hole_{i}.Get_ISolid()), PyGeometry_1)\n"
   
    script = script + f"PyBoolean_1.Get_IBoolean().AddTools((radiator_blank_{i}.Get_ISolid()), PyGeometry_1)\n"
    n = n + 2


script = script + f"port_circle = PyGeometry_1.SurfaceCircle(\"Circle_115\", \"World\", \"Air\", \"XY\", \"0, 0, wafer_t+0.5*metal_t\", \"d1+g1\", True)\n"
script = script + f"port_blank = PyGeometry_1.SurfaceCircle(\"Circle_116\", \"World\", \"Air\", \"XY\", \"0, 0, wafer_t+0.5*metal_t\", \"d1\", True)\n"
script = script + f"PyBoolean_port = PyGeometry_1.Boolean(\"Subtraction_144\", \"World\", \"Air\", \"Subtraction\", (port_circle.Get_ISolid()), (port_blank.Get_ISolid()), \"True\", True)\n"


script = script + f"PyAttributeSet_1 = PyStructure_1.GetAttributeSet(0)\n"
script = script + f"PyRF_Port_1 = PyAttributeSet_1.NewExcitation(\"RF_Port\", \"Port_1\")\n"
script = script + f"PyApplication_1 = PyBoolean_port.Get_ISolid().ApplyAttributeFaces(PyRF_Port_1, (0,), (0,), 0)\n"
script = script + f"PyRF_Port_1.Get_IRFPort().SetExType(\"Lumped\")\n"


#PyFD_Open_1 = PyAttributeSet_1.NewBoundaryCondition("FD_Open", "Open_1")
#PyApplication_2 = PySolid_5.Get_ISolid().ApplyAttributeFaces(PyFD_Open_1, (0, 0, 0, 0), (3, 0, 1, 2), 0)

#print(script)
#input()
external_editor.RunScript(script)

external_editor.Save()
external_editor.Close()
    
em.Update()     #upadtes the 3d EM structure view in AWR

