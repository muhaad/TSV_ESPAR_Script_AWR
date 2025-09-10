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

num_parasitic_elements = 6
wafer_t = 700
wafer_r = 1000
metal_t = 10


parasitic_r = 5
via_t = 1
spacing = 200


n = 1        #current solid number (gives unique name to each created solid)

#PyVariable_2 = PyVariableManager_1.NewVariable("v1", "0", "my new variable (v1)", "0, 0, 0, 0", "0", "0", "0", "Global")

script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"wafer_t\", \"({wafer_t}*1e-6)*m\", \"wafer thickness in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"
script = script + f"PyVariable_{n} = PyVariableManager_1.NewVariable(\"wafer_r\", \"({wafer_r}*1e-6)*m\", \"wafer thickness in um\", \"0, 0, 0, 0\", \"0\", \"0\", \"0\", \"Global\")\n"


script = script + f"wafer = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Diel_1\", \"0, 0, 0\", \"0, 0, 1\", \"wafer_t\", \"wafer_r\", True)\n"


for i in range(1, num_parasitic_elements + 1):
    curr_angle = i * 2*math.pi / num_parasitic_elements
    # name   #coord  #material  #center    #axil direction #height #radius
    script = script + f"parasitic_blank_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect_Conductor\", \"{round(60 * math.cos(curr_angle), 4)*1e-6}, {round(60 * math.sin(curr_angle), 4)*1e-6}, 0\", \"0, 0, 1\", \"wafer_t\", \"22*1e-6\", True)\n"
    script = script + f"parasitic_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect_Conductor\", \"{round(60 * math.cos(curr_angle), 4)*1e-6}, {round(60 * math.sin(curr_angle), 4)*1e-6}, 0\", \"0, 0, 1\", \"wafer_t\", \"22*1e-6\", True)\n"
    script = script + f"via_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n+1}\", \"World\", \"Air\", \"{round(60 * math.cos(curr_angle), 4)*1e-6}, {round(60*math.sin(curr_angle), 4)*1e-6}, 0\", \"0, 0, 1\", \"wafer_t\", \"12*1e-6\", True)\n"
    script = script + f"PyBoolean_{i} = PyGeometry_1.Boolean(\"Subtraction_{i}\", \"World\", \"Air\", \"Subtraction\", (parasitic_{i}.Get_ISolid()), (via_{i}.Get_ISolid()), \"True\", True)\n"

    if(i == 1):
        script = script + f"PyBoolean_1 = PyGeometry_1.Boolean(\"Subtraction_1\", \"World\", \"Diel_1\", \"Subtraction\", (wafer.Get_ISolid()), (parasitic_blank_1.Get_ISolid()), \"True\", True)\n"
    else:
        script = script + f"PyBoolean_1.Get_IBoolean().AddTools((parasitic_blank_{i}.Get_ISolid()), PyGeometry_1)\n"

    #script = script + f"parasitic_{i} = PyGeometry_1.Cylinder(\"Cylinder_{n}\", \"World\", \"Perfect_Conductor\", \"{round(60 * math.cos(curr_angle), 4)*1e-6}, {round(60 * math.sin(curr_angle), 4)*1e-6}, 0\", \"0, 0, 1\", \"wafer_t\", \"22*1e-6\", True)\n"

    n = n + 2

print(script)
input()
external_editor.RunScript(script)

external_editor.Save()
external_editor.Close()
    
em.Update()     #upadtes the 3d EM structure view in AWR

