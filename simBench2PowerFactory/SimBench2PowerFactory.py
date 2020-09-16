"""#################################################################################################
                                        FUNCTIONS                                                        
####################################################################################################"""

# function for opening csv files
def importCSVdata(folder, filename, delim = ";"):
    try:
        with open(os.path.join(folder,filename+".csv"), "r") as csv_file:
            reader = csv.DictReader(csv_file, delimiter= delim)
            return list(reader)
    except IOError:
        app.PrintWarn("File {0}.csv does not appear to exist.".format(filename))

# convert a list that contains PowerFactory objects into a dict
def pfList2Dict(pflist):
    newdict = {}
    for listelement in pflist:
        newdict[listelement.loc_name] = listelement
    return newdict

# function for creating new graphical objects in PowerFactory
def createGraphic(grf_path, grf_name, grf_symbol, dataobject, x, y):
    gridGrf = grf_path
    Grf = gridGrf.CreateObject("IntGrf",grf_name)
    Grf.sSymNam = grf_symbol
    Grf.pDataObj = dataobject
    Grf.rCenterX = x
    Grf.rCenterY = y

def findSlacks(app, nodes, pfNodes):
    slacks = {}
    for row in nodes:
        if (row["vmSetp"] != "NULL" and row["vmSetp"] != "0") and row["vaSetp"] != "NULL" and row["type"] != "auxiliary":
            slackname = row["id"]
            slacks[slackname] = [pfNodes.get(slackname), row["vaSetp"]]
        else:
            continue
    if slacks:
        return slacks
    else:
        app.PrintWarn("No slack-nodes found!")
        return None

"""#################################################################################################
                                        MAIN                                                        
####################################################################################################"""
# ===== Import packages =====
import powerfactory as pf
import os
import csv
from datetime import datetime

#  ===== Import self written modules =====
import PFObjectCreator as pfoc

# ===== Calculate the seconds for setting the studycase time in PF =====
starttime = datetime(1970, 1, 1, 0, 0, 0)
timeprofiletime = datetime(2016, 1, 1, 0, 0, 0)
diff_seconds = (timeprofiletime-starttime).total_seconds()-3600

# ===== Get the current PF-Application =====
app = pf.GetApplication()
if app is None:
    raise Exception("Getting PowerFactory application failed.")
user = app.GetCurrentUser()
project = app.GetActiveProject()
if project is None:
    raise Exception("No project activated. Python Importscript stopped.")
script = app.GetCurrentScript()

#Delete old messages in the PF-output window
app.ClearOutputWindow()
#print to PowerFactory output window
app.PrintInfo("SimBench to PowerFactory converter started...")
app.EchoOff()

# ===== Save relevant global variables =====
studycase = app.GetActiveStudyCase()
ldf = app.GetFromStudyCase('ComLdf') #Calling Loadflow Command object
grid = app.GetCalcRelevantObjects("*.ElmNet",)[1] # grid objects folder
gridGrf = app.GetCurrentDiagram()
feederfolder = app.GetDataFolder("ElmFeeder",1) #Feeder
areafolder = app.GetDataFolder("ElmArea", 1) #Area folder
zonefolder = app.GetDataFolder("ElmZone", 1) #Zone folder
libfolder = app.GetProjectFolder('equip') #Component library
charFolder = app.GetProjectFolder('chars') #Profiles (characteristics) folder
scenfolder = app.GetProjectFolder('scen') #Folder that contains 'scenarios' (study cases)

# Set the datetime in the active studycase
studycase.SetStudyTime(diff_seconds)

# --- Set folder path containing SimBench csv-files to import
thisScript = app.GetCurrentScript()     # Get this script
folderpath = thisScript.folder          # folder defined as input parameter in the PowerFactory Python Object is the folderpath
if folderpath is False:
    app.PrintError('The chosen folder path "{0}" does not appear to exist!'.format(folderpath))

# --- Set up loadflow options
ldf.iopt_plim = 1

"""
Read the csv data and save it temporarily in variables for further steps
"""
app.PrintPlain("=======START CHECKING THE CSV-FILES=======")
coordinates = importCSVdata(folderpath, "Coordinates")
substations = importCSVdata(folderpath, "Substation")
nodes = importCSVdata(folderpath, "Node")
switches = importCSVdata(folderpath, "Switch")
linetypes = importCSVdata(folderpath, "LineType")
dclinetypes = importCSVdata(folderpath, "DCLineType")
lines = importCSVdata(folderpath, "Line")
transformertypes = importCSVdata(folderpath, "Transformertype")
transformers = importCSVdata(folderpath, "Transformer")
# transformertypes3w = importCSVdata(folderpath, "Transformer3Wtype")
# transroformers3w = importCSVdata(folderpath, "Transformer3W")
shunts = importCSVdata(folderpath, "Shunt")
xnets = importCSVdata(folderpath, "ExternalNet")
loads = importCSVdata(folderpath, "Load")
powerplants = importCSVdata(folderpath, "PowerPlant")
# Check if it is a EHV model, therefore check if data contains powerplants in order to activate time profiles (powerplant are only contained in EHV models)
activate_timeprofile = 1 # 1 = time profiles are initially out of service/deactivated and must be activated manually in the model if needed
if powerplants:
    activate_timeprofile = 0
    # Set loadflow calculation settings to get the same results as in Integral and pandapower
    ldf.iopt_plim = 0
    ldf.iPbalancing = 4
reses = importCSVdata(folderpath, "RES")
storages = importCSVdata(folderpath, "Storage")
measurements = importCSVdata(folderpath, "Measurement")
studycases = importCSVdata(folderpath, "StudyCases")

if switches:
    no_sw = False
else:
    no_sw = True
app.PrintPlain("=======FINISHED CHECKING THE CSV-FILES======="+"\n")

"""
Set needed global variables for import to PowerFactory
"""
area_set = set()
zone_set = set()
feeder_set = set()
auxnodes = {}   # variable for saving auxiliary nodes

"""
Save relevant PowerFactory objects of the current application to avoid creating them again
"""
#Areas
areaPFO_list = areafolder.GetContents("*.ElmArea")
pfAreas = pfList2Dict(areaPFO_list)
#Zones
zonePFO_list = zonefolder.GetContents("*.ElmZone")
pfZones = pfList2Dict(zonePFO_list)
#Substations
substationPFO_list = app.GetCalcRelevantObjects("*.ElmSubstat")
pfSubstations = pfList2Dict(substationPFO_list)
#Nodes
nodePFO_list = app.GetCalcRelevantObjects("*.ElmTerm")
pfNodes = pfList2Dict(nodePFO_list)
#Auxnodes
pfAuxnodes = {}
for node in nodePFO_list:
    if node.iUsage == 2:
        pfAuxnodes[node.loc_name] = node
#Couplers
coupPFO_list = app.GetCalcRelevantObjects("*.ElmCoup")
pfCouplers = pfList2Dict(coupPFO_list)
#Switches
switchPFO_list = app.GetCalcRelevantObjects("*.StaSwitch")
pfSwitches = pfList2Dict(switchPFO_list)
#Cubicles
cubiclePFO_list = app.GetCalcRelevantObjects("*.StaCubicle")
pfCubicles = pfList2Dict(cubiclePFO_list)
#LineTypes
linetypePFO_list = libfolder.GetContents("*.TypLne")
pfLineTypes = pfList2Dict(linetypePFO_list)
#Lines
linePFO_list = app.GetCalcRelevantObjects("*.ElmLne")
pfLines = pfList2Dict(linePFO_list)
#TranformerTypes
transformertypePFO_list = libfolder.GetContents("*.TypTr2")
pfTransformerTypes = pfList2Dict(transformertypePFO_list)
#Transformers
transformerPFO_list = app.GetCalcRelevantObjects("*.ElmTr2")
pfTransformers = pfList2Dict(transformerPFO_list)
#External nets
xnetsPFO_list = app.GetCalcRelevantObjects("*.ElmXnet")
pfXnets = pfList2Dict(xnetsPFO_list)
#LoadTypes
loadtypePFO_list = app.GetCalcRelevantObjects("*.TypLod")
pfLoadTypes = pfList2Dict(loadtypePFO_list)
#Loads
loadPFO_list = app.GetCalcRelevantObjects("*.ElmLod")
pfLoads = pfList2Dict(loadPFO_list)
#RES
resPFO_list = app.GetCalcRelevantObjects("*.ElmGenStat")
pfRES = pfList2Dict(resPFO_list)
#PowerPlants
ppPFO_list = app.GetCalcRelevantObjects("*.ElmSym")
pfPP = pfList2Dict(ppPFO_list)
#PowerPlants
shuntsPFO_list = app.GetCalcRelevantObjects("*.ElmShnt")
pfShunts = pfList2Dict(shuntsPFO_list)
#Profiles
profilePFO_list = charFolder.GetContents("*.Chatime")
pfProfiles = pfList2Dict(profilePFO_list)
#Measurements
measurementsPFO_list = app.GetCalcRelevantObjects("*.StaExt*mea")
pfMeasurements = pfList2Dict(measurementsPFO_list)
#Studycases
studycasesPFO_list = scenfolder.GetContents("*.IntScenario")
pfStudyCases = pfList2Dict(studycasesPFO_list)

"""
Process coordinates data for a graphical representation of the network in PowerFactory
"""
max_x, min_x, max_y, min_y = None, None, None, None
coordinates_dict = {}
if coordinates:
    x_list = []
    y_list = []
    for row in coordinates:
        x_list.append(float(row["x"]))
        y_list.append(float(row["y"]))
        coordinates_dict[row["id"]] = row
    max_x = max(x_list)
    min_x = min(x_list)
    max_y = max(y_list)
    min_y = min(y_list)
    gridGrf.iUTrSet = 1
    gridGrf.rULBotX = (min_x)*1.1
    gridGrf.rURTopX = (max_x)*1.1
    gridGrf.rULBotY = (min_y)*1.1
    gridGrf.rURTopY = (max_y)*1.1

"""
---------------------------------------------------------------------------------------
                    CREATE POWERFACTORY OBJECTS
---------------------------------------------------------------------------------------
"""
"""
Zones and Areas
"""
app.PrintPlain("=======START IMPORTING ZONES AND AREAS=======")
#First save needed subnets and voltLvls then create them
if substations:
    for row in substations:
        area_set.add(row["subnet"])
        zone_set.add(row["voltLvl"])
for row in nodes:
    if "Feeder" or "feeder" not in row["subnet"]:
        area_set.add(row["subnet"])
    else:
        feeder_set.add(row["subnet"])
    area_set.add(row["subnet"])
    zone_set.add(row["voltLvl"])

# Create PF objects for areas and zones
for areaelement in area_set:
    # check if already an PF object exists with that areaname, if not then create a new object
    if not pfAreas.get(areaelement):
        newarea = pfoc.createArea(areafolder, areaelement)
        pfAreas[newarea.loc_name] = newarea
for zoneelement in zone_set:
    if not pfZones.get(zoneelement):
        newzone = pfoc.createZone(zonefolder, zoneelement)
        pfZones[newzone.loc_name] = newzone
app.PrintPlain("=======FINISHED IMPORTING ZONES AND AREAS======="+"\n")

"""
Import Substations
"""
app.PrintPlain("=======START IMPORTING SUBSTATIONS=======")
if substations:
    for row in substations:
        subArea = pfAreas.get(row["subnet"])
        subZone = pfZones.get(row["voltLvl"])
        if not pfSubstations.get(row["id"]):
            newsubstat = pfoc.createSubstation(grid, row["id"], subArea, subZone)
            pfSubstations[newsubstat.loc_name] = newsubstat
app.PrintPlain("=======FINISHED IMPORTING SUBSTATIONS========="+"\n")

"""
Import Nodes
"""
app.PrintPlain("=======START IMPORTING NODES=======")
bb_idlist = []  # Contains the id (or names) of busbars that belong to a doublebusbar
dbb_list = []  # Contains lists that contain doublebusbar-pairs
doublebusbars = []
pfDbusbars = {}
for row in nodes:
    if row["type"] == "double busbar":
        doublebusbars.append(row)
seen = set()
seen_add = seen.add
# adds all elements it doesn't know yet to seen and all other to seen_twice
seen_twice = set(row["coordID"] for row in doublebusbars if ((row["coordID"] in seen or seen_add(row["coordID"]))))
# check for busbars that have the same coordID to create doublebusbars
if seen_twice:
    for coordid in seen_twice:
        bb_list = []
        for row in doublebusbars:  # This loop is for finding the nodes with the same coordID
            # put the busbars with same coordID in dbb_list
            if (row["coordID"] == coordid):
                bb_list.append(row)
        # now check for busbars in dbb_list with the same voltage -> these are now doublebusbars
        voltage_seen = set()
        voltage_seen_add = voltage_seen.add
        v_seen_twice = set(row["vmR"] for row in bb_list if
                           row["vmR"] in voltage_seen or voltage_seen_add(row["vmR"]))
        if v_seen_twice:
            # loop over the busbar_list with same coordID and delete
            # the ones that are not doublebusbars, i.e. busbars with the same rated voltage are
            # considered to be doublebusars
            nodeAList = []
            nodeBList = []
            for vmR in v_seen_twice:
                bb_one = False
                for bb in bb_list:
                    if (bb["vmR"] == vmR):
                        if bb_one == False:
                            nodeAList.append(bb)
                            bb_one = True
                        elif bb_one == True:
                            nodeBList.append(bb)
            # now add these busbars to a list
            if len(nodeAList) == len(nodeBList):
                for i in range(0,len(nodeAList)):
                    nodeA = nodeAList[i]
                    nodeB = nodeBList[i]
                    dbb_list.append([nodeA, nodeB])
        for bb in bb_list:
            if not v_seen_twice or bb["vmR"] not in v_seen_twice:
                substatcoordID = bb["coordID"]
                substatcoord = coordinates_dict.get(substatcoordID)
                substat_x = float(substatcoord.get("y"))
                substat_y = float(substatcoord.get("x"))
                if not pfSubstations.get(bb["substation"]):
                    substat = pfoc.createSubstation(grid, "Substation_" + bb["id"], pfAreas.get(bb["subnet"]),
                                                    pfZones.get(bb["voltLvl"]), substat_x, substat_y)
                    pfSubstations[substat.loc_name] = substat
                else:
                    substat = pfSubstations.get(bb["substation"])
                dbb_node = pfoc.createNode(substat, bb, area=pfAreas.get(bb["subnet"]),
                                           zone=pfZones.get(bb["voltLvl"]), x=substat_x, y=substat_y)
                pfDbusbars[dbb_node.loc_name] = dbb_node
                pfNodes[dbb_node.loc_name] = dbb_node
    # Check the node_list for doublebusbars that do not belong to a substation and
    # in these cases create a new substation
    for doublebusbar in dbb_list:
        nodeA = doublebusbar[0]
        nodeB = doublebusbar[1]
        # check if vmsetp is a float, if there is no value in the csv file then set the value of vmSetp to 1
        if nodeA["vmSetp"] and nodeA["vmSetp"] != "NULL":
            vtarget = float(nodeA["vmSetp"])
        else:
            vtarget = 1.0
        uknom = float(nodeA["vmR"])
        if nodeA["vmMin"]:
            vmin = float(nodeA["vmMin"])
        else:
            vmin = 0.95
        if nodeA["vmMax"]:
            vmax = float(nodeA["vmMax"])
        else:
            vmax = 1.05
        if coordinates:        # Get the coordinates
            substatcoordID = nodeA["coordID"]
            substatcoord = coordinates_dict.get(substatcoordID)
            substat_x = float(substatcoord.get("y"))
            substat_y = float(substatcoord.get("x"))
        if not pfSubstations.get(row["substation"]):
            substat = pfoc.createSubstation(grid, nodeA["id"] + "-" + nodeB["id"], pfAreas.get(nodeA["subnet"]), pfZones.get(nodeA["voltLvl"]), substat_x, substat_y)
            pfSubstations[substat.loc_name] = substat
        else:
            substat = pfSubstations.get(row["substation"])

        newnodeA, newnodeB = pfoc.createDoubleBusbar(substat, nodeA["id"], nodeB["id"], 0, vtarget, uknom, vmin, vmax, pfAreas.get(nodeA["subnet"]), pfZones.get(nodeA["voltLvl"]), substat_x, substat_y)
        pfDbusbars[newnodeA.loc_name] = newnodeA
        pfDbusbars[newnodeB.loc_name] = newnodeB

        pfNodes[newnodeA.loc_name] = newnodeA
        pfNodes[newnodeB.loc_name] = newnodeB
for row in doublebusbars:
    if row["coordID"] not in seen_twice:
        if coordinates:
            substatcoordID = row["coordID"]
            substatcoord = coordinates_dict.get(substatcoordID)
            substat_x = float(substatcoord.get("y"))
            substat_y = float(substatcoord.get("x"))
        if not pfSubstations.get(row["substation"]):
            substat = pfoc.createSubstation(grid, "Substation_" + row["id"], pfAreas.get(row["subnet"]),
                                            pfZones.get(row["voltLvl"]), substat_x, substat_y)
            pfSubstations[substat.loc_name] = substat
        else:
            substat = pfSubstations.get(row["substation"])
        dbb_node = pfoc.createNode(substat, row, area = pfAreas.get(row["subnet"]), zone = pfZones.get(row["voltLvl"]), x = substat_x, y = substat_y)
        pfDbusbars[dbb_node.loc_name] = dbb_node
        pfNodes[dbb_node.loc_name] = dbb_node

for row in nodes:
    nodeArea = pfAreas.get(row["subnet"])
    nodeZone = pfZones.get(row["voltLvl"])
    if row["type"] == "busbar" or row["type"] == "node":
        if coordinates:        # Get the coordinates
            nodecoordID = row["coordID"]
            nodecoord = coordinates_dict.get(nodecoordID)
            node_x = float(nodecoord.get("y"))
            node_y = float(nodecoord.get("x"))
        usage = 0   #set the usage of the PowerFactory node object, 0 for busbar or node, 2 for auxiliary node
        if (row["substation"]) != "NULL": #check if the node is inside a substation
            substation = pfSubstations.get(row["substation"])
            if not pfNodes.get(row["id"]):
                newnode = pfoc.createNode(substation, row, nodeArea, nodeZone, usage, node_x, node_y)
                pfNodes[newnode.loc_name] = newnode
        else: #if it is not inside a substation it is a regular PowerFactory "Terminal"-object
            if not pfNodes.get(row["id"]):
                newnode = pfoc.createNode(grid, row, nodeArea, nodeZone, usage, node_x, node_y)
                pfNodes[newnode.loc_name] = newnode
        # Add a graphic-object with coordinates in PowerFactory that represents the node
        if coordinates:
            for coord in coordinates:
                if (coord["id"] == row["coordID"]):
                    x_coord = (float(coord["x"])) + 10 * abs(min_x)
                    y_coord = (float(coord["y"])) + 10 * abs(min_y)
            if 'newnode' in locals(): #check if newnode exists
                if row["substation"] == "NULL":
                    createGraphic(gridGrf, newnode.loc_name, "TermStrip", newnode, x_coord, y_coord)
                elif pfSubstations.get(row["substation"]):
                    createGraphic(gridGrf, row["substation"], "GeneralCompCirc",
                                  pfSubstations.get(row["substation"]), x_coord, y_coord)
    elif row["type"] == "auxiliary": #check if it is an auxiliary node
        auxnodes[row["id"]] = row

slacknodes = findSlacks(app, nodes, pfNodes)
app.PrintPlain("=======FINISHED IMPORTING NODES========="+"\n")

"""
Import Switches
"""
app.PrintPlain("=======START IMPORTING SWITCHES=======")
if not no_sw:
    for row in switches:
        # check if it is a coupler or a switch -> check if one node is an auxiliary node (switch) or not (coupler)
        if (row["nodeA"] not in auxnodes and row["nodeB"] not in auxnodes):  # check if it is a coupler
            if not pfCouplers.get(row["id"]):
                nodeA = pfNodes.get(row["nodeA"])
                nodeB = pfNodes.get(row["nodeB"])
                #First create the cubicles to connect the coupler to
                newcubicle1 = pfoc.createCubicle(nodeA, row["nodeA"])
                newcubicle2 = pfoc.createCubicle(nodeB, row["nodeB"])
                #Now create the coupler
                substation = nodeA.cpSubstat
                newcoupler = pfoc.createCoupler(substation, row, newcubicle1, newcubicle2)
                pfCouplers[newcoupler.loc_name] = newcoupler
        # check if a bay needs to be created for a double busbar
        elif not pfSwitches.get(row["id"]):
            if pfDbusbars.get(row["nodeA"]) and auxnodes.get(row["nodeB"]):
                node = pfDbusbars.get(row["nodeA"])
                auxnoderow = auxnodes.get(row["nodeB"])
                usage = 2
                # find the corresponding PF-object to that auxiliary-node if it exists otherwise create it
                if pfAuxnodes.get(row["nodeB"]):
                    auxnode = pfAuxnodes.get(row["nodeB"])
                else:
                    if node.cpSubstat:
                        substation = node.cpSubstat
                        auxnode = pfoc.createNode(substation, auxnoderow, usage=usage)
                        pfAuxnodes[auxnode.loc_name] = auxnode
                    else:
                        auxnode = pfoc.createNode(grid, auxnoderow, usage=usage)
                        pfAuxnodes[auxnode.loc_name] = auxnode
                # First create the cubicles to connect the coupler to
                newcubicle1 = pfoc.createCubicle(node, row["nodeA"]+"_"+row["nodeB"])
                pfCubicles[newcubicle1.loc_name] = newcubicle1
                newcubicle2 = pfoc.createCubicle(auxnode, row["nodeB"]+"_"+row["nodeA"])
                pfCubicles[newcubicle2.loc_name] = newcubicle2
                # Now create the cubicle to connect the equipment to(i.e. lines and Transformers)
                if not pfCubicles.get(row["nodeB"]):
                    newcubicle3 = pfoc.createCubicle(auxnode, row["nodeB"])
                    pfCubicles[newcubicle3.loc_name] = newcubicle3
                else:
                    newcubicle3 = pfCubicles.get(row["nodeB"])
                # Now create the switch
                if node.cpSubstat:
                    newcoupler = pfoc.createCoupler(substation, row, newcubicle1, newcubicle2)
                else:
                    newcoupler = pfoc.createCoupler(grid, row, newcubicle1, newcubicle2)
                pfSwitches[newcoupler.loc_name] = newcoupler
            elif auxnodes.get(row["nodeA"]) and (pfDbusbars.get(row["nodeB"]) or pfNodes.get(row["nodeB"])):
                if row["nodeB"] in pfDbusbars.keys():
                    node = pfDbusbars.get(row["nodeB"])
                else:
                    node = pfNodes.get(row["nodeB"])
                auxnoderow = auxnodes.get(row["nodeA"])
                usage = 2
                # find the corresponding PF-object to that auxiliary-node if it exists otherwise create it
                if pfAuxnodes.get(row["nodeA"]):
                    auxnode = pfAuxnodes.get(row["nodeA"])
                else:
                    if node.cpSubstat:
                        substation = node.cpSubstat
                        auxnode = pfoc.createNode(substation, auxnoderow, usage=usage)
                        pfAuxnodes[auxnode.loc_name] = auxnode
                    else:
                        auxnode = pfoc.createNode(grid, auxnoderow, usage=usage)
                        pfAuxnodes[auxnode.loc_name] = auxnode
                # First create the cubicles to connect the coupler to
                newcubicle1 = pfoc.createCubicle(node, row["nodeB"] + "_" + row["nodeA"])
                pfCubicles[newcubicle1.loc_name] = newcubicle1
                newcubicle2 = pfoc.createCubicle(auxnode, row["nodeA"] + "_" + row["nodeB"])
                pfCubicles[newcubicle2.loc_name] = newcubicle2
                # Now create the cubicle to connect the equipment to(i.e. lines and Transformers)
                if not pfCubicles.get(row["nodeA"]):
                    newcubicle3 = pfoc.createCubicle(auxnode, row["nodeA"])
                    pfCubicles[newcubicle3.loc_name] = newcubicle3
                else:
                    newcubicle3 = pfCubicles.get(row["nodeA"])
                # Now create the switch
                if node.cpSubstat:
                    newcoupler = pfoc.createCoupler(substation, row, newcubicle1, newcubicle2)
                else:
                    newcoupler = pfoc.createCoupler(grid, row, newcubicle1, newcubicle2)
                pfSwitches[newcoupler.loc_name] = newcoupler
            else:
                nodeA = None
                nodeB = None
                if pfNodes.get(row["nodeA"]):
                    nodeA = pfNodes.get(row["nodeA"])
                elif pfNodes.get(row["nodeB"]):
                    nodeB = pfNodes.get(row["nodeB"])
                # First create the cubicle to connect the switch to
                # Create a cubicle for "nodeA" of the SimBench csv-file but the name of the cubicle is in column "nodeB"
                if nodeA:
                    newcubicle = pfoc.createCubicle(nodeA, row["nodeB"])
                    pfCubicles[newcubicle.loc_name] = newcubicle
                elif nodeB:
                    newcubicle = pfoc.createCubicle(nodeB, row["nodeA"])
                    pfCubicles[newcubicle.loc_name] = newcubicle
                # Now create the switch
                newswitch = pfoc.createSwitch(row, newcubicle)
                pfSwitches[newswitch.loc_name] = newswitch
app.PrintPlain("=======FINISHED IMPORTING SWITCHES========="+"\n")

"""
Import LineTypes
"""
app.PrintPlain("=======START IMPORTING LINETYPES=======")
for row in linetypes:
    # check if the lineType exist before creating new lineTypes to avoid duplicates
    if not pfLineTypes.get(row["id"]):
        newlinetype = pfoc.createLineType(libfolder, row)
        pfLineTypes[newlinetype.loc_name] = newlinetype
######### save DCLineTypes in a Python dictionary #########
if dclinetypes:
    dclinetypes_dct = {}
    for row in dclinetypes:
        dclinetypes_dct[row["id"]] = row
app.PrintPlain("=======FINISHED IMPORTING LINETYPES========="+"\n")

"""
Import Lines
"""
app.PrintPlain("=======START IMPORTING LINES=======")
for row in lines:
    if "dcline" in row["id"]:
        if not pfRES.get(row["id"]+"_from") or not pfRES.get(row["id"]+"_to"):
            if pfNodes.get(row["nodeA"]):
                cubicleA = pfoc.createCubicle(pfNodes.get(row["nodeA"]), "Cubicle_" + row["id"])
            else:
                cubicleA = pfoc.createCubicle(pfAuxnodes.get(row["nodeA"]), "Cubicle_" + row["id"])
            if pfNodes.get(row["nodeB"]):
                cubicleB = pfoc.createCubicle(pfNodes.get(row["nodeB"]), "Cubicle_" + row["id"])
            else:
                cubicleB = pfoc.createCubicle(pfAuxnodes.get(row["nodeB"]), "Cubicle_" + row["id"])
        dclinetype = dclinetypes_dct.get(row["type"])
        newdcline = pfoc.createDCLine(grid, row, cubicleA, cubicleB, dclinetype)
        pfRES[newdcline[0].loc_name] = newdcline[0]
        pfRES[newdcline[1].loc_name] = newdcline[1]
    else:
        if not pfLines.get(row["id"]):
            linetype = pfLineTypes.get(row["type"])
            cubicleA = pfCubicles.get(row["nodeA"])
            cubicleB = pfCubicles.get(row["nodeB"])
            if cubicleA == None:
                nodeA = pfNodes.get(row["nodeA"])
                cubicleA = pfoc.createCubicle(nodeA, nodeA.loc_name+"_"+row["id"])
            if cubicleB == None:
                nodeB = pfNodes.get(row["nodeB"])
                cubicleB = pfoc.createCubicle(nodeB, nodeB.loc_name+"_"+row["id"])
            newline = pfoc.createLine(grid, row, linetype, cubicleA, cubicleB)
            if cubicleA.GetParent().uknom > newline.typ_id.uline:
                newline.typ_id.uline = cubicleA.GetParent().uknom
            pfLines[newline.loc_name] = newline
app.PrintPlain("=======FINISHED IMPORTING LINES========="+"\n")

"""
Import TransformerTypes
"""
app.PrintPlain("=======START IMPORTING TRANSFORMERTYPES=======")
for row in transformertypes:
    if not pfTransformerTypes.get(row["id"]):
        newtransformertype = pfoc.createTransformerType(libfolder, row)
        pfTransformerTypes[newtransformertype.loc_name] = newtransformertype
app.PrintPlain("=======FINISHED IMPORTING TRANSFORMERTYPES========="+"\n")

"""
Import Transformers
"""
app.PrintPlain("=======START IMPORTING TRANSFORMERS=======")
for row in transformers:
    if not pfTransformers.get(row["id"]):
        transformertype = pfTransformerTypes.get(row["type"])
        cubicleHV = pfCubicles.get(row["nodeHV"])
        cubicleLV = pfCubicles.get(row["nodeLV"])
        if cubicleHV == None:
            nodeHV = pfNodes.get(row["nodeHV"])
            cubicleHV = pfoc.createCubicle(nodeHV, nodeHV.loc_name+"_"+row["id"])
        if cubicleLV == None:
            nodeLV = pfNodes.get(row["nodeLV"])
            cubicleLV = pfoc.createCubicle(nodeLV, nodeLV.loc_name+"_"+row["id"])
        if (row["substation"]) != "NULL":  # check if transformer is inside a substation
            substation = pfSubstations.get(row["substation"])
            newtransformer = pfoc.createTransformer(substation, row, transformertype, cubicleHV, cubicleLV)
            pfTransformers[newtransformer.loc_name] = newtransformer
        else:  # if it is not inside a substation
            newtransformer = pfoc.createTransformer(grid, row, transformertype, cubicleHV, cubicleLV)
            pfTransformers[newtransformer.loc_name] = newtransformer
app.PrintPlain("=======FINISHED IMPORTING TRANSFORMERS========="+"\n")

# """
# Import Transformer3Ws
# """
# app.PrintPlain("=======START IMPORTING TRANSFORMER3W=======")
#
# app.PrintPlain("=======FINISHED IMPORTING TRANSFORMER3W========="+"\n")
#
# """
# Import Transformer3WTypes
# """
# app.PrintPlain("=======START IMPORTING TRANSFORMER3WTYPES=======")
#
# app.PrintPlain("=======FINISHED IMPORTING TRANSFORMER3WTYPES========="+"\n")

"""
Import ExternalNets
"""
app.PrintPlain("=======START IMPORTING EXTERNALNETS=======")
if xnets:
    for row in xnets:
        if not pfXnets.get(row["id"]):
            #elements like ExternalNets or Loads are connectet directly to nodes (without switches), therefore the needed cubicles are not in pfCubicles and need to be created
            cubicle = pfoc.createCubicle(pfNodes.get(row["node"]), "Cubicle_"+row["id"])
            newXnet = pfoc.createXnet(grid, row, cubicle)
            pfXnets[newXnet.loc_name] = newXnet
app.PrintPlain("=======FINISHED IMPORTING EXTERNALNETS========="+"\n")

"""
Import Powerplants
"""
app.PrintPlain("=======START IMPORTING POWERPLANTS=======")
if powerplants:
    for row in powerplants:
        if not pfPP.get(row["id"]):
            cubicle = pfoc.createCubicle(pfNodes.get(row["node"]), "Cubicle_" + row["id"])
            newPP = pfoc.createPowerplant(grid, libfolder, row, cubicle)
            pfPP[newPP.loc_name] = newPP
app.PrintPlain("=======FINISHED IMPORTING POWERPLANTS========="+"\n")

"""
Set slack angle
"""
slacks = []
for key in slacknodes:
    slack_cubicles = slacknodes[key][0].GetContents("*.StaCubic")
    for cubicle in slack_cubicles:
        if cubicle.obj_id.GetClassName() == 'ElmXnet':
            slack = cubicle.obj_id
            slacks.append(slack)
            if slack.bustp == "SL":
                slack.phiini = float(slacknodes[key][1])
        elif cubicle.obj_id.GetClassName() == 'ElmSym':
            slack = cubicle.obj_id
            slacks.append(slack)

"""
Import Loads
"""
app.PrintPlain("=======START IMPORTING LOADS=======")
if loads:
    for row in loads:
        if not pfLoads.get(row["id"]):
            cubicle = pfoc.createCubicle(pfNodes.get(row["node"]), "Cubicle_" + row["id"])
            if not pfLoadTypes.get(row["profile"]):
                newloadtype = pfoc.createLoadType(libfolder, row["profile"])
                pfLoadTypes[newloadtype.loc_name] = newloadtype
            newload = pfoc.createLoad(grid, row, cubicle, pfLoadTypes.get(row["profile"]))
            pfLoads[newload.loc_name] = newload
app.PrintPlain("=======FINISHED IMPORTING LOADS========="+"\n")

"""
Import RES
"""
app.PrintPlain("=======START IMPORTING RES=======")
if reses:
    for row in reses:
        if not pfRES.get(row["id"]):
            if pfNodes.get(row["node"]):
                cubicle = pfoc.createCubicle(pfNodes.get(row["node"]), "Cubicle_" + row["id"])
            else:
                cubicle = pfoc.createCubicle(pfAuxnodes.get(row["node"]), "Cubicle_" + row["id"])
            newres = pfoc.createRES(grid, row, cubicle)
            pfRES[newres.loc_name] = newres
app.PrintPlain("=======FINISHED IMPORTING RES========="+"\n")

"""
Import Storage
"""
app.PrintPlain("=======START IMPORTING STORAGE=======")
if storages:
    for row in storages:
        if not pfRES.get(row["id"]):
            if pfNodes.get(row["node"]):
                cubicle = pfoc.createCubicle(pfNodes.get(row["node"]), "Cubicle_" + row["id"])
            else:
                cubicle = pfoc.createCubicle(pfAuxnodes.get(row["node"]), "Cubicle_" + row["id"])
            newstor = pfoc.createStorage(grid, row, cubicle)
            pfRES[newstor.loc_name] = newstor
app.PrintPlain("=======END IMPORTING STORAGE========="+"\n")

"""
Import Shunts
"""
app.PrintPlain("=======START IMPORTING SHUNTS=======")
if shunts:
    for row in shunts:
        if not pfShunts.get(row["id"]):
            if pfNodes.get(row["node"]):
                cubicle = pfoc.createCubicle(pfNodes.get(row["node"]), "Cubicle_" + row["id"])
            else:
                cubicle = pfoc.createCubicle(pfAuxnodes.get(row["node"]), "Cubicle_" + row["id"])
            newshunt = pfoc.createShunt(grid, row, cubicle)
            pfShunts[newshunt.loc_name] = newshunt
app.PrintPlain("=======FINISHED IMPORTING SHUNTS========="+"\n")

"""
Import Measurements
"""
app.PrintPlain("=======START IMPORTING MEASUREMENTS=======")
if measurements:
    for row in measurements:
        if not pfMeasurements.get(row["id"]):
            if pfNodes.get(row["element1"]) and row["element2"] == "NULL" and row["variable"] == "v":
                node = pfNodes.get(row["element1"])
                newmeas = pfoc.createMeasurement(node, row)
            elif pfNodes.get(row["element1"]) and row["element2"] == "NULL" and row["variable"] != "v":
                #In PowerFactory 2019 there is no external p- or q-measurement possible directly at a node, only at a bay of a node
                newmeas = None
            elif pfNodes.get(row["element1"]):
                node = pfNodes.get(row["element1"])
                cubicles = node.GetContents("*.StaCubic")
                for cubicle in cubicles:
                    if cubicle.obj_id.loc_name == row["element2"]:
                        cubicleA = cubicle
                        newmeas = pfoc.createMeasurement(cubicleA, row)
                    else:
                        continue
            elif pfCubicles.get(row["element1"]):
                cubicle = pfCubicles.get(row["element1"])
                newmeas = pfoc.createMeasurement(cubicle, row)
            if newmeas:
                pfMeasurements[newmeas.loc_name] = newmeas
app.PrintPlain("=======FINISHED IMPORTING MEASUREMENTS========="+"\n")

"""
Create station controllers for "pv" nodes
"""
newnodePFO_list = app.GetCalcRelevantObjects("*.ElmTerm")
for node in newnodePFO_list:
    StaCtrlNeed = False
    nodecontents = node.GetContents("*.StaCubic")
    GenUnits = []
    for cubicle in nodecontents:
        if cubicle.obj_id:
            if (cubicle.obj_id.GetClassName() == 'ElmGenstat' and cubicle.obj_id.av_mode == 'constv') or (cubicle.obj_id.GetClassName() == 'ElmSym' and cubicle.obj_id.av_mode == 'constv'):
                GenUnits.append(cubicle.obj_id)
                StaCtrlNeed = True
    if StaCtrlNeed:
        pfoc.createStaCtrl(grid, node, GenUnits)

"""
Create graphical objects
"""
# Final steps for creating graphical objects in PowerFactory
layout = app.GetFromStudyCase('ComSgllayout')
layout.iAction = 0
layout.iGenType = 0
layout.pGrids = grid
layout.Execute()

"""
---------------------------------------------------------------------------------------
                                    ADD PROFILES
---------------------------------------------------------------------------------------
"""
"""
Import Loadprofiles
"""
app.PrintPlain("=======START IMPORTING LOAD PROFILES=======")
try:
    with open(os.path.join(folderpath,"LoadProfile.csv"), "r") as csv_file:
        reader = csv.DictReader(csv_file, delimiter=';')
        colnames = list(reader.fieldnames)  #save columnnames of the csv-file in a list
        colnames.pop(0)                     #remove first columname (time-cloumn) from list
        colindex=2
        for colname in colnames:
            if not pfProfiles.get(colname):
                newChar = charFolder.CreateObject("ChaTime",colname)
                newChar.source = 1  # defining external file to be the source
                newChar.iopt_stamp = 1  # setting Time Stamped Data to be true
                newChar.timeformat = "DD.MM.YYYY hh:mm"  # Time Format
                # Path to csv file
                newChar.f_name = os.path.join(folderpath,"LoadProfile.csv")
                newChar.usage = 1  # value usage (1 = multiply the parameter value with char-values)
                newChar.datacol = colindex #setting the column of the inputfile that contains the data
                colindex = colindex+1
                newChar.iopt_sep = 0 # defining seperation manually
                newChar.col_Sep = ";"  # defining column seperator
                newChar.dec_Sep = "."  # defining decimal seperator
            pfProfiles[newChar.loc_name] = newChar
except IOError:
    app.PrintWarn("File 'LoadProfile.csv' does not appear to exist.")
app.PrintPlain("=======FINISHED IMPORTING LOAD PROFILES========="+"\n")

"""
Assign profiles to loads
"""
app.PrintPlain("=======START ASSIGNING LOAD PROFILES=======")
allPFloads = app.GetCalcRelevantObjects('*.ElmLod')
for load in allPFloads:
    #Look for all existing (old) characteristics saved in that load-objecr and delete it
    sOld = load.GetContents('*.Cha*')
    for i in sOld:
        i.Delete()
    # Assign P-profile
    refObj = load.CreateObject('ChaRef', 'plini')   # Create ChaRef object and name it plini
    refObj.outserv = activate_timeprofile #initial state of the time characteristics is out of service
    # Assign created ChaTime to ChaRef
    if pfProfiles.get(load.typ_id.loc_name+"_pload") != None:
        refObj.typ_id = pfProfiles.get(load.typ_id.loc_name+"_pload")
    # Assign Q-profile
    refObj = load.CreateObject('ChaRef', 'qlini')   # Create ChaRef object and name it qlini
    refObj.outserv = activate_timeprofile #initial state of the time characteristics is out of service
    # Assign created ChaTime to ChaRef
    if pfProfiles.get(load.typ_id.loc_name+"_qload") != None:
        refObj.typ_id = pfProfiles.get(load.typ_id.loc_name+"_qload")
app.PrintPlain("=======FINISHED ASSIGNING LOAD PROFILES========="+"\n")

"""
Import RESProfiles
"""
app.PrintPlain("=======START IMPORTING RES PROFILES=======")
try:
    with open(os.path.join(folderpath,"RESProfile.csv"), "r") as csv_file:
        reader = csv.DictReader(csv_file, delimiter=';')
        colnames = list(reader.fieldnames)  #save columnnames of the csv-file in a list
        colnames.pop(0)                     #remove first columname (time-cloumn) from list
        colindex=2
        for colname in colnames:
            if not pfProfiles.get(colname):
                newChar = charFolder.CreateObject("ChaTime",colname)
                newChar.source = 1  # defining external file to be the source
                newChar.iopt_stamp = 1  # setting Time Stamped Data to be true
                newChar.timeformat = "DD.MM.YYYY hh:mm"  # Time Format
                # Path to csv file
                newChar.f_name = os.path.join(folderpath,"RESProfile.csv")
                newChar.usage = 1  # value usage (1 = multiply the parameter value with char-values
                newChar.datacol = colindex #setting the column of the inputfile that contains the data
                colindex = colindex+1
                newChar.iopt_sep = 0 # defining seperation manually
                newChar.col_Sep = ";"  # defining column seperator
                newChar.dec_Sep = "."  # defining decimal seperator
                pfProfiles[newChar.loc_name] = newChar
except IOError:
    app.PrintWarn("File 'RESProfile.csv' does not appear to exist.")
app.PrintPlain("=======FINISHED IMPORTING RES PROFILES========="+"\n")

"""
Import StorageProfiles
"""
app.PrintPlain("=======START IMPORTING STORAGE PROFILES=======")
try:
    with open(os.path.join(folderpath,"StorageProfile.csv"), "r") as csv_file:
        reader = csv.DictReader(csv_file, delimiter=';')
        colnames = list(reader.fieldnames)  #save columnnames of the csv-file in a list
        colnames.pop(0)                     #remove first columname (time-cloumn) from list
        colindex=2
        for colname in colnames:
            if not pfProfiles.get(colname):
                newChar = charFolder.CreateObject("ChaTime",colname)
                newChar.source = 1  # defining external file to be the source
                newChar.iopt_stamp = 1  # setting Time Stamped Data to be true
                newChar.timeformat = "DD.MM.YYYY hh:mm"  # Time Format
                # Path to csv file
                newChar.f_name = os.path.join(folderpath,"StorageProfile.csv")
                newChar.usage = 1  # value usage (1 = multiply the parameter value with char-values
                newChar.datacol = colindex #setting the column of the inputfile that contains the data
                colindex = colindex+1
                newChar.iopt_sep = 0 # defining seperation manually
                newChar.col_Sep = ";"  # defining column seperator
                newChar.dec_Sep = "."  # defining decimal seperator
                pfProfiles[newChar.loc_name] = newChar
except IOError:
    app.PrintWarn("File 'StorageProfile.csv' does not appear to exist.")
app.PrintPlain("=======FINISHED IMPORTING RES PROFILES========="+"\n")


"""
Assign profiles to RES and Storages
"""
app.PrintPlain("=======START ASSIGNING RES PROFILES=======")
allPFreses = app.GetCalcRelevantObjects('*.ElmGenstat')
for res in allPFreses:
    #Look for all existing characteristics in that RES and delete it
    sOld = res.GetContents('*.Cha*')
    for i in sOld:
        i.Delete()
    # Create ChaRef object and name it plini
    refObj = res.CreateObject('ChaRef', 'pgini')
    refObj.outserv = activate_timeprofile #initial state of the time characteristics is out of service
    # Assign created ChaTime randomly to ChaRef
    # if res.desc != None: #pfProfiles.get(res.desc[0]):
    if res.desc:
        refObj.typ_id = pfProfiles.get(res.desc[0])
app.PrintPlain("=======FINISHED ASSIGNING RES PROFILES========="+"\n")

"""
Import PowerPlant profiles
"""
app.PrintPlain("=======START IMPORTING POWERPLANT PROFILES=======")
try:
    with open(os.path.join(folderpath,"PowerPlantProfile.csv"), "r") as csv_file:
        reader = csv.DictReader(csv_file, delimiter=';')
        colnames = list(reader.fieldnames)  #save columnnames of the csv-file in a list
        colnames.pop(0)                     #remove first columname (time-cloumn) from list
        colindex=2
        for colname in colnames:
            if not pfProfiles.get(colname):
                newChar = charFolder.CreateObject("ChaTime",colname)
                newChar.source = 1  # defining external file to be the source
                newChar.iopt_stamp = 1  # setting Time Stamped Data to be true
                newChar.timeformat = "DD.MM.YYYY hh:mm"  # Time Format
                # Path to csv file
                newChar.f_name = os.path.join(folderpath,"PowerPlantProfile.csv")
                newChar.usage = 1  # value usage (1 = multiply the parameter value with char-values)
                newChar.datacol = colindex #setting the column of the inputfile that contains the data
                colindex = colindex+1
                newChar.iopt_sep = 0 # defining seperation manually
                newChar.col_Sep = ";"  # defining column seperator
                newChar.dec_Sep = "."  # defining decimal seperator
            pfProfiles[newChar.loc_name] = newChar
except IOError:
    app.PrintWarn("File 'PowerPlantProfile.csv' does not appear to exist.")
app.PrintPlain("=======FINISHED IMPORTING POWERPLANT PROFILES========="+"\n")

"""
Assign profiles to powerplants
"""
app.PrintPlain("=======START ASSIGNING POWERPLANT PROFILES=======")
allPFpp = app.GetCalcRelevantObjects('*.ElmSym')
for pp in allPFpp:
    #Look for all existing characteristics in that RES and delete it
    sOld = pp.GetContents('*.Cha*')
    for i in sOld:
        i.Delete()
    # Create ChaRef object and name it plini
    refObj = pp.CreateObject('ChaRef', 'pgini')
    refObj.outserv = activate_timeprofile #initial state of the time characteristics is out of service
    # Assign created ChaTime randomly to ChaRef
    if pp.desc:
        refObj.typ_id = pfProfiles.get(pp.desc[0])
app.PrintPlain("=======FINISHED ASSIGNING POWERPLANT PROFILES========="+"\n")

"""
---------------------------------------------------------------------------------------
                                    ADD STUDYCASES
---------------------------------------------------------------------------------------
"""

"""
Add studycases
"""
app.PrintPlain("=======START CREATING STUDY CASES=======")
if studycases:
    # Get all loads from the PowerFactory model
    allloads = app.GetCalcRelevantObjects("*.ElmLod")
    # Get all RESes from the PowerFactory model
    allRES = app.GetCalcRelevantObjects("*.ElmGenStat")

    for row in studycases:
        if not pfStudyCases.get(row["Study Case"]):
            newstudycase = pfoc.createStudyCase(scenfolder, row, allloads, allRES, slacks)
            pfStudyCases[newstudycase.loc_name] = newstudycase
app.PrintPlain("=======FINISHED CREATING STUDY CASES========="+"\n")

app.EchoOn()