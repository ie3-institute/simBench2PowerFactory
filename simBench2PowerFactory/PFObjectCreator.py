"""#################################################################################################
                            Functions for creating PowerFactory objects
####################################################################################################"""

def createArea(areafolder, areaname):
    newarea = areafolder.CreateObject("ElmArea", areaname)
    return newarea

def createZone(zonefolder, zonename):
    newzone = zonefolder.CreateObject("ElmZone", zonename)
    return newzone

def createSubstation(gridfolder, substat_name, area, zone, x = None, y = None):
    newsubstat = gridfolder.CreateObject("ElmSubstat", substat_name)
    newsubstat.pArea = area
    newsubstat.pZone = zone
    if x != None:
        newsubstat.GPSlat = x
    if y != None:
        newsubstat.GPSlon = y
    return newsubstat

def createNode(folder, row, area = None, zone = None, usage = 0, x = None, y = None):
    newnode = folder.CreateObject("ElmTerm", row["id"])
    # check if vmsetp is a float, if there is no value in the csv file then set the value of vmSetp to 1
    if row["vmSetp"] != "NULL":
        newnode.vtarget = float(row["vmSetp"])
    else:
        newnode.vtarget = 1.0
    newnode.uknom = float(row["vmR"])
    newnode.vmin = float(row["vmMin"])
    newnode.vmax = float(row["vmMax"])
    if x != None:
        newnode.GPSlat = x
    if y != None:
        newnode.GPSlon = y
    newnode.cpArea = area
    newnode.cpZone = zone
    newnode.iUsage = usage
    return newnode

#Function for creating a single busbar of a doublebusbar
def createBusbar(folder, name, iusage=0, vtarg=1.0, uknom=110, vmin=0.95, vmax=1.05, cparea=None, cpzone=None, x = None, y = None):
    newnode = folder.CreateObject("ElmTerm")
    newnode.loc_name = name
    newnode.iUsage = iusage
    newnode.vtarget = vtarg
    newnode.uknom = uknom
    newnode.vmin = vmin
    newnode.vmax = vmax
    if x != None:
        newnode.GPSlat = x
    if y != None:
        newnode.GPSlon = y
    newnode.cpArea = cparea
    newnode.cpZone = cpzone
    return newnode

def createDoubleBusbar(folder, bb1_name, bb2_name, iusage=0, vtarg=1.0, uknom=110, vmin=0.95, vmax=1.05, cparea=None, cpzone=None, x = None, y = None):
    # Create 2 nodes
    bb1 = createBusbar(folder, bb1_name, iusage, vtarg, uknom, vmin, vmax, cparea, cpzone, x, y)
    bb2 = createBusbar(folder, bb2_name, iusage, vtarg, uknom, vmin, vmax, cparea, cpzone, x, y)
    return bb1, bb2

def createCubicle(node, cubiclename):
    newcubicle = node.CreateObject("StaCubic", cubiclename)
    return newcubicle

def createCoupler(folder, row, cubicle1, cubicle2):
    newcoupler = folder.CreateObject("ElmCoup", row["id"])
    if (row["type"] == "CB"):
        newcoupler.aUsage = "cbk"
    elif (row["type"] == "LS"):
        newcoupler.aUsage = "swt"
    elif (row["type"] == "LBS"):
        newcoupler.aUsage = "sdc"
    elif (row["type"] == "DS"):
        newcoupler.aUsage = "dct"
    if (row["cond"] == "0"):
        newcoupler.on_off = 0
    else:
        newcoupler.on_off = 1
    newcoupler.bus1 = cubicle1
    newcoupler.bus2 = cubicle2
    return newcoupler

def createSwitch(row, cubicle):
    newswitch = cubicle.CreateObject("StaSwitch", row["id"])
    if (row["type"] == "CB"):
        newswitch.aUsage = "cbk"
    elif (row["type"] == "LS"):
        newswitch.aUsage = "swt"
    elif (row["type"] == "LBS"):
        newswitch.aUsage = "sdc"
    elif (row["type"] == "DS"):
        newswitch.aUsage = "dct"
    if (row["cond"] == "0"):
        newswitch.on_off = 0
    else:
        newswitch.on_off = 1
    return newswitch

#Create a coupler connecting the two Busbars
def createdbbCoupler(folder, row, nodeA, nodeB):
    newcoupler = folder.CreateObject("ElmCoup", row["id"])
    if (row["type"] == "CB"):
        newcoupler.aUsage = "cbk"
    elif (row["type"] == "LS"):
        newcoupler.aUsage = "swt"
    elif (row["type"] == "LBS"):
        newcoupler.aUsage = "sdc"
    elif (row["type"] == "DS"):
        newcoupler.aUsage = "dct"
    if (row["cond"] == "0"):
        newcoupler.on_off = 0
    else:
        newcoupler.on_off = 1
    # Create cubicle and connect it to the terminals
    # Connection for nodeA
    sta_cubicle1 = nodeA.CreateObject("StaCubic")
    sta_cubicle1.loc_name = nodeA.loc_name
    newcoupler.bus1 = sta_cubicle1
    # Connection for nodeB
    sta_cubicle2 = nodeB.CreateObject("StaCubic")
    sta_cubicle2.loc_name = nodeB.loc_name
    newcoupler.bus2 = sta_cubicle2
    return newcoupler

def createLineType(libfolder, row):
    newlineType = libfolder.CreateObject("TypLne", row["id"])
    newlineType.rline = float(row["r"])  # positive sequence resistence @ 20 °C
    newlineType.xline = float(row["x"])  # positive sequence reactance
    newlineType.bline = float(row["b"])  # positive sequence susceptance
    newlineType.sline = float(row["iMax"]) / 1000  # iMax
    if row["type"] == "ohl":
        newlineType.cohl_ = 1
    else:
        newlineType.cohl_ = 0
    return newlineType

def createDCLine(folder, dcline_row, cubicleA, cubicleB, dclintype_row):
    # Create static generator at node A
    newDCgenA = folder.CreateObject("ElmGenStat", dcline_row["id"]+"_from")
    newDCgenA.cCategory = "hvdc"
    newDCgenA.mode_inp = "PQ"
    if float(dclintype_row["pDCLine"]) > 0:
        newDCgenA.pgini = -float(dclintype_row["pDCLine"])
    else:
        newDCgenA.pgini = 1
    newDCgenA.cosn = 1  # set defaultvalue of the powerfactor to 1
    newDCgenA.sgn = newDCgenA.pgini*1.5
    newDCgenA.Pmin_uc = 0
    if dclintype_row["pMax"] == "NULL":
        newDCgenA.Pmax_uc = 0
    else:
        newDCgenA.Pmax_uc = float(dclintype_row["pMax"])
    if dclintype_row["qMinA"] == "NULL":
        newDCgenA.cQ_min = 0
    else:
        newDCgenA.cQ_min = float(dclintype_row["qMinA"])
    if dclintype_row["qMaxA"] == "NULL":
        newDCgenA.cQ_max = 0
    else:
        newDCgenA.cQ_max = float(dclintype_row["qMaxA"])
    newDCgenA.bus1 = cubicleA
    # Create static generator at node B
    newDCgenB = folder.CreateObject("ElmGenStat", dcline_row["id"]+"_to")
    newDCgenB.cCategory = "hvdc"
    newDCgenB.mode_inp = "PQ"
    newDCgenB.pgini = (newDCgenA.pgini - float(dclintype_row["fixPLosses"]) * (1 - (float(dclintype_row["relPLosses"]) / 100)))
    newDCgenB.cosn = 1  # set defaultvalue of the powerfactor to 1
    newDCgenB.sgn = newDCgenB.pgini*1.5
    newDCgenB.Pmin_uc = 0
    if dclintype_row["pMax"] == "NULL":
        newDCgenB.Pmax_uc = 0
    else:
        newDCgenB.Pmax_uc = float(dclintype_row["pMax"])
    if dclintype_row["qMinB"] == "NULL":
        newDCgenB.cQ_min = 0
    else:
        newDCgenB.cQ_min = float(dclintype_row["qMinB"])
    if dclintype_row["qMaxB"] == "NULL":
        newDCgenB.cQ_max = 0
    else:
        newDCgenB.cQ_max = float(dclintype_row["qMaxB"])
    newDCgenB.bus1 = cubicleB
    return [newDCgenA, newDCgenB]

def createLine(folder, row, linetype, cubicleA, cubicleB):
    newline = folder.CreateObject("ElmLne", row["id"])
    newline.typ_id = linetype
    newline.bus1 = cubicleA
    newline.bus2 = cubicleB
    newline.dline = float(row["length"])
    if row["loadingMax"]:
        newline.maxload = float(row["loadingMax"])
    else:
        newline.maxload = 100
    return newline

def createTransformerType(libfolder, row):
    newtransformertype = libfolder.CreateObject("TypTr2", row["id"])
    newtransformertype.strn = float(row["sR"])
    newtransformertype.utrn_h = float(row["vmHV"])
    newtransformertype.utrn_l = float(row["vmLV"])
    newtransformertype.nt2ag = float(row["va0"]) / 30
    newtransformertype.uktr = float(row["vmImp"])
    newtransformertype.pcutr = float(row["pCu"])
    newtransformertype.pfe = float(row["pFe"])
    newtransformertype.curmg = float(row["iNoLoad"])
    # Stufung
    if row["tapable"] == "1":
        newtransformertype.itapch = 1
        newtransformertype.tapside = 0 if row["tapside"] is "HV" else 1
        newtransformertype.dutap = float(row["dVm"])
        newtransformertype.phitr = float(row["dVa"])
        newtransformertype.nntap0 = int(row["tapNeutr"])
        newtransformertype.ntpmn = int(row["tapMin"])
        newtransformertype.ntpmx = int(row["tapMax"])
    return newtransformertype

def createTransformer(folder, row, transformertype, cubicleHV, cubicleLV):
    newtransformer = folder.CreateObject("ElmTr2", row["id"])
    newtransformer.bushv = cubicleHV
    newtransformer.buslv = cubicleLV
    newtransformer.typ_id = transformertype
    newtransformer.nntap = int(row["tappos"])
    newtransformer.ntrcn = 1 if row["autoTap"] == "1" else 0
    newtransformer.t2ldc = 0 if row["autoTapSide"] == "HV" else 1
    newtransformer.maxload = float(row["loadingMax"])
    return newtransformer

def createXnet(folder, row, cubicle):
    newxnet = folder.CreateObject("ElmXnet", row["id"])
    newxnet.usetp = cubicle.GetParent().vtarget
    newxnet.bus1 = cubicle
    if row["calc_type"] == "pq":
        newxnet.bustp = "PQ"
    if row["calc_type"] == "pv":
        newxnet.bustp = "PV"
    if row["calc_type"] == "vavm":
        newxnet.bustp = "SL"
    return newxnet

def createSMType(libfolder, node, row):
    newSMtype = libfolder.CreateObject("TypSym", row["id"]+"_type")
    newSMtype.sgn = float(row["sR"])
    newSMtype.ugn = node.uknom
    newSMtype.cosn = 0.95
    return newSMtype

def createPowerplant(folder, libfolder, row, cubicle):
    newpp = folder.CreateObject("ElmSym", row["id"])
    #create synchronous machine (SM) type and assign it to this SM
    newpp.typ_id = createSMType(libfolder, cubicle.GetParent(), row)
    # check the type and set the category in PowerFactory
    if row["type"] == "hard coal":
        newpp.cCategory = 'coal'
        newpp.cSubCategory = 'hardcoal'
    elif row["type"] == "lignite":
        newpp.cCategory = 'coal'
        newpp.cSubCategory = 'lignite'
    elif row["type"] == "nuclear":
        newpp.cCategory = 'nuc'
    elif row["type"] == "gas":
        newpp.cCategory = 'gas'
    elif row["type"] == "oil":
        newpp.cCategory = 'oil'
    # check the controltype and set the category in PowerFactory
    if row["calc_type"] == "vavm":
        newpp.ip_ctrl = 1
        newpp.av_mode = 'constv'
    elif row["calc_type"] == "pvm":
        newpp.av_mode = 'constv'
    elif row["calc_type"] == "pq":
        newpp.av_mode = 'constq'
    if row["pPP"] == "NULL":
        newpp.pgini = 0
    else:
        newpp.pgini = float(row["pPP"])
    if row["qPP"] == "NULL":
        newpp.qgini = 0
        newpp.mode_inp = "SP"
        newpp.sgini = float(row["sR"])
    else:
        newpp.qgini = float(row["qPP"])
    newpp.Pmin_uc = float(row["pMin"])
    newpp.Pmax_uc = float(row["pMax"])
    newpp.cQ_min = float(row["qMin"])
    newpp.cQ_max = float(row["qMax"])
    newpp.bus1 = cubicle
    newpp.desc = [row["profile"]]
    return newpp

def createStaCtrl(folder, node, genunitlist):
    newstactrl = folder.CreateObject("ElmStactrl", "Stactrl_" + node.loc_name)
    newstactrl.psym = genunitlist
    newstactrl.i_ctrl = 0  # 0 => voltage control
    newstactrl.selBus = 0
    newstactrl.uset_mode = 1
    newstactrl.rembar = node

def createLoad(folder, row, cubicle, loadtype):
    newload = folder.CreateObject("ElmLod", row["id"])
    newload.plini = abs(float(row["pLoad"]))
    newload.qlini = abs(float(row["qLoad"]))
    newload.typ_id = loadtype
    newload.bus1 = cubicle
    return newload

def createLoadType(libfolder, name):
    newloadtype = libfolder.CreateObject("TypLod", name)
    newloadtype.systp = 0
    newloadtype.phtech = 2
    return newloadtype

def createRES(folder, row, cubicle):
    newres = folder.CreateObject("ElmGenStat", row["id"])
    if ("pv" in row["type"].lower()):
        newres.cCategory = "pv"
    elif ("wind" in row["type"].lower()):
        newres.cCategory = "wgen"
        if ("offshore" in row["type"].lower()):
            newres.cSubCategory = "offshore"
    elif ("biomass" in row["type"].lower()):
        newres.cCategory = "bgas"
    elif ("hydro" in row["type"].lower() or "river" in row["type"].lower()):
        newres.cCategory = "hydr"
    if (row["calc_type"] == "pq"):
        newres.mode_inp = "PQ"
        newres.pgini = float(row["pRES"])
        newres.qgini = float(row["qRES"])
    newres.cosn = 1  # set defaultvalue of the powerfactor to 1
    if float(row["sR"]) > 0:
        newres.sgn = float(row["sR"])
    else:
        newres.sgn = 1
    newres.Pmin_uc = 0
    newres.Pmax_uc = float(row["pRES"])
    newres.cQ_min = -float(row["qRES"])
    newres.cQ_max = float(row["qRES"])
    newres.bus1 = cubicle
    newres.desc = [row["profile"]]
    return newres

def createStorage(folder, row, cubicle):
    newstor = folder.CreateObject("ElmGenStat", row["id"])
    newstor.cCategory = "stor"
    newstor.sgn = float(row["sR"])
    newstor.cosn = 1
    newstor.pgini = -float(row["pStor"])
    newstor.qgini = float(row["qStor"])
    newstor.Pmin_uc = float(row["pMin"])
    newstor.Pmax_uc = float(row["pMax"])
    newstor.cQ_min = float(row["qMin"])
    newstor.cQ_max = float(row["qMax"])
    newstor.bus1 = cubicle
    newstor.desc = [row["profile"]]
    return newstor

def createShunt(folder, row, cubicle):
    newshunt = folder.CreateObject('ElmShnt', row["id"])
    newshunt.bus1 = cubicle
    newshunt.ushnm = row["vmR"]
    newshunt.ncapx = row["Step"]
    newshunt.ncapa = 1          #set current step to startstep
    newshunt.qtotn = row["q0"]
    return newshunt

def createMeasurement(folder, row):
    newmeas = folder.CreateObject('StaExt'+row["variable"]+"mea", row["id"])
    return newmeas

def createStudyCase(folder, row, loads, reses, slacks):
    newstudycase = folder.CreateObject('IntScenario', row["Study Case"])
    newstudycase.Activate()
    #Set study case loadvalues
    for load in loads:
        load.plini = load.plini * float(row["pload"])
        load.qlini = load.qlini * float(row["qload"])
    # Set study case RES-values
    for res in reses:
        if res.cCategory == 'Wind':
            res.pgini = res.pgini * float(row["Wind_p"])
        elif res.cCategory == 'Fotovoltaik':
            res.pgini = res.pgini * float(row["PV_p"])
        else:
            res.pgini = res.pgini * float(row["RES_p"])
    # Set slack voltages
    if slacks:
        for slack in slacks:
            if slack.GetClassName() == 'ElmXnet':
                slack.usetp = float(row["Slack_vm"])
    newstudycase.Save()
    newstudycase.Deactivate()
    return newstudycase