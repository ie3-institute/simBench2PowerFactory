# simBench2PowerFactory
This tool converts SimBench data sets into DIgSILENT PowerFactory. 
SimBench has been research project of Kassel university, Fraunhofer IEE, RWTH Aachen university and TU Dortmund 
university, aiming at provision of realistic models of electrical distribution grids including time series. 
The data is availabe at [https://simbench.de](https://simbench.de/en/).

Another conversion tool for converting SimBench data sets to ie3's power system data model is available 
[here](https://github.com/ie3-institute/simBench2psdm]).

Furthermore, the tool [here](https://github.com/e2nIEE/simbench) enables the use of the SimBench data set in the 
simulation tool [pandapower](https://github.com/e2nIEE/pandapower).

## How to use
1. First additional packeges need to be installed, these are:
    - os
    - csv
    - datetime
2. To use the converter in PowerFactory a Python command object (ComPython) needs to be created in the PowerFactory 
project. 
This object creates a link between PowerFactory and a python script file. 
A new Python command object can be created in a PowerFactory project under "Library" -> "Scripts". 
In a newly created ComPython object, the path of the SimBench converter file ('SimBench2PowerFactory.py') can be set 
under the "Script" tab. 
In the "Basic Options" tab of the ComPython object an additional input parameter must be created. 
String needs to be  chosen as "type" and the name  of the parameter must  be "folder". 
The "value" of the parameter must be the folder path  under which the SimBench csv data to be imported 
into PowerFactory can be found. 
For example this could be something  like "D:\SimBench\1-LV-semiurb4--0-sw”. 
The converter is then set and can be executed within PowerFactory. Further general information on how to build and 
use a Python command object is described in detail in the PowerFactory manual.