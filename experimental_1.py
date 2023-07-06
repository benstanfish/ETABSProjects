# Import the library that will allow us to grab CSi Objects
import comtypes.client

# Get the active ETABS process
# The ETABSObject is a cOAPI object, refer to documentation
ETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")

# Get the open ETABS model. This is a cSapModel object.
SapModel = ETABSObject.SapModel

# API calls return 0 when the program responds properly.
# Make sure the model is unlocked, otherwise commands may not work. This is also a good way to test that you have
# control of the model.
SapModel.SetModelIsLocked(False)

# Create list of dummy names required for interacting with model. Use the C# documentation (despite this being Python)
NumberNames = 0
MyName = []

# Set the present units of the project. Refer to eUnits enumeration
kip_in_F = 4
SapModel.SetPresentUnits(kip_in_F)

# Set some of the project information
project_info = {"Client Name": "",
                "Project Name": "",
                "Project Number": "",
                "Company Name": "",
                "Company Logo": "",
                "Engineer": "",
                "Checker": "",
                "Supervisor": "",
                "Model Name": "",
                "Model Description": "",
                "Revision Number": "1",
                "Issue Number": ""}
for i, key in enumerate(project_info.keys()):
    SapModel.SetProjectInfo(key, project_info[key])

# Create dummy variables per the cGridSys.GetNameList method documentation

grid_system_count = SapModel.GridSys.GetNameList(NumberNames, MyName)[0]
grid_system_names = SapModel.GridSys.GetNameList(NumberNames, MyName)[1]

# Use the GetGridSys_2 method to review detailed information about the grid system.
# First create dummy variables per cGridSys.GetGridSys_2 method documentation
Xo = 0.0                #1
Yo = 0.0                #2
RZ = 0.0                #3
GridSysType = ""        #4
NumXLines = 0           #5
NumYLines = 0           #6
GridLineIDX = []        #7
GridLineIDY = []        #8
OrdinateX = []          #9
OrdinateY = []          #10
VisibleX = []           #11
VisibleY = []           #12
BubbleLocX = []         #13
BubbleLocY = []         #14

grid_system_data = SapModel.GridSys.GetGridSys_2(grid_system_names[0], Xo, Yo, RZ, GridSysType, NumXLines, NumYLines, GridLineIDX,
                                    GridLineIDY, OrdinateX, OrdinateY,VisibleX,VisibleY,BubbleLocX,BubbleLocY)

# Note that the SetGridSys method actually adds a new Grid System to the model.



# Let's deal with point objects
GUID = ""
X = 0.0
Y = 0.0
Z = 0.0
CSys = ""

# Get All the Point Objects
points = SapModel.PointObj.GetNameList(NumberNames,MyName)
point_count = points[0]
point_names = points[1]
point_success = points[2]     # Will be 0 if operation successful

# Get all the point guids
point_guids = {}
for pnt in point_names:
    point_guids[pnt] = SapModel.PointObj.GetGuid(pnt,GUID)[0]

# Let's pull all the coordinates for the points
point_coords = {}
for i, pnt in enumerate(point_names):
    point_name = point_names[i]
    ret = SapModel.PointObj.GetCoordCartesian(point_name,X,Y,Z)
    point_coords[point_name] = (ret[0],ret[1],ret[2])

# Let's add new points using the AddCartesian method; create dummy variables
CSys = "Global"
MergeOff = False
MergeNumber = 0

new_point = SapModel.PointObj.AddCartesian(2,2,50,Name='777',UserName="u777",CSys="Global",MergeOff=False,MergeNumber=0)
new_point_unique = new_point[0]
new_point_success = new_point[1]     # Will be 0 if operation successful

