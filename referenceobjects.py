from startup import *

# Create list of dummy names required for interacting with model.
# Use the C# documentation (despite this being Python)
NumberNames = 0
MyName = []

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

grid_system_data = SapModel.GridSys.GetGridSys_2(grid_system_names[0], Xo, Yo, RZ, GridSysType,
                                                 NumXLines, NumYLines, GridLineIDX, GridLineIDY,
                                                 OrdinateX, OrdinateY, VisibleX, VisibleY,
                                                 BubbleLocX, BubbleLocY)

# Note that the SetGridSys method actually adds a new Grid System to the model.



# Let's deal with point objects
GUID = ""
X = 0.0
Y = 0.0
Z = 0.0
CSys = ""
