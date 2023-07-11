#######################################################################
#                       Get active instance
#######################################################################
# Call this as a dependency to grab an existing ETABS process

# Import the library that will allow us to grab CSi Objects
import comtypes.client

# Get the active ETABS process
# The ETABSObject is a cOAPI object, refer to documentation
ETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")

# Get the open ETABS model. This is a cSapModel object.
SapModel = ETABSObject.SapModel

# API calls return 0 when the program responds properly.
# Make sure the model is unlocked, otherwise commands may not work.
# This is also a good way to test that you have control of the model.
SapModel.SetModelIsLocked(False)