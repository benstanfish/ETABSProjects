# Import the library that will allow us to grab CSi Objects
import comtypes.client
import os, sys

# Create a CSi API Helper object
helper = comtypes.client.CreateObject('ETABSv1.Helper')
helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

try:
    # create an instance of the ETABS object from the latest installed ETABS
    # The ETABSObject is a cOAPI object, refer to documentation
    ETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")

except (OSError, comtypes.COMError):
    print("Cannot start a new instance of the program.")
    sys.exit(-1)


ETABSObject.ApplicationStart()


clientAPIVersion = helper.GetOAPIVersionNumber()
programAPIVersion = ETABSObject.GetOAPIVersionNumber()
if clientAPIVersion > programAPIVersion:
    print('API client uses newer version than program API.')
    print('Some API functionality may not be available.')
else:
    print('Client API: '+clientAPIVersion)
    print('Program API: '+programAPIVersion)

# initialize a new model, API returns ret = 0 if successful
SapModel = ETABSObject.SapModel
ret = SapModel.InitializeNewModel()
if ret == 0:
    print('Model successfully created.')
else:
    print('There was a error, could not create model.')

file_save = False
ETABSObject.ApplicationExit(file_save)
print('Program successfully closed without saving')
