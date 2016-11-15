
# coding: utf-8

from win32api import GetSystemMetrics
import comtypes
from comtypes.client import CreateObject

# Open STK and create a new scenario
# Create an instance of STK
app=CreateObject("STK11.Application")
app.Visible=True
app.UserControl=True

app.Top=0
app.Left=0
app.Width=int(GetSystemMetrics(0)/2)
app.Height=int(GetSystemMetrics(1)-30)

# Get our IAgStkObjectRoot interface
root=app.Personality2

from comtypes.gen import STKUtil
from comtypes.gen import STKObjects

# Create new Scenario
root.NewScenario("IPython_DIY")
sc=root.CurrentScenario

# Set Scenario Analysis Interval
sc2=sc.QueryInterface(STKObjects.IAgScenario)
sc2.SetTimePeriod("10 Jun 2016 04:00:00","11 Jun 2016 04:00:00")
root.Rewind();

# Insert Facility and set its position
fac= sc.Children.New(STKObjects.eFacility,"codeFacility")
fac2 = fac.QueryInterface(STKObjects.IAgFacility)
fac2.Position.AssignGeodetic(38.9943,-76.8489,0)

# Insert Satellite and set its orbital elements
sat = sc.Children.New(STKObjects.eSatellite, "codeSat")
sat2= sat.QueryInterface(STKObjects.IAgSatellite)
sat2.SetPropagatorType(STKObjects.ePropagatorJ2Perturbation)

satProp = sat2.Propagator
satProp=satProp.QueryInterface(STKObjects.IAgVePropagatorJ2Perturbation)
satProp.InitialState.Epoch="08 Jun 2016 15:14:26"

# Use the Classical Element interface
keplerian = satProp.InitialState.Representation.ConvertTo(STKUtil.eOrbitStateClassical)
# Changes from SMA/Ecc to MeanMotion/Ecc
keplerian2 = keplerian.QueryInterface(STKObjects.IAgOrbitStateClassical)
keplerian2.SizeShapeType =STKObjects.eSizeShapeMeanMotion
# Makes sure Mean Anomaly is being used
keplerian2.LocationType = STKObjects.eLocationMeanAnomaly
# Use RAAN instead of LAN for data entry
keplerian2.Orientation.AscNodeType = STKObjects.eAscNodeRAAN

# Set unit preferences for revs/day
root.UnitPreferences.Item('AngleUnit').SetCurrentUnit('revs')
root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('day')

# Assign the perigee and apogee altitude values:
keplerian2.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeMeanMotion).MeanMotion = 15.08385840
keplerian2.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeMeanMotion).Eccentricity = 0.0002947

# Return unit preferences for degrees and seconds
root.UnitPreferences.Item('AngleUnit').SetCurrentUnit('deg')
root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('sec')

# Assign the other desired orbital parameters:
keplerian2.Orientation.Inclination = 28.4703
keplerian2.Orientation.ArgOfPerigee = 114.7239
keplerian2.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = 315.1965
keplerian2.Location.QueryInterface(STKObjects.IAgClassicalLocationMeanAnomaly).Value = 332.9096

# Apply the changes made to the satellite's state and propagate:
satProp.InitialState.Representation.Assign(keplerian)
satProp.Propagate()


cartVel=sat.DataProviders("Cartesian Velocity")
cartVel=cartVel.QueryInterface(STKObjects.IAgDataProviderGroup)

cartVelJ2000=cartVel.Group.Item("J2000")
cartVelJ2000TimeVar = cartVelJ2000.QueryInterface(STKObjects.IAgDataPrvTimeVar)

rptElements=['Time','x','y','z']

velResult=cartVelJ2000TimeVar.ExecElements(sc2.StartTime,sc2.StopTime,60,rptElements)

time=velResult.DataSets.Item(0).GetValues()
x=velResult.DataSets.Item(1).GetValues()
y=velResult.DataSets.Item(2).GetValues()
z=velResult.DataSets.Item(3).GetValues()


import pandas as pd

df=pd.DataFrame({'time':time,'x':x,'y':y,'z':z});


#ALTcartVelJ2000TimeVar=sat.DataProviders.GetDataPrvTimeVarFromPath("Cartesian Velocity//J2000")

#ALTvelResults=ALTcartVelJ2000TimeVar.ExecElements(sc2.StartTime,sc2.StopTime,60,rptElements)

#del root;
#app.Quit();
#del app
