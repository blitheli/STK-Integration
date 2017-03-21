
# coding: utf-8

# 
# # STK Integration with Python
# <br/>
# <br/>
# ## Frank Snyder
# <br/>
# ### Product Evangelist - Analytical Graphics, Inc.
# #### fsnyder@agi.com
# #### www.agi.com
# https://github.com/agifsnyder/STK-Integration/

# ## We're going to need the following things:
# ### - A working instance of Python, complete with the following modules
#   - comtypes, pandas, IPython as well as a working Jupyter Notebooks environment
# 
# Now we're going to learn something new
# The simplest way to set up the right kind of Python environment is to download the Anaconda Python installer from Continuum Analytics [here](https://www.continuum.io/downloads). Not only is it extremely complete in terms of useful scientific libraries, it's **Free** and properly licensed for both Personal as well as Commercial Use, including redistribution.
# We'll pick the right 32/64-bit version for our Windows Operating System, and it's highly recommend to stick with the Python 3.5 version (unless there is a specific dependency that's known already to have to use Python 2.7)
# It's also recommended to let the installer default to installing in the local user folders rather than in a Global {Program Files} kind of location.  _This avoids the need to have Administrative priviledges to perform the install._
# 
# ### - A properly installed, and licensed, instance of STKv11.1, which can be requested [here](http://www.agi.com/products/stk/license/).
#   - Specifically we will need the STK/Integration License which activates the STK API for use by external environments, like Python|Matlab|Java|etc...
# 
# ### - A working internet connection if you wish to explore the STK Online documentation linked below in this Notebook

# First we set up a few imports to gain access to their nested properties and functions.

# In[1]:

from win32api import GetSystemMetrics
from IPython.display import Image, display, SVG
import os as os


# Now we need to import references to `comtypes` and some of it's methods so we can tie to the STK COM-based objects.

# In[2]:

import comtypes
from comtypes.client import CreateObject


# ### Let's start by creating a new running instance of the STK (AgUiApplication) Application

# In[3]:

app=CreateObject("STK11.Application")


#  If this is the first time you're running this Notebook against STKv11.1, you might see the following:
# 
# ```
# # Generating comtypes.gen._781C4C18_C2C9_4E16_B620_7B22BC8DE954_0_1_0
# # Generating comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
# # Generating comtypes.gen.stdole
# # Generating comtypes.gen._9B797FC6_9EF1_4779_9691_A85B091560A1_0_1_0
# # Generating comtypes.gen.AgUiCoreLib
# # Generating comtypes.gen.AgUiApplicationLib
# ```
# 
# as `comtypes` generates Python "wrappers" around the Classes and Interfaces found at the **_AgUiApplication_** level.
# 
# 
# let's check what kind of object we've been given back. To do that we can use the Python `type()`
# call and pass in the variable referencing the object we're interested in.

# In[4]:

type(app)


# It's important to recognize that we're getting a **POINTER** to the **_IAgUiApplication_** which is normal behavior for Python.  This means that this variable points-to the representation of the **_AgUiApplication_** class, or instance of that class in memory.

# Let's also point out the **very** important difference between the _Object_ that is **_AgUiApplication_** and the _Interface_ **_IAgUiApplication_** that it exposes for use. As per AGI's naming conventions used in the API, you can usually knock the **_I_** off the front end of the keyword to figure out what kind of object it is, as opposed to what kind of interface it's currently exposing.

# In[5]:

app.Visible=True
app.UserControl= True


# As per most Windows Applications, this is where you can usually do things like position and resize the Application's main window.  Here we're going to place it on the left side of the screen, extending halfway to the right.

# In[6]:

app.Top=0
app.Left=0
app.Width=int(GetSystemMetrics(0)/2)
app.Height=int(GetSystemMetrics(1)-30)


# #### Now lets get to work. 
# We need to start using the STK Object model, and to do that we want to get a reference to the Object Model Tree itself.
# In AGI terms, we refer to the particular version of the structure of that Object Model as it's "Personality".  To date there are two existing versions of that structure, and we want to use the latest one, simply called **_Personality2_**

# In[7]:

root=app.Personality2


# If this is the first time you've run this Notebook, you'll see how the Python `comtypes` module generates Python "wrappers" around the rest of the Classes and Interfaces that are defined in the STK Object Model Tree, from the root level on down:
# 
# ```
# # Generating comtypes.gen._D6A1725B_89FF_43A4_995B_7F055549F4EB_0_1_0
# # Generating comtypes.gen._00DD7BD4_53D5_4870_996B_8ADB8AF904FA_0_1_0
# # Generating comtypes.gen.STKUtil
# # Generating comtypes.gen._8B49F426_4BF0_49F7_A59B_93961D83CB5D_0_1_0
# # Generating comtypes.gen.AgSTKVgtLib
# # Generating comtypes.gen._42D2781B_8A06_4DB2_9969_72D6ABF01A72_0_1_0
# # Generating comtypes.gen.AgSTKGraphicsLib
# # Generating comtypes.gen.STKObjects
# ```
# 
# Two of the top-level ones we care about, and will be using repeatedly through our introductory exploration of the STK Object Model API are:
# - **STKUtil**
# - **STKObjects**
# 
# lets keep their python "paths" in mind so we can import them later.  If you've run a python script that used `comtypes` to "touch" STK in the past, those generated wrappers will already exist in the `comtypes.client.gen-dir` folder location

# In[9]:

comtypes.client.gen_dir


# In[10]:

os.listdir(comtypes.client.gen_dir)


# It's important to know where these cached python wrappers are located in case you want to remove them, like when trying to use this same script against a different version of STK.  At the time of writing this, we're using STKv11.1, so the classID's you see listed in that folder reflect the "current" version of each Type Library used by STKv11.1
# 
# In general, those classID's won't change except for released updates at the "Major Number" level of STK, such as STKv9.x.x, STKv10.x.x, STKv11.x.x, etc...
# 
# As a general rule of thumb, you'll want to clear the contents of the `gen` folder of the STK wrappers after you've done a major number upgrade of STK.
# 
# 

# #### So, now let's also check what kind of object we've been given back.

# In[11]:

type(root)


# The Interface is named **_IAgStkObjectRoot_** **NOT** **_IAgStkObject_** (no root) so it looks like a "special" kind of STK object, and it is.  The Root object in the STK API sits at the very **top** of the Object Heirarchy Tree.
# 
# #### From this `root` object we can reach all the sub-objects, methods, and properties by traversing the API Object Model.

# Next, we're going to assign two new variables to the autogenerated Python wrappers of the STK Classes found in `comtypes.gen.STKObjects` and `comtypes.gen.STKUtil`.  These will be useful when transitioning between Interfaces exposed by each STK object we will use, as well as all the readable Enumerations (ex. eScenario = 19)

# In[12]:

from comtypes.gen import STKUtil
from comtypes.gen import STKObjects


# ### Now let's create a Scenario, the start of any kind of STK-related work

# In[13]:

root.NewScenario("IPython_DIY")


# and lets set the variable `sc` to the current scenario using the exposed **CurrentScenario** method off the `root` object

# In[14]:

sc=root.CurrentScenario


# In[15]:

type(sc)


# The object interface we get back is of type **_IAgStkObject_** which is a generic base class for **every** Object found in STK.  As such, {_Satellites, Ships, Aircraft, Facilities_, etc...} ALL implement the **_IAgStkObject_** interface consistently accross their respective Classes, including the _Scenario_ itself, but for our next step we need to work with _Scenario_-specific kind of properties and methods, such as the `SetTimePeriod()` method, which are only going to be found when the `sc` Object "behaves" like a _Scenario_ rather than a generic StkObject, so we need to "cast" the `sc` object into it's **_IAgScenario_** Interface.  
# 
# We do that via the `QueryInterface`() generic method that Python added onto the `sc` object returned by STK's `root` Object, and for future convenience, simply set a different `sc2` object to that change. Now we can use both the generic **_IAgStkObject_** features of the _Scenario_ by using the `sc` variable, and the more _Scenario_-specific ones of **_IAgScenario_** by using the `sc2` variable.
# 
# If you forget which one represents which exposed interface at any time, you can simply use the tab-completion feature to see what Properties and Methods exist for each `sc`/`sc2` variable.
# 
# 
# We use the `STKObjects.IAgScenario` "knowledge" gained when the STK Type Libraries were scanned and those equivalent Python Classes created. Both `sc` and `sc2` "point" to the very same object, but different exposed interfaces.
# ##### Online docs for [_IAgSTKObject_](http://help.agi.com/resources/help/online/stkdevkit/11.1/index.html?page=source%2Fextfile%2FSTKObjects%2FSTKObjects~IAgStkObject.html)
# 
# ##### Online docs for [_IAgSCenario_](http://help.agi.com/resources/help/online/stkdevkit/11.1/index.html?page=source%2Fextfile%2FSTKObjects%2FSTKObjects~IAgScenario.html)

# In[19]:

sc2=sc.QueryInterface(STKObjects.IAgScenario)
type(sc2)


# Now that it's implementing the right interface, we can set the Time Period to be what we want
# <table>
# <tr><td>Start Time</td><td>10 Jun 2016 04:00:00 UTCG</td></tr>
# <tr><td>Stop Time</td><td>11 Jun 2016 04:00:00 UTCG</td></tr>
# </table>
# 
# The times are conveniently in a **Gregorian** format, referencing UTC, so we don't need to convert any Date/Time units because STK's Default Date settings are to use UTCG, but it's always good to pay attention to details like this.

# In[20]:

sc2.SetTimePeriod("10 Jun 2016 04:00:00","11 Jun 2016 04:00:00")


# and while we're at it, lets reset the Animation time for the scenario to the begining.
# this is an action (method) that can called from the root level.

# In[21]:

root.Rewind();


# ### Next we need to create a Facility on the ground
# It's position will be defined by Latitude, Longitude and Altitude, which are traditionally termed **Geodetic** coordinates.
# <table>
# <tr><td>Latitude</td><td>38.9943 deg</td></tr>
# <tr><td>Longitude</td><td>-76.8489 deg (- implies West)</td></tr>
# <tr><td>Altitude (above WGS84)</td><td>0 m</td></tr>
# </table>

# So let's first create the Facility, using the sc object and it's **_IAgStkObject_** Interface

# The `sc` object can return a Collection of "Children" that hierarchically fall under it.  That collection also has a `New` Method we can use to create any kind of STK object that is allowed to be attached underneath the _Scenario_ level of the STK Object Browser Tree.  In the following example, we can "gang" multiple traverses through the object model using additional '.' and tab-completions, as follows:

# In[22]:

fac= sc.Children.New(STKObjects.eFacility,"codeFacility")


# Now we want to define it's position, preferably using a Geodetic frame of reference.  To do that, we need to use the objects **IAgFacility** interface instead of it's **IAgStkObject** interface.

# In[23]:

fac2 = fac.QueryInterface(STKObjects.IAgFacility)
fac2.Position.AssignGeodetic(38.9943,-76.8489,0)


# ### Now we want to create a satellite with the following characteristics:
#  - Use the J2 Propagator
#  - Use Classical (keplerian) Elements, but slightly modified from "traditional" Classical Elements
#    - **Mean Motion** for the size instead of Semi-Major Axis
#    - **Mean Anomaly** for the location in the orbit instead of True Anomaly
# 
# We also want to use slightly different units for the Mean Motion (revs/day)
# 
# <table>
# <th>Specific Values</th>
# <tr><td>Orbit Epoch (state Epoch)</td><td>08 Jun 2016 15:14:26</td></tr>
# <tr><td>Mean Motion</td><td>15.08385840 revs/day</td></tr>
# <tr><td>Eccentricity</td><td>0.0002947</td></tr>
# <tr><td>Inclination</td><td>28.4703 deg</td></tr>
# <tr><td>Argument of Perigee</td><td>114.7239 deg</td></tr>
# <tr><td>RAAN</td><td>315.1965 deg</td></tr>
# <tr><td>Mean Anomaly</td><td>332.9096 deg</td></tr>
# </table>
# 

# First we create the new "codeSat" Satellite object in the same way we created the Facility, but using a different `STKObjects.eSatellite` enumeration.  That object, of default Type **_IAgStkObject_** gets reflected in the `sat` variable.
# 
# Of course we want to do _Satellite_ things with this Object, so let's create a second `sat2` variable that's set to use the **_IAgSatellite_** interface.

# In[24]:

sat = sc.Children.New(STKObjects.eSatellite, "codeSat")
sat2= sat.QueryInterface(STKObjects.IAgSatellite)


# There is a "pattern" to how a Satellite is created, it's properties defined, and it's propagator asked to calculate it's position over a specified time period.  That pattern can be easily understood by following the sequence of actions performed in the GUI when setting up a Satellite.  
# 
# <div class="alert alert-info">NOTE: This same philosophy applies to working with any other STK object as well.</div>
# 
# 
# **When laying out the STK Object Model hierarchial structure back in STKv5.0, when it first appeared, the developers chose to mimic the functional layout of the GUI, but in code.** 
# 
# So, by simply remembering how we set up an object in the GUI (mouse clicks and selections), we have a basic "roadmap" for how to do the same in code.  The **sequence** of events done in the GUI usually needs to be replicated in a similar sequential fashion using code.
# 
# So, when the first step in defining a Satellite is the selection of the Propagator, as seen below:

# <img src="./images/1.SelectJ2PropagatorFromPulldown.png"/>

# where the very first thing you do in this GUI panel is select from a dropdown of available Propagators the J2Perturbation propagator. That selection will change the GUI layout to reflect the type of propagator selected, and the data entry options that are allowed.
# 
# Here's the code equivalent:

# To get a list of the allowed Propagators for this object, we can "ask" it for what's supported

# In[25]:

sat2.PropagatorSupportedTypes


# and we see that J2Perturbation is available, along with a few others that have been activated by the valid licenses that STK saw on startup. 
# 
# <div class="alert alert-info">NOTE: Only items that are activated by valid licenses will show up in this list, or be available when assignining at runtime.</div>
# 
# Now we can set the Propagator type using a related method `SetPropagatorType` and passing it the `STKObjects.ePropagatorJ2Perturbation` enumeration (which actualy resolves to the number "1")

# In[26]:

sat2.SetPropagatorType(STKObjects.ePropagatorJ2Perturbation)


# Now the .Propagator "property" of the `sat2` object will return an instance of the right kind of J2Perturbation object for us to work with.

# In[27]:

satProp = sat2.Propagator


# In[28]:

type(satProp)


# `satProp` is currently showing that it's exposing an **_IAgVePropagator_** interface. Looking up that "_IAgVePropagator_" keyword in the docs returns a page showing co-Classes that implement **_IAgPropagator_**. That means that all these **_AgVe_**... Classes also reference and expose the **_IAgVePropagator_** interface, and possibly other interfaces as well.
# 
# The class **_AgVePropagatorJ2Perturbation_** looks interesting and is more applicable to our needs, so lets take a closer look at that...
# 
# It has 2 Interfaces:
#  - **_IAgVePropagator_** : Base vehicle propagator interface, with no properties or methods
#  - **_IAgVePropagatorJ2Perturbation_** : Class defining the J2 perturbation propagator
# 
# and since the base interface **_IAgVePropagator_** didn't have anything useful for us to use, let's Cast `satProp` into it's **_IAgVePropagatorJ2Perturbation_** interface to see what we can do with that.

# In[29]:

satProp=satProp.QueryInterface(STKObjects.IAgVePropagatorJ2Perturbation)
type(satProp)


# This is one of the strange cases in the STK Object Model where we initially get an "object" exposing a generic base Interface (**_IAgVePropagator_**), but there is no equivalent **_AgVePropagator_** (base) class or object that goes along with it.  There isn't even one in the documentation.  Instead, this propagator object is **already** an instamce of the specific **_AgVePropagatorJ2Perturbation_** class of object because it came out of the _Satellite_ object **after** it had been configured to use the J2Perturbation propagator.  It was an **_AgVePropagatorJ2Perturbation_** when the satProp variable was assigned, but it simply reverted to the most base-level's of interfaces that it had, the **_IAgVePropagator_** interface.
# 
# So, if we had, just out of curiosity, tried to cast `satProp`'s interface over to an **_IAgVePropagatorJ4Perturbation_** (which is similar) we would get an error because that interface doesn't exist for a J2Perturbation propagator.
# 
# Also, you cannot simply create an instance AgVePropagatorJ2Perturbation from scratch.  It has to be manufactured by a _Satellite_ object, specifically when that object is exposing it's **_IAgSatellite_** interface.  This behavior, in a manner of speaking, conforms to the "Factory" design pattern often used in Object Oriented programming.  Only certain factories (object class) can produce certain objects, and only when running in a particular way (interface).
# 
# This also means that if you wanted, for whatever reason, to change the propagator in use, you would have to go BACK UP to the `sat2` object, change it's selected propagator to what you wanted, and re-request the `.Propagator` property to get the changed propagator object back out of it.
# 
# This very particular workflow is reflected in the GUI behavior as well.  You would need to go back to the top of the GUI panel, change the Propagator selected from the dropdown list and THEN start making changes to the propagator settings.
# 

# Moving forwards through the Propagator setup, we can see from the Object Model "roadmap" below that the InitialState property obtained from the **_IAgVePropagatorJ2Perturbation_** interface will have an **_IAgVeJxInitialState_**.  The **_Jx_** is there because this exact same interface is available for the J4Perturbation propagator as well.
# 
# 

# <img src="./images/IAgVePropagator.png" />
# 
# This visual is derived from the published PDF of the STKv11.1 Object Model Diagram found [here](http://help.agi.com/resources/help/online/stkdevkit/11.1/ObjectModel/pdf/ObjectModel_diagram.pdf)

# #### Setting the Initial State
# Now we need to set up the Initial State conditions for this J2Perturbation propagator.
# This sequence parallels the following GUI actions:
# 
# First, set the **Orbit Epoch** Time (GUI):
# <img src="./images/2.Set_OrbitEpoch.png"/>
# 
# and via `code`:

# In[30]:

satProp.InitialState.Epoch="08 Jun 2016 15:14:26"


# Next, change **Semimajor Axis** to **Mean Motion** and **True Anomaly** to **Mean Anomaly**
# <img src="./images/3.Classical_Orbital_Elements_MeanMotion_MeanAnomaly_RAAN.png"/>
# 
# This is a little trickier to follow.  Essentially the _InitialState_, in it's entirety, has a **Representation** of Type **_IAgOrbitState_** that can be manipulated.  Essentially it's the code/class/interface equivalent of all the parameters found within the dashed blue border seen above.
# 
# That includes the _Coordinate Type_ {Cartesian, Classical/Keplerian, etc...}, it's _Coordinate System of Reference_ {ICRF, J2000, B1950, etc...} and the actual parameters, as appropriate, for how the Coordinates are applied.
# 

# In[31]:

type(satProp.InitialState.Representation)


# We're going to set a new variable, `keplerian`, to the _.Representation_ when it is converted to be a Classical (i.e. Keplerian) Type of parameters.

# In[32]:

keplerian = satProp.InitialState.Representation.ConvertTo(STKUtil.eOrbitStateClassical)


# In[33]:

type(keplerian)


# despite being "converted" to **`eOrbitStateClassical`** the `keplerian` object is still exposing the generic "base" **_IAgOrbitState_** interface.  We want the more specific, and applicable to our type of data, **_IAgOrbitStateClassical_** interface where we can make some changes to how both the _Size_ of the orbit (_MeanMotion_) and the _Location_ (_MeanAnomaly_) in the orbit is specified.
# 
# <div class="alert alert-info">NOTE: It's only because we want to specify these alternative _MeanMotion_ and _MeanAnomaly_ parameters, because they're convenient to describing the orbit we want, that we have to go to this extra level of detail.  The base **_IAgOrbitState_** interface actually has a few "helper" methods such as _AssignClassical_ where the "traditional" Keplerian elements can be set all at once.</div>
# 
# So, `keplerian2` gets assigned with the **_IAgOrbitStateClassical_** interface

# In[34]:

keplerian2 = keplerian.QueryInterface(STKObjects.IAgOrbitStateClassical)


# and now we can get to work changing the Size, Location, and Orientation to suite how our Orbit State data is specified.

# In[35]:

keplerian2.SizeShapeType =STKObjects.eSizeShapeMeanMotion
keplerian2.LocationType = STKObjects.eLocationMeanAnomaly
keplerian2.Orientation.AscNodeType = STKObjects.eAscNodeRAAN


# One last thing we need to do prior to setting values and that's to make sure that the Units will match those of our values.  Again this parallels the GUI actions in selecting what's approriate from the unit drop-downs.
# 
# <div class="alert alert-info">NOTE: Each instance of an STK Root object can have **ONLY ONE** global set of units at any particular time.  So if you have values that are specified in multiple different units for a single kind of measurement type (angles, distance, etc...) you'll need to flip-flop to the right Unit prior to setting & assigning it.</div>
# 
# Hanging off the `root` object is a **_UnitPreferences_** property which returns a Collection of **_IAgUnitPrefsDim_** with it's own interface **_IAgUnitPrefsDimCollection_**.  In this case, while it's possible to traverse down to each _Item_ in the collection, and modify them one layer deeper there are some "shortcuts" you can take by "ganging" calls together.  Since the Collection interface has both an **_Item_** Property as well as a **_SetCurrentUnit_** method they can be combined to make the whole process a little more streamlined as one long "call" per Dimension and Unit as follows:

# In[36]:

root.UnitPreferences.Item('AngleUnit').SetCurrentUnit('revs')
root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('day')


# One anoying detail about how Python and the comtypes module works when interacting with Microsoft COM objects is that when a new object is referenced, if it's capable of exposing multiple Interfaces, the lowest or base class of the bunch is automatically selected, even if it doesn't have any usefull methods or properties.
# 
# Case in point: the SizeShape property of our `keplerian2` object. 

# In[37]:

type(keplerian2.SizeShape)


# We can see that despite having been "set" to a Type enumeration of _eSizeShapeMeanMotion_ a few cells above, it still defaults to the **_IAgClassicalSizeShape_** interface, which is simply a base class for all the following co-classes:
# - **_AgClassicalSizeShapeAltitude_**
# - **_AgClassicalSizeShapeMeanMotion_**
# - **_AgClassicalSizeShapePeriod_**
# - **_AgClassicalSizeShapeRadius_**
# - **_AgClassicalSizeShapeSemimajorAxis_**
# - **_AgEquinoctialSizeShapeMeanMotion_**
# - **_AgEquinoctialSizeShapeSemimajorAxis_**
# 
# each of which having an equivalently specific {...|Altitude|MeanMotion|Period|Radius|SemimajorAxis} Interface as well as the **_IAgClassicalSizeShape_** base interface.
# 
# Since our SizeShape is actually an instance of the **_AgClassicalSizeShapeMeanMotion_** class, the only usefull & specific interface is the **_IAgClassicalSizeShapeMeanMotion_** one.
# 
# Knowing this apriori, as we get more familiar reading through how the documentation is structured, we can "quick-cast" the interface into the interface that we know will work for us, tacking on the **_.MeanMotion_** property because we know it's going to be there.
# 
# The process looks like this:

# In[38]:

keplerian2.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeMeanMotion).MeanMotion = 15.08385840


# In[39]:

keplerian2.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeMeanMotion).Eccentricity = 0.0002947


# Rememberr that our **'AngleUnit'** and **'TimeUnit'** are both still set to **'revs'** and **'days'** respectively from a few cells above. To properly assign the following angular values as degrees we need to flip-flop back to **'deg'**.  We'll change the **'TimeUnit'** as well because 'seconds' is a common default for units that we're going to be referencing later on.

# In[40]:

root.UnitPreferences.Item('AngleUnit').SetCurrentUnit('deg')
root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('sec')
keplerian2.Orientation.Inclination = 28.4703
keplerian2.Orientation.ArgOfPerigee = 114.7239


# Once again, the "base Class" Interface selection mechanism used by Python is going to trip us up.  For the `keplerian2` object both the **_.Orientation.AscNode_** and **_.Location_** will default to fairly useless **_IAgOrientationAscNode_** and **_IAgClassicalLocation_** interfaces respectively.  We need the **_.Orientation.AscNode_** returned object to expose its **_IAgOrientationAscNodeRAAN_** interface in order to set the RAAN Value.  We also want the **_.Location_** returned object to expose its **_IAgClassicalLocationMeanAnomaly_** interface so we can set the Mean Anomaly value as well.
# 
# Again, in-line "quick-casting" to interfaces that we already know, and expect to see, can be done to streamline the code, albeit at the expense of loosing a bit of tab-completion interactivity.

# In[41]:

keplerian2.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = 315.1965
keplerian2.Location.QueryInterface(STKObjects.IAgClassicalLocationMeanAnomaly).Value = 332.9096


# We're almost done with this Propagator.  Now that all the properties are properly set in the `keplerian` object, we can Assign it back into the InitialState.Representation of the propagator itself.
# 
# <div class="alert alert-warning">This is an important step that **MUST** be done in order for the propagator to "learn" about all these updated values. Simply updating the `keplerian` variable will NOT affect the "factory" Propagator where it was produced.  It has to be explicitly Assigned back in order to take effect.<div>

# In[42]:

satProp.InitialState.Representation.Assign(keplerian)


# The last step is to tell the Propagator to **_.Propagate_**.  Without this step, none of the position data will be calculated and the Satellite will remain in a partially configured state. All the properties will have been set, but the code used to to take those properties and calculate its position over its time span won't execute.

# In[43]:

satProp.Propagate()


# ## Getting Data Out -  Data Providers and STK GUI Reports
# 
# ### STK Objects and the Data they can provide
# 
# Every **_AgStkObject_**, through it's **_IAgStkObject_** interface, will expose a **_.DataProviders_** property that will return a Collection of valid DataProvider Groups for that type of object.
# 
# Each _Item()_ of the returned **_IAgDataProviderCollection_** conforms to an **_IAgDataProviderInfo_** interface in addition to exposing one of the three interfaces listed below.
# 
# The **_IAgDataProviderInfo_** interface is another "base" and common interface that only has a few properties a single method.  Essentially, it provides the human-readable _Name_ for the Data Provider as well as the _Type_ of Data Provider it is.

# Data Providers fall into 3 basic categories, or Types:
# - Static Data, defined in the **_AgDataPrvFixed_** Class and handled through the **_IAgDataPrvFixed_** Interface.
#   - examples of this would be Fixed position information, object initial state data, and other parameters that won't vary over time.
# - Time Varying (TimeSeries) Data, defined in the **_AGDataPrvTimeVar_** Class and handled through the **_IAgDataPrvTimeVar_** Interface.
#   - examples of this would be position data for moving vehicles reported over a time interval at specified time steps.
# - Time Intervals, defined in the **_AGDataPrvInterval_** Class and handled through the **_IAgDataPrvInterval_** Interface.
#   - examples of this would be Start|Stop Date/Time values for one or more time intervals, such as Access, No_Access, etc...
# 
# 
# We can better understand the organization of Groups, Data Providers, and Elements found in the Object Model follows by reviewing this structure:
# <img src="./images/Report-2-DatProvider-Structure.png"/>
# which bears a strinking similarity to the GUI layout of the Data Providers selector panel when customizing reports and graphs.
# <img src="./images/GUI_Report_DataProviders.png"/>
# So when planning to pull data from STK, it's generally a good idea to be familiar with the GUI layout prior to setting up the code.  The names, relative positions in heirarchy, and contents will all help define the Interfaces we will need to be looking to work with.

# For our first example, we want to ouput the **_Cartesian Position_** and **_Velocity_** of the _Satellite_ that we just created.  That information will need to be specified in a particular Coordinate System, and for this report we want to use the **_J2000_** Coordinate System.
# 
# Both of those data groups are Time Varying (TimeSeries) data types, so we know we'll need to work with the **_IAgDataPrvTimeVar_** interface as part of the process.
# 
# So, for the velocity data we will be using is the "Cartesian Velocity" _Group_ name, and below it, the "J2000" _Data Provider_ Name, to hone in on the actual _Elements_ {x,y,z,etc...}

# In[44]:

cartVel=sat.DataProviders("Cartesian Velocity")
type(cartVel)


# We've selected the "Cartesian Velocity" group, but need to cast `cartVel` to it's **_IAgDataProviderGroup_** interface to do anything usefull with it.

# In[45]:

cartVel=cartVel.QueryInterface(STKObjects.IAgDataProviderGroup)


# Now we can pull from the collection of DataProviders found in this Group the one named "J2000"

# In[46]:

cartVelJ2000=cartVel.Group.Item("J2000")
type(cartVelJ2000)


# again, we get back the "default" **_IAgDataProviderInfo_** interface, which won't do us much good here.  We need to cast it to the more specific **_IAgDataPrvTimeVar_** interface because it's a Time Varying type of Data Provider, and that interface will let us specify the time span over which we want the data to be generated, as well as which specific elements we want.

# In[47]:

cartVelJ2000TimeVar = cartVelJ2000.QueryInterface(STKObjects.IAgDataPrvTimeVar)
type(cartVelJ2000TimeVar)


# Now, we're going to use the **_ExecElements_** method to return ONLY the {Time,x,y,z} values rather than the entire {Time,x,y,z,speed,radial,in-track} list of _Elements_ that exist for the J2000 _Data Provider_ found in the Cartesian Velocity _Group_.
# 
# We need to pass an list of these named Elements into the method call, so we need to create that list first.

# In[48]:

rptElements=['Time','x','y','z']


# Now we're ready to call _ExecElements_ with a Starting and Ending Time exactly that of the Scenario's, and a time step between generated data points of 60 seconds.

# In[49]:

velResult=cartVelJ2000TimeVar.ExecElements(sc2.StartTime,sc2.StopTime,60,rptElements)
type(velResult)


# In[50]:

time=velResult.DataSets.Item(0).GetValues()


# In[51]:

x=velResult.DataSets.Item(1).GetValues()


# In[52]:

y=velResult.DataSets.Item(2).GetValues()


# In[53]:

z=velResult.DataSets.Item(3).GetValues()


# In[54]:

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# In[55]:

df=pd.DataFrame({'time':time,'x':x,'y':y,'z':z});
df


# In[56]:

df.columns


# In[57]:

ALTcartVelJ2000TimeVar=sat.DataProviders.GetDataPrvTimeVarFromPath("Cartesian Velocity//J2000")
type(ALTcartVelJ2000TimeVar)


# In[58]:

ALTvelResults=ALTcartVelJ2000TimeVar.ExecElements(sc2.StartTime,sc2.StopTime,60,rptElements)
type(ALTvelResults)


# Close things down for a clean exit.
# 
# (commented out for live use.)

# In[59]:

#del root; 
#app.Quit();
#del app;

