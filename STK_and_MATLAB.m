clear; clc;

%Open STK and create a new scenario
    %Create an instance of STK
    app = actxserver('STK11.application');
    app.Visible = 1;

%Get our IAgStkObjectRoot interface
    root = app.Personality2;

%Create new Scenario
    root.NewScenario('Matlab_Training_Scenario');
    sc = root.CurrentScenario;

%Set Scenario Analysis Interval
    sc.SetTimePeriod('10 Jun 2016 04:00:00', '11 Jun 2016 04:00:00');
    root.Rewind();

%Insert Facility and set its position
    fac = sc.Children.New('eFacility', 'codeFacility');
    fac.Position.AssignGeodetic(38.9943,-76.8489,0);

%Insert Satellite and set its orbital elements
    sat = sc.Children.New('eSatellite', 'codeSat');
    
    sat.Propagator.InitialState.Epoch = '08 Jun 2016 15:14:26';
    satProp = sat.Propagator.InitialState.Representation.ConvertTo('eOrbitStateClassical'); % Use the Classical Element interface
    satProp.SizeShapeType = 'eSizeShapeMeanMotion';  % Changes from SMA/Ecc to MeanMotion/Ecc
    satProp.LocationType = 'eLocationMeanAnomaly'; % Makes sure Mean Anomaly is being used
    satProp.Orientation.AscNodeType = 'eAscNodeRAAN'; % Use RAAN instead of LAN for data entry

    %Set unit preferences for revs/day
    root.UnitPreferences.Item('Angle').SetCurrentUnit('revs');
    root.UnitPreferences.Item('Time').SetCurrentUnit('day');
        
    % Assign the perigee and apogee altitude values:
    satProp.SizeShape.MeanMotion = 15.08385840;   % revs/day
    satProp.SizeShape.Eccentricity = 0.0002947;   % unitless

    %Return unit preferences for degrees and seconds
    root.UnitPreferences.Item('Angle').SetCurrentUnit('deg');
    root.UnitPreferences.Item('Time').SetCurrentUnit('sec');
    
    % Assign the other desired orbital parameters:
    satProp.Orientation.Inclination = 28.4703;    % deg
    satProp.Orientation.ArgOfPerigee = 114.7239;  % deg
    satProp.Orientation.AscNode.Value = 315.1965; % deg
    satProp.Location.Value = 332.9096;            % deg

    % Apply the changes made to the satellite's state and propagate:
    sat.Propagator.InitialState.Representation.Assign(satProp);
    sat.Propagator.Propagate;

    %% Output LLA of sat using access times
    dp = sat.DataProviders.Item('Cartesian Velocity').Group.Item('J2000');
    res = dp.ExecElements(sc.StartTime, sc.StopTime, 60, {'Time'; 'x'; 'y'; 'z';});
    time = res.DataSets.GetDataSetByName('Time').GetValues;
    x = res.DataSets.GetDataSetByName('x').GetValues;
    y = res.DataSets.GetDataSetByName('y').GetValues;
    z = res.DataSets.GetDataSetByName('z').GetValues;

    %