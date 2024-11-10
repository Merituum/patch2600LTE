'# MWS Version: Version 2024.1 - Oct 16 2023 - ACIS 33.0.1 -

'# length = mm
'# frequency = MHz
'# time = ns
'# frequency range: fmin = 100 fmax = 4000
'# created = '[VERSION]2024.1|33.0.1|20231016[/VERSION]


'@ use template: b7LTE.cfg

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
'set the units
With Units
    .SetUnit "Length", "mm"
    .SetUnit "Frequency", "MHz"
    .SetUnit "Voltage", "V"
    .SetUnit "Resistance", "Ohm"
    .SetUnit "Inductance", "nH"
    .SetUnit "Temperature",  "degC"
    .SetUnit "Time", "ns"
    .SetUnit "Current", "A"
    .SetUnit "Conductance", "S"
    .SetUnit "Capacitance", "pF"
End With

ThermalSolver.AmbientTemperature "0"

'----------------------------------------------------------------------------

'set the frequency range
Solver.FrequencyRange "2400", "2800"

'----------------------------------------------------------------------------

Plot.DrawBox True

With Background
     .Type "Normal"
     .Epsilon "1.0"
     .Mu "1.0"
     .XminSpace "0.0"
     .XmaxSpace "0.0"
     .YminSpace "0.0"
     .YmaxSpace "0.0"
     .ZminSpace "0.0"
     .ZmaxSpace "0.0"
End With

With Boundary
     .Xmin "expanded open"
     .Xmax "expanded open"
     .Ymin "expanded open"
     .Ymax "expanded open"
     .Zmin "expanded open"
     .Zmax "expanded open"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With

' optimize mesh settings for planar structures

With Mesh
     .MergeThinPECLayerFixpoints "True"
     .RatioLimit "20"
     .AutomeshRefineAtPecLines "True", "6"
     .FPBAAvoidNonRegUnite "True"
     .ConsiderSpaceForLowerMeshLimit "False"
     .MinimumStepNumber "5"
     .AnisotropicCurvatureRefinement "True"
     .AnisotropicCurvatureRefinementFSM "True"
End With

With MeshSettings
     .SetMeshType "Hex"
     .Set "RatioLimitGeometry", "20"
     .Set "EdgeRefinementOn", "1"
     .Set "EdgeRefinementRatio", "6"
End With

With MeshSettings
     .SetMeshType "HexTLM"
     .Set "RatioLimitGeometry", "20"
End With

With MeshSettings
     .SetMeshType "Tet"
     .Set "VolMeshGradation", "1.5"
     .Set "SrfMeshGradation", "1.5"
End With

' change mesh adaption scheme to energy
' 		(planar structures tend to store high energy
'     	 locally at edges rather than globally in volume)

MeshAdaption3D.SetAdaptionStrategy "Energy"

' switch on FD-TET setting for accurate farfields

FDSolver.ExtrudeOpenBC "True"

PostProcess1D.ActivateOperation "vswr", "true"
PostProcess1D.ActivateOperation "yz-matrices", "true"

With FarfieldPlot
	.ClearCuts ' lateral=phi, polar=theta
	.AddCut "lateral", "0", "1"
	.AddCut "lateral", "90", "1"
	.AddCut "polar", "90", "1"
End With

'----------------------------------------------------------------------------

Dim sDefineAt As String
sDefineAt = "2400;2500;2550;2620;2650;2690;2800"
Dim sDefineAtName As String
sDefineAtName = "2400;2500;2550;2620;2650;2690;2800"
Dim sDefineAtToken As String
sDefineAtToken = "f="
Dim aFreq() As String
aFreq = Split(sDefineAt, ";")
Dim aNames() As String
aNames = Split(sDefineAtName, ";")

Dim nIndex As Integer
For nIndex = LBound(aFreq) To UBound(aFreq)

Dim zz_val As String
zz_val = aFreq (nIndex)
Dim zz_name As String
zz_name = sDefineAtToken & aNames (nIndex)

' Define E-Field Monitors
With Monitor
    .Reset
    .Name "e-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Efield"
    .MonitorValue  zz_val
    .Create
End With

' Define H-Field Monitors
With Monitor
    .Reset
    .Name "h-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Hfield"
    .MonitorValue  zz_val
    .Create
End With

' Define Farfield Monitors
With Monitor
    .Reset
    .Name "farfield ("& zz_name &")"
    .Domain "Frequency"
    .FieldType "Farfield"
    .MonitorValue  zz_val
    .ExportFarfieldSource "False"
    .Create
End With

Next

'----------------------------------------------------------------------------

With MeshSettings
     .SetMeshType "Hex"
     .Set "Version", 1%
End With

With Mesh
     .MeshType "PBA"
End With

'set the solver type
ChangeSolverType("HF Time Domain")

'----------------------------------------------------------------------------

'@ define material: FR-4 (lossy)

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Material
     .Reset
     .Name "FR-4 (lossy)"
     .Folder ""
     .FrqType "all"
     .Type "Normal"
     .SetMaterialUnit "GHz", "mm"
     .Epsilon "4.3"
     .Mu "1.0"
     .Kappa "0.0"
     .TanD "0.025"
     .TanDFreq "10.0"
     .TanDGiven "True"
     .TanDModel "ConstTanD"
     .KappaM "0.0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstKappa"
     .DispModelEps "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "General 1st"
     .DispersiveFittingSchemeMu "General 1st"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMu "False"
     .Rho "0.0"
     .ThermalType "Normal"
     .ThermalConductivity "0.3"
     .SetActiveMaterial "all"
     .Colour "0.94", "0.82", "0.76"
     .Wireframe "False"
     .Transparency "0"
     .Create
End With

'@ new component: component1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Component.New "component1"

'@ define brick: component1:DIELECTRIC

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "DIELECTRIC" 
     .Component "component1" 
     .Material "FR-4 (lossy)" 
     .Xrange "-WG/2", "WG/2" 
     .Yrange "-LG/2", "LG/2" 
     .Zrange "0", "-h" 
     .Create
End With

'@ define material: Copper (annealed)

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Material
     .Reset
     .Name "Copper (annealed)"
     .Folder ""
     .FrqType "static"
     .Type "Normal"
     .SetMaterialUnit "Hz", "mm"
     .Epsilon "1"
     .Mu "1.0"
     .Kappa "5.8e+007"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .KappaM "0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstTanD"
     .DispModelEps "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .DispersiveFittingSchemeMu "Nth Order"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMu "False"
     .FrqType "all"
     .Type "Lossy metal"
     .SetMaterialUnit "GHz", "mm"
     .Mu "1.0"
     .Kappa "5.8e+007"
     .Rho "8930.0"
     .ThermalType "Normal"
     .ThermalConductivity "401.0"
     .SpecificHeat "390", "J/K/kg"
     .MetabolicRate "0"
     .BloodFlow "0"
     .VoxelConvection "0"
     .MechanicsType "Isotropic"
     .YoungsModulus "120"
     .PoissonsRatio "0.33"
     .ThermalExpansionRate "17"
     .Colour "1", "1", "0"
     .Wireframe "False"
     .Reflection "False"
     .Allowoutline "True"
     .Transparentoutline "False"
     .Transparency "0"
     .Create
End With

'@ define brick: component1:GND

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "GND" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-WG/2+z", "WG/2-z" 
     .Yrange "-LG/2+z", "-we" 
     .Zrange "-h", "-h-mt" 
     .Create
End With

'@ define brick: component1:Patch

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "Patch" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-WP/2", "WP/2" 
     .Yrange "-LP/2", "LP/2" 
     .Zrange "0", "mt" 
     .Create
End With

'@ activate local coordinates

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.ActivateWCS "local"

'@ activate global coordinates

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.ActivateWCS "global"

'@ define brick: component1:Microstrip

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "Microstrip" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-Wt/2", "Wt/2" 
     .Yrange "-LP/2", "-LP/2-14.3" 
     .Zrange "0", "mt" 
     .Create
End With

'@ activate local coordinates

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.ActivateWCS "local"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "Wt/2", "-LP/2", "0.0"

'@ define brick: component1:vacuum_right

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "vacuum_right" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "g" 
     .Yrange "0", "x0" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:vacuum_right

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Patch", "component1:vacuum_right"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "-Wt", "0.0", "0.0"

'@ define brick: component1:vacuum_left

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "vacuum_left" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-g", "0" 
     .Yrange "0", "x0" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:vacuum_left

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Patch", "component1:vacuum_left"

'@ pick face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.PickFaceFromId "component1:Microstrip", "3"

'@ define port:1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "1.6*5.23", "1.6*5.23"
  .YrangeAdd "0", "0"
  .ZrangeAdd "1.6", "1.6*5.23"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define time domain solver parameters

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Mesh.SetCreator "High Frequency" 

With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-40"
     .MeshAdaption "False"
     .AutoNormImpedance "False"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .RunDiscretizerOnly "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With

'@ set PBA version

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Discretizer.PBAVersion "2023101624"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "-WP/2", "0.0", "0.0"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "0.0", "WP/2", "0.0"

'@ define brick: component1:test_lewo

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "test_lewo" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "0", "-k" 
     .Yrange "0", "-12" 
     .Zrange "0", "0" 
     .Create
End With

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "WP+Wt", "0.0", "0.0"

'@ define brick: component1:test_prawo

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "test_prawo" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "0", "k" 
     .Yrange "0", "-12" 
     .Zrange "0", "0" 
     .Create
End With

'@ activate global coordinates

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.ActivateWCS "global"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-19", "-17" 
     .Yrange "-1", "1" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "17", "19" 
     .Yrange "-1", "1" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ delete shapes

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid1" 
Solid.Delete "component1:solid2"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-5", "-1" 
     .Yrange "-22-j", "-18" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ delete shape: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid1"

'@ delete shapes

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:test_lewo" 
Solid.Delete "component1:test_prawo"

'@ activate local coordinates

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.ActivateWCS "local"

'@ set wcs properties

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With WCS
     .SetNormal "0", "0", "1"
     .SetOrigin "18.35", "3.5", "0"
     .SetUVector "1", "0", "0"
End With

'@ align wcs with face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.ForceNextPick 
Pick.PickFaceFromId "component1:Patch", "21" 
WCS.AlignWCSWithSelected "Face"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "-WP/2", "-LP/2", "0.0"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "TW" 
     .Yrange "0", "TL" 
     .Zrange "-Mt", "0" 
     .Create
End With

'@ delete shape: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid1"

'@ align wcs with face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.ForceNextPick 
Pick.PickFaceFromId "component1:DIELECTRIC", "1" 
WCS.AlignWCSWithSelected "Face"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "-WP/2", "-LP/2", "0.0"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "TW" 
     .Yrange "0", "TL" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ delete shape: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid1"

'@ define brick: component1:vac1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "vac1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "TW" 
     .Yrange "0", "TL" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:vac1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Patch", "component1:vac1"

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "0.0", "LP", "0.0"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "TW" 
     .Yrange "0", "-TL" 
     .Zrange "0", "0" 
     .Create
End With

'@ delete shape: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "TW" 
     .Yrange "0", "-TL" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Patch", "component1:solid1"

'@ delete monitor: farfield (f=2800)

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Monitor 
     .Delete "farfield (f=2800)" 
End With

'@ define farfield monitor: farfield (broadband)

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (broadband)" 
     .Domain "Time" 
     .Accuracy "1e-3" 
     .Samples "21" 
     .FieldType "Farfield" 
     .TransientFarfield "False" 
     .ExportFarfieldSource "False" 
     .Create 
End With

'@ align wcs with face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.ForceNextPick 
Pick.PickFaceFromId "component1:DIELECTRIC", "1" 
WCS.AlignWCSWithSelected "Face"

'@ define curve polygon: curve1:polygon1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Polygon 
     .Reset 
     .Name "polygon1" 
     .Curve "curve1" 
     .Point "-20", "10" 
     .LineTo "-25", "-0" 
     .LineTo "-20", "-5" 
     .Create 
End With

'@ define curve analytical: curve1:analytical1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With AnalyticalCurve
     .Reset 
     .Name "analytical1" 
     .Curve "curve1" 
     .LawX "t" 
     .LawY "s*exp(r*t)" 
     .LawZ "0" 
     .ParameterRange "0", "q" 
     .Create
End With

'@ move wcs

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.MoveWCS "local", "0.0", "16.0", "0.0"

'@ delete curve item: curve1:analytical1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Curve.DeleteCurveItem "curve1", "analytical1"

'@ activate global coordinates

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
WCS.ActivateWCS "global"

'@ farfield plot options

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With FarfieldPlot 
     .Plottype "3D" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "5" 
     .Step2 "5" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "False" 
     .SymmetricRange "False" 
     .SetTimeDomainFF "False" 
     .SetFrequency "2800" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "True" 
     .ShowStructureProfile "True" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "True" 
     .AspectRatio "Free" 
     .ShowGridlines "True" 
     .InvertAxes "False", "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .IncludeUnitCellSidewalls "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .SetMaxReferenceMode "abs" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1.0" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAntennaType "unknown" 
     .Phistart "1.000000e+00", "0.000000e+00", "0.000000e+00" 
     .Thetastart "0.000000e+00", "0.000000e+00", "1.000000e+00" 
     .PolarizationVector "0.000000e+00", "1.000000e+00", "0.000000e+00" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+00 
     .Origin "bbox" 
     .Userorigin "0.000000e+00", "0.000000e+00", "0.000000e+00" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+00" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+01" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .ClearCuts 
     .AddCut "lateral", "0", "1"  
     .AddCut "lateral", "90", "1"  
     .AddCut "polar", "90", "1"  

     .StoreSettings
End With

'@ delete curve: curve1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Curve.DeleteCurve "curve1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-17", "-12" 
     .Yrange "-13.6", "-8.6" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Patch, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Patch", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-17", "-12" 
     .Yrange "-8.6", "-8.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-17.06", "-17" 
     .Yrange "-13.6", "-8.54" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Patch, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Patch", "component1:solid1"

'@ boolean add shapes: component1:Patch, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Patch", "component1:solid2"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-17", "-12" 
     .Yrange "8", "13.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Patch, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Patch", "component1:solid1"

'@ boolean add shapes: component1:Microstrip, component1:Patch

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:Patch"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-15", "-3" 
     .Yrange "-12", "-10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "3", "15" 
     .Yrange "-12", "-10" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-20", "-4" 
     .Yrange "-20", "-18" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "4", "20" 
     .Yrange "-20", "-18" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-14", "-12" 
     .Yrange "-26", "-20" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid4

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "12", "14" 
     .Yrange "-26", "-20" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid5

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid5" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-25", "-20" 
     .Yrange "-20", "15" 
     .Zrange "0", "mt" 
     .Create
End With

'@ define brick: component1:solid6

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid6" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "20", "26" 
     .Yrange "-20", "16" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid7

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid7" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-25", "26" 
     .Yrange "16", "20" 
     .Zrange "0", "mt" 
     .Create
End With

'@ define brick: component1:solid8

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid8" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-25", "-20" 
     .Yrange "15", "16" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid1", "component1:solid2"

'@ boolean add shapes: component1:solid3, component1:solid4

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid3", "component1:solid4"

'@ boolean add shapes: component1:solid5, component1:solid6

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid5", "component1:solid6"

'@ boolean add shapes: component1:solid7, component1:solid8

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid7", "component1:solid8"

'@ boolean add shapes: component1:solid1, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid1", "component1:solid3"

'@ boolean add shapes: component1:solid5, component1:solid7

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid5", "component1:solid7"

'@ boolean add shapes: component1:solid1, component1:solid5

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid1", "component1:solid5"

'@ delete port: port1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.PickFaceFromId "component1:Microstrip", "84"

'@ define port:1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "1.6*5.23", "1.6*5.23"
  .YrangeAdd "0", "0"
  .ZrangeAdd "1.6", "1.6*5.23"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define frequency range

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solver.FrequencyRange "100", "4000"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-14", "14" 
     .Yrange "-2", "10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-16", "14" 
     .Yrange "-2", "10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-10", "-3" 
     .Yrange "-12", "-10" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "3", "10" 
     .Yrange "-12", "-10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-15", "-10" 
     .Yrange "-12", "-10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "10", "15" 
     .Yrange "-12", "-10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-14", "-12" 
     .Yrange "-26", "-20" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "12", "14" 
     .Yrange "-26", "-20" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-3", "2" 
     .Yrange "13", "16" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ delete shape: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-3", "3" 
     .Yrange "13", "16" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-10", "-7.5" 
     .Yrange "-18", "-13.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "7.5", "10" 
     .Yrange "-18", "-13.5" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-20", "-17" 
     .Yrange "-2", "2" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "17", "20" 
     .Yrange "-2", "2" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "20", "24" 
     .Yrange "2", "16" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-23", "-19" 
     .Yrange "2", "16" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-23", "-19" 
     .Yrange "-18", "-2" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "19", "23.5" 
     .Yrange "-18", "-2" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-23", "-17" 
     .Yrange "5", "10" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "17", "24" 
     .Yrange "5", "10" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid3"

'@ define brick: component1:GND#2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "GND#2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-32", "32" 
     .Yrange "LG/2", "LG/2-we" 
     .Zrange "-h", "-h-Mt" 
     .Create
End With

'@ delete port: port1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.PickFaceFromId "component1:Microstrip", "162"

'@ define port:1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "1.6*6.87", "1.6*6.87"
  .YrangeAdd "0", "0"
  .ZrangeAdd "1.6", "1.6*6.87"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-6", "-4" 
     .Yrange "-14", "-0" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "4", "6" 
     .Yrange "-14", "0" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ delete shape: component1:GND#2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:GND#2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-10", "-7.5" 
     .Yrange "-16.5", "-15.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "7.5", "10" 
     .Yrange "-16.5", "-15.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-6", "6" 
     .Yrange "-0", "2" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-5.5", "-4" 
     .Yrange "-24.5", "-20" 
     .Zrange "mt", "0" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "4", "5.5" 
     .Yrange "-24.5", "-20" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ define brick: component1:solid4

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-4", "4" 
     .Yrange "-24.5", "-23" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:solid2, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid2", "component1:solid3"

'@ boolean add shapes: component1:solid2, component1:solid4

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid2", "component1:solid4"

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-23", "-17" 
     .Yrange "5", "10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "17", "24" 
     .Yrange "5", "10" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-6", "6" 
     .Yrange "2", "5" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-6", "6" 
     .Yrange "2", "5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid3"

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-6", "6" 
     .Yrange "2", "5" 
     .Zrange "-0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:solid2, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid2", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-25", "-5.5" 
     .Yrange "-24.5", "-22.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ delete shape: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-25", "-5.5" 
     .Yrange "-24.5", "-22.5" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-25", "-22" 
     .Yrange "-23", "-20" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "5.5", "26.2" 
     .Yrange "-24.5", "-22.6" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "23", "26" 
     .Yrange "-24.5", "-20" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ define brick: component1:solid4

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "26", "26.2" 
     .Yrange "-24.5", "-22.6" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:solid2, component1:solid4

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid2", "component1:solid4"

'@ boolean add shapes: component1:solid2, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:solid2", "component1:solid3"

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-14", "-13" 
     .Yrange "-14", "14" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ delete shape: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-15", "-14" 
     .Yrange "-12", "12" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "14", "16" 
     .Yrange "-12", "12" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-10", "10" 
     .Yrange "8", "10" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "WG/2", "25" 
     .Yrange "-16", "28" 
     .Zrange "-h-Mt", "-h" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-32", "-24" 
     .Yrange "-16", "28" 
     .Zrange "-h-Mt", "-h" 
     .Create
End With

'@ boolean add shapes: component1:GND, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:GND", "component1:solid2"

'@ boolean add shapes: component1:GND, component1:solid3

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:GND", "component1:solid3"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-10", "10" 
     .Yrange "8", "10" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-15", "-13" 
     .Yrange "-12", "12" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-22", "-5.5" 
     .Yrange "-26", "-22" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "5.5", "23" 
     .Yrange "-25", "-21" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid2"

'@ boolean add shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-25", "-22" 
     .Yrange "-25", "-20" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "22", "26" 
     .Yrange "-25", "-20" 
     .Zrange "0", "Mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define cylinder: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Cylinder 
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .OuterRadius "14" 
     .InnerRadius "1" 
     .Axis "z" 
     .Zrange "-h", "-h-Mt" 
     .Xcenter "0" 
     .Ycenter "0" 
     .Segments "0" 
     .Create 
End With

'@ delete port: port1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.PickFaceFromId "component1:Microstrip", "410"

'@ define port:1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "1.6*6.87", "1.6*6.87"
  .YrangeAdd "0", "0"
  .ZrangeAdd "1.6", "1.6*6.87"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ delete shape: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-20", "20" 
     .Yrange "-10", "28" 
     .Zrange "-h-Mt", "-h" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-10", "10" 
     .Yrange "16", "28" 
     .Zrange "-h-Mt", "-h" 
     .Create
End With

'@ delete shape: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Delete "component1:solid2"

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-12", "12" 
     .Yrange "16", "28" 
     .Zrange "-h-Mt", "-h" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ boolean add shapes: component1:GND, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:GND", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-23", "-3" 
     .Yrange "15", "18" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "3", "24" 
     .Yrange "14", "18" 
     .Zrange "Mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-2", "4" 
     .Yrange "-10", "16" 
     .Zrange "-h", "-h-mt" 
     .Create
End With

'@ boolean subtract shapes: component1:GND, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:GND", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-5.5", "-4" 
     .Yrange "-23", "-20" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "4", "5.5" 
     .Yrange "-23", "-20" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-7.5", "-4" 
     .Yrange "-20", "-18" 
     .Zrange "mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "4", "7.5" 
     .Yrange "-20", "-18" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-2", "2" 
     .Yrange "6", "19" 
     .Zrange "mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-14", "-4" 
     .Yrange "8", "12" 
     .Zrange "mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "4", "16" 
     .Yrange "8", "12" 
     .Zrange "mt", "0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "10", "14" 
     .Yrange "-0", "8" 
     .Zrange "mt", "-0" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-13", "-10" 
     .Yrange "0", "8" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-13", "-10" 
     .Yrange "-7", "-0" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "10", "14" 
     .Yrange "-7", "-0" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-24", "24" 
     .Yrange "20", "22" 
     .Zrange "mt", "0" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid1"

'@ delete port: port1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Pick.PickFaceFromId "component1:Microstrip", "494"

'@ define port:1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "1.6*6.87", "1.6*6.87"
  .YrangeAdd "0", "0"
  .ZrangeAdd "1.6", "1.6*6.87"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-24", "-16" 
     .Yrange "-18", "18" 
     .Zrange "mt", "0" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "17", "24" 
     .Yrange "-18", "18" 
     .Zrange "mt", "0" 
     .Create
End With

'@ boolean add shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid1"

'@ boolean add shapes: component1:Microstrip, component1:solid2

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Add "component1:Microstrip", "component1:solid2"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-26", "-18" 
     .Yrange "-20", "22" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "20", "26" 
     .Yrange "-20", "22" 
     .Zrange "0", "mt" 
     .Create
End With

'@ boolean subtract shapes: component1:Microstrip, component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
Solid.Subtract "component1:Microstrip", "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2024.1|33.0.1|20231016[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-8", "6" 
     .Yrange "20", "28" 
     .Zrange "-h", "-h-mt" 
     .Create
End With

