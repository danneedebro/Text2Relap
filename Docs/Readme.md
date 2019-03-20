# Main documentation for Text2Relap

## Recognised first words
### Settings
```
Title       Input file title
Include     Path to include folder
Timestep    A timestep control block
Refvol      A reference volume block
Cvformat    The format of cntrlvars (999 or 9999)
Tripformat  The format of trips. 0 = Default where variable trips 401-599 and logical trips 601-799
ForceCalc   The way to calculate forces
Comment     Verbose output with comments
CntrlvarNr  Sets the start of cntrlvar numbering. Can be used multiple times
```
### Components
```
Pipe                A pipe segment
Junction / Sngljun  A single junction component
Mtrvlv              A motor valve component
Srvvlv              A servo valve component
Inrvlv              An intertial swing check valve (non ideal check valve)
Chkvlv              A check valve component
Trpvlv              A trip valve
Tmdpvol             A time dependant volume
Snglvol             A single volume component
Tmdpjun             A time dependant junction component
Pump                A pump component
Custom              A custom component
TripVar             A variable trip
TripLog             A logical trip
*                   New flowpath comment
**                  Inline commment
```

### Hydrodynamic component initialisation
```
Init                Sets the (volume) initial conditions 
InitGas             Sets the (volume) initial conditions to non-condensible
```

### Misc input
```
IGNORE              Starts an ignore-input block
/IGNORE             Ends an ignore-input block
Triggerwords        [WORDREPLACE1;REPLACEWITH1;WORDREPLACE2;REPLACEWITH2]
Replacements        Same as above
```


# Special processes

## Connecting volumes

## Initialising components

## Custom components / Include files

## Built in functions

## Time-dependant input for certain components

## Loop (elevation) check
