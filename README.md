# Flyer-Cut-Opt
A host of projects designed to optimize the machining of flyers for HELIX with AIMD. 

## Execution
To execute create flyer stack, `cfstack.py`, provide the required positional arguments:
`python cfstack.py {config-json}`

The excel-sheet expects columns `maxPower`, `QPulseWidth`, `speed`,	`frequency`, and `numPasses` (with these labels in row1). The script should automatically parse a properly formatted template lightburn file specificated in the config.

## Contents
`cfstack.py`: script accepts a excel sheet of laser parameters and JSON and compiles a `.lbrn2` lightburn project file according to the provided specifications.

`config_app.html`: html form that helps create a JSON of run parameters for the script. 

`LB-TEMPLATES`: directory of formatted lightburn files for the script.

`IGSN-CONFIGS`: directory of foil/material metadata for data linking. 

`EXCEL`: directory of formatted excel sheets for the script

`CONFIGS`: directory of formatted config json for the script. 
