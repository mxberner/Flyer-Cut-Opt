# Flyer-Cut-Opt
A host of projects designed to optimize the machining of flyers for HELIX with AIMD. 

## Execution
To execute create flyer stack, `cfstack.py`, provide the required positional arguments:
`python cfstack.py {excel-sheet} {config-json}`


The excel-sheet expects columns maxPower, QPulseWidth, speed,	frequency, and numPasses (with these labels in row1). The script should automatically parse a properly formatted template lightburn file specificated in the config.


## Contents
`cfstack.py`: script accepts a excel sheet of laser parameters and JSON and compiles a `.lbrn2` lightburn project file according to the provided specifications.

`config.json`: formatted JSON file containing miscellaneous parameters. 

`data.xslx`: excel sheet with rows that contain tested laser parameters.

`5x5stack.lbrn2`: formatted lightburn project template for use in script.

`6x2stack.lbrn2`: formatted lightburn project template for use in script.

`7x7stack.lbrn2`: row-wise formatted lightburn project template for use in script.
