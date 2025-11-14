# LightBurn Project Templates
The main script programmatically parses formatted LightBurn `.lbrn2` project files for data extraction and templatting. 

## Creating a New Template
For preservation purposes, please do not edit the names or the content of the existing LightBurn templates, as their use in prior executions will be logged for traceability. Creating a new template from a copy of another is encouraged.

The script expects template project files to follow this structure:
1. "Flyer" cuts, labelled "F#", starting from 1 to N. The numbers must be consecutive to be found. 
2. A text label with the value matching `template_placeholder_id` within the config. 


## Expected Behavior
The script will inject the data from the excel sheet row_i into the respective flyer f_i. The script will replace the string of the textbox containing `template_placeholder_id` with `ID`.