# LightBurn Project Templates
The main script programmatically parses formatted LightBurn `.lbrn2` project files for data extraction and templatting. 

Lightburn has a software limit of 30 layers (unique cut setting / set of laser parameters), which each can contain many shapes (objects cut). For this reason, some stack layouts cannot have a unique layer (cut setting) for every single flyer (shape), i.e. 7x7=49 > 30. As a result, there is some disconnect between physical flyers and flyer layers; specifically, multiple physical flyers often live on the same layer. 
For disambiguation, each lightburn project is packaged with a `.json` of the same name to associate a set of laser parameters with each shape (flyer) following a traceable coordinate system. 

## Creating a New Template
For preservation purposes, please do not edit the names or the content of the existing LightBurn templates, as their use in prior executions will be logged for traceability. Creating a new template from a copy of another is encouraged.

The script expects template project files to follow this structure:
1. "Flyer" layers, labelled "F#", starting from 1 to N. The numbers must be consecutive to be found. 
2. A text label with the value matching `template_placeholder_id` within the config. 
3. A `.json` file containing the row/col position and layer for each physical flyer shape.


## Expected Behavior
The script will inject the data from the excel sheet row_i into the respective flyer f_i. The script will replace the string of the textbox containing `template_placeholder_id` with `ID`.