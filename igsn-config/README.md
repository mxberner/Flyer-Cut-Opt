# Foil IGSN Configs
Foil configs define the default parameter set used for a specific metal foil across experimental trials. These files should not be edited since their use in prior executions will be logged for traceability. If any edits or duplicates are made, `version` and `timestamp` should be updated.

## Purpose
Foil configs serve three core functions:

**Default Parameter Management**
Provide defaults for control variables that appropriate for the metal. Ensure that repeated trials using the same foil share the same baseline operating conditions.

**Metadata Packaging**
Bundle key metadata relevant to the foil, such as IGSN identifiers, material type and thickness, and logging control variables. 