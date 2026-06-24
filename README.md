# Flyer Stack Machining

Tools for generating a LightBurn `.lbrn2` Flyer Stack.

## Run the tool

```bash
python cfstack.py [config.json]
```

## Repository layout

- `cfstack.py`: main script.
- `inputs/configs/`: run-config JSON inputs.
- `inputs/excel/`: laser parameter excel sheets.
- `inputs/igsn/`: foil/material IGSN.
- `inputs/templates/`: LightBurn template files and sidecar JSON metadata.
- `config_app.html`: browser form for producing config JSON files.
- `output/`: generated outputs written to `output/<igsn>/<stack_id>/`.