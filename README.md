# Flyer Stack Machining

Tools for generating a LightBurn `.lbrn2` Flyer Stack.

## Run the tool

```bash
python cfstack.py [config.json]
```

## Repository layout

- `cfstack.py`: main script.
- `config_app.html`: browser form for creating run-config JSON files.
- `inputs/configs/`: directory for run-config JSON inputs.
- `inputs/excel/`: directory for laser parameter excel sheets.
- `inputs/igsn/`: directory for foil/material IGSN.
- `inputs/templates/`: directory for lightBurn template files with sidecar JSON.
- `output/`: generated outputs written to `output/<igsn>/<stack_id>/`.