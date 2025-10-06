# NorthData API Integration

## Setup

1. Create a virtual environment and install dependencies:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -e .
   ```

2. Provide your NorthData API key:

   **Linux / macOS (Bash):**

   ```bash
   export NORTHDATA_API_KEY=your_real_key
   ```

   **Windows (PowerShell):**

   ```powershell
   $Env:NORTHDATA_API_KEY = "your_real_key"
   ```

3. Run the integration:

   ```bash
   python cli.py --excel "Liste.xlsx" --sheet "Tabelle1" --start 3 --end 20 --name-col C --zip-col I --country-col J --source api
   ```

## Column Mapping

You can provide a YAML file to control how Excel columns map to the NorthData `CompanyRecord` fields. An example mapping is included at `mappings/example_mapping.yaml`.

Run the CLI with the `--mapping-yaml` flag to use the mapping file and automatically download additional documents:

```bash
python cli.py --excel "Liste.xlsx" --sheet "Tabelle1" --start 3 --end 10 --name-col C --zip-col I --country-col J --source api --mapping-yaml mappings/example_mapping.yaml --download-ad
```
