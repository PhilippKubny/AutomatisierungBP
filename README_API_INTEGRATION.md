# NorthData API Integration

## Setup

1. Create a virtual environment and install dependencies:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -e .
   ```

2. Provide your NorthData API key:

   ```bash
   export NORTHDATA_API_KEY=your_real_key
   ```

   ```powershell
   $Env:NORTHDATA_API_KEY = "your_real_key"
   ```

3. Run the integration:

   ```bash
   python cli.py --excel "Liste.xlsx" --sheet "Tabelle1" --start 3 --end 20 --name-col C --zip-col I --country-col J --source api
   ```
