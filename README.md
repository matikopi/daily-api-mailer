# Noga ISO API Data Extraction

This repository contains Python scripts to fetch electricity market data from Israel's Independent System Operator (Noga ISO) APIs and export them to Excel files.

## Overview

The scripts pull data from four different Noga ISO APIs and generate separate Excel files with both recent (yesterday) and historical data:

- **Production Mix API**: Energy generation by source type
- **CO2 API**: Carbon dioxide emissions data
- **Demand API**: Electricity demand forecasts and actuals  
- **SMP API**: System Marginal Price (wholesale electricity pricing)

## Data Coverage

| API | Data Frequency | Earliest Date | Variables |
|-----|---------------|---------------|-----------|
| Production Mix | Every 5 minutes | Feb 22, 2023 | Coal, Natural Gas, Solar, Wind, etc. |
| CO2 | Every 5 minutes | Mar 23, 2023 | CO2 emissions data |
| Demand | Every 30 minutes | Dec 27, 2022 | Electricity demand forecasts |
| SMP | Every 30 minutes | Dec 27, 2021 | Constrained & Unconstrained prices |

## API Details

### Production Mix API
**Data includes energy generation by source:**
- Coal
- Natural Gas  
- Solar (thermal and photovoltaic)
- Wind
- Mazut (heavy fuel oil)
- Pumped Storage
- Bio Gas
- Other sources
- Actual Demand

### SMP (System Marginal Price) API
**Includes two price types:**
- **Constrained SMP**: Price reflecting transmission network constraints and congestion
- **Unconstrained SMP**: Theoretical price without transmission limitations

**What is Constrained vs Unconstrained Pricing?**

In electricity markets, the difference between constrained and unconstrained pricing reflects transmission network limitations:

- **Unconstrained Price**: The theoretical wholesale electricity price if there were infinite transmission capacity. This represents the pure economic merit order of generation.

- **Constrained Price**: The actual market price that accounts for transmission line limits, network congestion, and physical constraints. When transmission lines reach capacity, cheaper generators may be "constrained off" and more expensive local generators must run instead, raising the price.

**Example**: If cheap wind power in northern Israel cannot reach demand centers in the south due to transmission constraints, expensive gas plants near the demand must run instead, creating a higher "constrained" price than the theoretical "unconstrained" price.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

Run individual scripts to fetch data for each API:

```bash
# Production Mix data
python fetch_production_mix.py

# CO2 emissions data  
python fetch_co2_data.py

# Demand data
python fetch_demand_data.py

# SMP pricing data
python fetch_smp_data.py
```

Each script will create an Excel file with two sheets:
- **Most Recent Day**: Yesterday's full-day data
- **Historical Data**: Complete history from earliest available date

## Output Files

- `production_mix.xlsx`
- `co2_data.xlsx` 
- `demand_data.xlsx`
- `smp_data.xlsx`

## Email Configuration (daily_api_mailer_v2.py)

For sending automated email reports, set these environment variables:

```bash
# Gmail SMTP Configuration
set SENDER_EMAIL=your_gmail@gmail.com
set GMAIL_APP_PASSWORD=your_16_char_app_password

# Multiple recipients (comma-separated, no spaces needed)
set RECIPIENT_EMAILS=email1@company.com,email2@company.com,email3@company.com

# Single recipient example:
set RECIPIENT_EMAILS=matikopi@gmail.com
```

**Gmail Setup Requirements:**
1. Enable 2-Factor Authentication on your Gmail account
2. Generate a 16-character App Password (not your regular password)
3. Use the App Password in `GMAIL_APP_PASSWORD`

**Usage Examples:**
```bash
# Send daily summary only
py daily_api_mailer_v2.py

# Send daily summary + historical files
py daily_api_mailer_v2.py --historical

# Use existing files (skip API fetch)
py daily_api_mailer_v2.py --skip-fetch
```

## Configuration

The scripts automatically try both API endpoints:
- Primary: `https://apim-api.noga-iso.co.il/`
- Fallback: `https://noga-apim-prod.azure-api.net/`

API keys are configured per script. In production, consider using environment variables instead of hardcoded keys.

## API Documentation

Based on Noga ISO's developer portal at https://apim-api.noga-iso.co.il/

**Request Format:**
```json
{
    "fromDate": "DD-MM-YYYY",
    "toDate": "DD-MM-YYYY"
}
```

**Authentication:**
```
Ocp-Apim-Subscription-Key: YOUR_API_KEY
```

## Energy Market Context

### Israel's Electricity Market
Israel operates a competitive electricity wholesale market managed by Noga ISO (Independent System Operator). The market includes:

- **Generation**: Mix of natural gas, renewables, and conventional sources
- **Transmission**: High-voltage network managed by Noga ISO
- **System Operation**: Real-time balancing of supply and demand
- **Market Operation**: Wholesale electricity trading and pricing

### Key Market Concepts

**System Marginal Price (SMP)**: The wholesale electricity price that clears the market, typically set by the most expensive generator needed to meet demand.

**Locational Marginal Pricing**: Prices can vary by location due to transmission constraints, with constrained areas experiencing different prices than unconstrained areas.

**Real-time Dispatch**: The system operator continuously adjusts generation to match changing demand, with prices updating every few minutes.

## Data Applications

This data can be used for:
- **Energy Market Analysis**: Understanding price patterns and generation mix
- **Renewable Energy Studies**: Tracking solar/wind integration
- **Carbon Footprint Analysis**: Using CO2 emissions data
- **Load Forecasting**: Analyzing demand patterns
- **Trading Strategy**: Understanding price formation mechanisms

## Dependencies

- `pandas`: Data manipulation and Excel export
- `requests`: API calls
- `openpyxl`: Excel file creation
- `datetime`: Date handling
