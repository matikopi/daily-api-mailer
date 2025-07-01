import json
import base64
import os
import subprocess
import time
import sys
from datetime import datetime, timedelta
import urllib3
import pandas as pd

# Initialize HTTP pool manager
http = urllib3.PoolManager()

# SendGrid API Key - consider moving to environment variable
SENDGRID_API_KEY = "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"
RECIPIENT_EMAIL = "matikopi@gmail.com"

def run_api_scripts():
    """Run all API fetch scripts and wait for completion."""
    scripts = [
        'fetch_production_mix.py',
        'fetch_co2_data.py', 
        'fetch_demand_data.py',
        'fetch_smp_data.py'
    ]
    
    print("Starting API data collection...")
    
    for script in scripts:
        if os.path.exists(script):
            print(f"Running {script}...")
            try:
                result = subprocess.run(['python', script], 
                                      capture_output=True, 
                                      text=True, 
                                      timeout=1800)  # 30 minute timeout
                
                if result.returncode == 0:
                    print(f"✓ {script} completed successfully")
                else:
                    print(f"✗ {script} failed with error: {result.stderr}")
                    
            except subprocess.TimeoutExpired:
                print(f"✗ {script} timed out after 30 minutes")
            except Exception as e:
                print(f"✗ Error running {script}: {str(e)}")
        else:
            print(f"✗ Script {script} not found")
    
    print("API data collection completed.\n")

def create_daily_summary_excel():
    """Create a single Excel file with yesterday's data from all 4 APIs in separate tabs."""
    current_date = datetime.now()
    yesterday = current_date - timedelta(days=1)
    yesterday_str = yesterday.strftime('%d-%m-%Y')
    
    output_filename = f"noga_daily_report_{current_date.strftime('%Y-%m-%d')}.xlsx"
    
    excel_files = [
        ('production_mix.xlsx', 'Production Mix'),
        ('co2_data.xlsx', 'CO2 Emissions'),
        ('demand_data.xlsx', 'Demand'),
        ('smp_data.xlsx', 'SMP Pricing')
    ]
    
    print("Creating daily summary Excel file...")
    
    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            sheets_created = 0
            
            for source_file, sheet_name in excel_files:
                if os.path.exists(source_file):
                    try:
                        # Read the "Most Recent Day" sheet from each API's Excel file
                        df = pd.read_excel(source_file, sheet_name='Most Recent Day')
                        
                        # Write to the combined file
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheets_created += 1
                        print(f"✓ Added {sheet_name} data ({len(df)} rows)")
                        
                    except Exception as e:
                        print(f"✗ Error processing {source_file}: {str(e)}")
                else:
                    print(f"✗ Source file not found: {source_file}")
        
        if sheets_created > 0:
            file_size = os.path.getsize(output_filename) / (1024 * 1024)  # MB
            print(f"✓ Created {output_filename} with {sheets_created} sheets ({file_size:.1f} MB)")
            return output_filename
        else:
            print("✗ No data sheets created")
            return None
            
    except Exception as e:
        print(f"✗ Error creating summary Excel: {str(e)}")
        return None

def encode_file_to_base64(file_path):
    """Read Excel file and encode to base64."""
    try:
        with open(file_path, 'rb') as file:
            file_content = file.read()
            encoded_content = base64.b64encode(file_content).decode('utf-8')
            return encoded_content
    except Exception as e:
        print(f"Error encoding {file_path}: {str(e)}")
        return None

def send_daily_summary(summary_filename):
    """Send the daily summary Excel file via email."""
    current_date = datetime.now()
    
    if not os.path.exists(summary_filename):
        print(f"✗ Summary file not found: {summary_filename}")
        return False
    
    file_size = os.path.getsize(summary_filename) / (1024 * 1024)  # MB
    
    encoded_content = encode_file_to_base64(summary_filename)
    if not encoded_content:
        print(f"✗ Failed to encode {summary_filename}")
        return False
    
    # Create email content
    email_body = f"""Noga ISO Daily Data Report

Generated on: {current_date.strftime('%Y-%m-%d %H:%M:%S')}

This email contains yesterday's electricity market data from Israel's Independent System Operator (Noga ISO) in a single Excel file with 4 tabs:

• Production Mix: Energy generation by source (5-minute intervals)
• CO2 Emissions: Carbon dioxide emissions data (5-minute intervals)  
• Demand: Electricity demand forecasts (30-minute intervals)
• SMP Pricing: System Marginal Pricing - constrained & unconstrained (30-minute intervals)

File size: {file_size:.1f} MB
Data source: Noga ISO APIs (https://apim-api.noga-iso.co.il/)

Note: This contains only yesterday's data. To receive historical data files, 
run the script with --historical flag.
"""

    # SendGrid email payload
    payload = {
        "personalizations": [
            {
                "to": [{"email": RECIPIENT_EMAIL}]
            }
        ],
        "from": {"email": RECIPIENT_EMAIL},
        "subject": f"Noga ISO Daily Report - {current_date.strftime('%Y-%m-%d')}",
        "content": [{
            "type": "text/plain",
            "value": email_body
        }],
        "attachments": [{
            "content": encoded_content,
            "filename": summary_filename,
            "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "disposition": "attachment"
        }]
    }
    
    # Send email via SendGrid
    url = "https://api.sendgrid.com/v3/mail/send"
    
    try:
        print(f"Sending daily summary ({file_size:.1f} MB)...")
        encoded_data = json.dumps(payload).encode('utf-8')
        
        response = http.request(
            'POST',
            url,
            body=encoded_data,
            headers={
                'Authorization': f'Bearer {SENDGRID_API_KEY}',
                'Content-Type': 'application/json'
            }
        )
        
        if response.status == 202:
            print(f"✓ Daily summary sent successfully to {RECIPIENT_EMAIL}!")
            return True
        else:
            print(f"✗ SendGrid error {response.status} for daily summary")
            print(f"Response: {response.data.decode()}")
            return False
            
    except Exception as e:
        print(f"✗ Error sending daily summary: {str(e)}")
        return False

def send_historical_file(filename, description, file_number, total_files):
    """Send a single historical Excel file via email."""
    current_date = datetime.now()
    
    if not os.path.exists(filename):
        print(f"✗ File not found: {filename}")
        return False
    
    file_size = os.path.getsize(filename) / (1024 * 1024)  # MB
    
    encoded_content = encode_file_to_base64(filename)
    if not encoded_content:
        print(f"✗ Failed to encode {filename}")
        return False
    
    # Create email content
    email_body = f"""Noga ISO Historical Data - Part {file_number} of {total_files}

Generated on: {current_date.strftime('%Y-%m-%d %H:%M:%S')}

This email contains: {description} (Full Historical Data)
File size: {file_size:.1f} MB

Data from Israel's Independent System Operator (Noga ISO)
Source: https://apim-api.noga-iso.co.il/

This file contains complete historical data from the earliest available date.
This is part {file_number} of {total_files} historical data emails.
"""

    # SendGrid email payload
    payload = {
        "personalizations": [
            {
                "to": [{"email": RECIPIENT_EMAIL}]
            }
        ],
        "from": {"email": RECIPIENT_EMAIL},
        "subject": f"Noga ISO Historical - {description} - {current_date.strftime('%Y-%m-%d')}",
        "content": [{
            "type": "text/plain",
            "value": email_body
        }],
        "attachments": [{
            "content": encoded_content,
            "filename": filename,
            "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "disposition": "attachment"
        }]
    }
    
    # Send email via SendGrid
    url = "https://api.sendgrid.com/v3/mail/send"
    
    try:
        print(f"Sending {filename} ({file_size:.1f} MB)...")
        encoded_data = json.dumps(payload).encode('utf-8')
        
        response = http.request(
            'POST',
            url,
            body=encoded_data,
            headers={
                'Authorization': f'Bearer {SENDGRID_API_KEY}',
                'Content-Type': 'application/json'
            }
        )
        
        if response.status == 202:
            print(f"✓ {filename} sent successfully!")
            return True
        else:
            print(f"✗ SendGrid error {response.status} for {filename}")
            print(f"Response: {response.data.decode()}")
            return False
            
    except Exception as e:
        print(f"✗ Error sending {filename}: {str(e)}")
        return False

def send_historical_files():
    """Send historical Excel files in separate emails."""
    excel_files = [
        ('production_mix.xlsx', 'Production Mix Data'),
        ('co2_data.xlsx', 'CO2 Emissions Data'),
        ('demand_data.xlsx', 'Electricity Demand Data'),
        ('smp_data.xlsx', 'System Marginal Price Data')
    ]
    
    successful_sends = 0
    total_files = len([f for f, _ in excel_files if os.path.exists(f)])
    
    if total_files == 0:
        print("No historical Excel files found to send!")
        return False
    
    print(f"Sending {total_files} historical files in separate emails...\n")
    
    file_number = 1
    for filename, description in excel_files:
        if os.path.exists(filename):
            success = send_historical_file(filename, description, file_number, total_files)
            if success:
                successful_sends += 1
            file_number += 1
            
            # Brief pause between emails
            time.sleep(2)
    
    print(f"\n✓ Successfully sent {successful_sends}/{total_files} historical files")
    return successful_sends == total_files

def main():
    """Main function with options for daily or historical data."""
    print("=" * 60)
    print("NOGA ISO API MAILER V2")
    print("=" * 60)
    
    # Check for command line arguments
    send_historical = "--historical" in sys.argv
    skip_fetch = "--skip-fetch" in sys.argv
    
    start_time = datetime.now()
    
    # Step 1: Run API scripts (unless skipped)
    if not skip_fetch:
        run_api_scripts()
        print("Waiting for files to be written...")
        time.sleep(5)
    else:
        print("Skipping API fetch (using existing files)...\n")
    
    # Step 2: Create and send daily summary
    print("Creating daily summary...")
    summary_filename = create_daily_summary_excel()
    
    if summary_filename:
        daily_success = send_daily_summary(summary_filename)
    else:
        daily_success = False
        print("✗ Failed to create daily summary")
    
    # Step 3: Send historical files if requested
    historical_success = True
    if send_historical:
        print("\nSending historical data files...")
        historical_success = send_historical_files()
    else:
        print("\nSkipping historical files (use --historical to include)")
    
    end_time = datetime.now()
    duration = end_time - start_time
    
    print("\n" + "=" * 60)
    if daily_success and historical_success:
        print("✓ MAILER COMPLETED SUCCESSFULLY")
    else:
        print("✗ MAILER COMPLETED WITH ERRORS")
    
    print(f"Total runtime: {duration}")
    print(f"Completed at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("\nUsage examples:")
    print("  py daily_api_mailer_v2.py                    # Daily summary only")
    print("  py daily_api_mailer_v2.py --historical       # Daily + historical")
    print("  py daily_api_mailer_v2.py --skip-fetch       # Use existing files")
    print("=" * 60)

if __name__ == "__main__":
    main() 