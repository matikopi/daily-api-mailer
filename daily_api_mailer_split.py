import json
import base64
import os
import subprocess
import time
from datetime import datetime
import urllib3

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

def send_individual_file(filename, description, file_number, total_files):
    """Send a single Excel file via email."""
    current_date = datetime.now()
    
    if not os.path.exists(filename):
        print(f"✗ File not found: {filename}")
        return False
    
    # Get file size for display
    file_size = os.path.getsize(filename) / (1024 * 1024)  # MB
    
    encoded_content = encode_file_to_base64(filename)
    if not encoded_content:
        print(f"✗ Failed to encode {filename}")
        return False
    
    # Create email content
    email_body = f"""Noga ISO Daily Data Report - Part {file_number} of {total_files}

Generated on: {current_date.strftime('%Y-%m-%d %H:%M:%S')}

This email contains: {description}
File size: {file_size:.1f} MB

Data from Israel's Independent System Operator (Noga ISO)
Source: https://apim-api.noga-iso.co.il/

The Excel file contains two sheets:
- Most Recent Day: Yesterday's complete data
- Historical Data: Full available history from earliest date

This is part {file_number} of {total_files} emails with your daily data report.
"""

    # SendGrid email payload
    payload = {
        "personalizations": [
            {
                "to": [{"email": RECIPIENT_EMAIL}]
            }
        ],
        "from": {"email": RECIPIENT_EMAIL},
        "subject": f"Noga ISO Data - {description} - {current_date.strftime('%Y-%m-%d')}",
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

def send_all_files_separately():
    """Send each Excel file in a separate email."""
    # Expected Excel files with descriptions
    excel_files = [
        ('production_mix.xlsx', 'Production Mix Data'),
        ('co2_data.xlsx', 'CO2 Emissions Data'),
        ('demand_data.xlsx', 'Electricity Demand Data'),
        ('smp_data.xlsx', 'System Marginal Price Data')
    ]
    
    successful_sends = 0
    total_files = len([f for f, _ in excel_files if os.path.exists(f)])
    
    if total_files == 0:
        print("No Excel files found to send!")
        return False
    
    print(f"Sending {total_files} files in separate emails...\n")
    
    file_number = 1
    for filename, description in excel_files:
        if os.path.exists(filename):
            success = send_individual_file(filename, description, file_number, total_files)
            if success:
                successful_sends += 1
            file_number += 1
            
            # Brief pause between emails
            time.sleep(2)
    
    print(f"\n✓ Successfully sent {successful_sends}/{total_files} files")
    return successful_sends == total_files

def main():
    """Main function to run API scripts and send emails."""
    print("=" * 60)
    print("NOGA ISO DAILY API MAILER (SPLIT VERSION)")
    print("=" * 60)
    
    start_time = datetime.now()
    
    # Step 1: Run all API fetch scripts
    run_api_scripts()
    
    # Step 2: Wait a moment for files to be fully written
    print("Waiting for files to be written...")
    time.sleep(5)
    
    # Step 3: Send files in separate emails
    success = send_all_files_separately()
    
    end_time = datetime.now()
    duration = end_time - start_time
    
    print("\n" + "=" * 60)
    if success:
        print("✓ DAILY MAILER COMPLETED SUCCESSFULLY")
    else:
        print("✗ DAILY MAILER COMPLETED WITH ERRORS")
    
    print(f"Total runtime: {duration}")
    print(f"Completed at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

if __name__ == "__main__":
    main() 