import json
import base64
import os
import subprocess
import time
from datetime import datetime
import urllib3

# Initialize HTTP pool manager
http = urllib3.PoolManager()

# SendGrid API Key - set as environment variable or replace with your key
SENDGRID_API_KEY = os.getenv('SENDGRID_API_KEY', 'YOUR_SENDGRID_API_KEY_HERE')
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

def send_email_with_attachments():
    """Send email with all Excel files as attachments."""
    current_date = datetime.now()
    
    # Expected Excel files
    excel_files = [
        ('production_mix.xlsx', 'Production Mix Data'),
        ('co2_data.xlsx', 'CO2 Emissions Data'),
        ('demand_data.xlsx', 'Electricity Demand Data'),
        ('smp_data.xlsx', 'System Marginal Price Data')
    ]
    
    # Prepare attachments
    attachments = []
    files_found = 0
    
    for filename, description in excel_files:
        if os.path.exists(filename):
            encoded_content = encode_file_to_base64(filename)
            if encoded_content:
                attachments.append({
                    "content": encoded_content,
                    "filename": filename,
                    "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "disposition": "attachment"
                })
                files_found += 1
                print(f"✓ Attached {filename}")
            else:
                print(f"✗ Failed to encode {filename}")
        else:
            print(f"✗ File not found: {filename}")
    
    if files_found == 0:
        print("No Excel files found to send!")
        return False
    
    # Create email content
    email_body = f"""Daily Noga ISO API Data Report

Generated on: {current_date.strftime('%Y-%m-%d %H:%M:%S')}

This email contains electricity market data from Israel's Independent System Operator (Noga ISO):

• Production Mix Data: Energy generation by source (5-minute intervals)
• CO2 Emissions Data: Carbon dioxide emissions (5-minute intervals)  
• Demand Data: Electricity demand forecasts (30-minute intervals)
• SMP Data: System Marginal Pricing - constrained & unconstrained (30-minute intervals)

Each Excel file contains two sheets:
- Most Recent Day: Yesterday's complete data
- Historical Data: Full available history from earliest date

Files attached: {files_found}/4

Data sources: Noga ISO APIs (https://apim-api.noga-iso.co.il/)
"""

    # SendGrid email payload
    payload = {
        "personalizations": [
            {
                "to": [{"email": RECIPIENT_EMAIL}]
            }
        ],
        "from": {"email": RECIPIENT_EMAIL},
        "subject": f"Noga ISO Daily Data Report - {current_date.strftime('%Y-%m-%d')}",
        "content": [{
            "type": "text/plain",
            "value": email_body
        }],
        "attachments": attachments
    }
    
    # Send email via SendGrid
    url = "https://api.sendgrid.com/v3/mail/send"
    
    try:
        print("Sending email via SendGrid...")
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
        
        print(f"SendGrid Response Status: {response.status}")
        
        if response.status == 202:
            print(f"✓ Email sent successfully to {RECIPIENT_EMAIL}!")
            print(f"✓ {files_found} Excel files attached")
            return True
        else:
            print(f"✗ SendGrid API returned status code {response.status}")
            print(f"Response: {response.data.decode()}")
            return False
            
    except Exception as e:
        print(f"✗ Error sending email: {str(e)}")
        return False

def main():
    """Main function to run API scripts and send email."""
    print("=" * 60)
    print("NOGA ISO DAILY API MAILER")
    print("=" * 60)
    
    start_time = datetime.now()
    
    # Step 1: Run all API fetch scripts
    run_api_scripts()
    
    # Step 2: Wait a moment for files to be fully written
    print("Waiting for files to be written...")
    time.sleep(5)
    
    # Step 3: Send email with attachments
    success = send_email_with_attachments()
    
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