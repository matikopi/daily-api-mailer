import json
import base64
import os
import subprocess
import time
import sys
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
import pandas as pd

# Gmail SMTP settings - set as environment variables
SMTP_USER = os.getenv('SENDER_EMAIL', 'your_gmail@gmail.com')
SMTP_PASS = os.getenv('GMAIL_APP_PASSWORD', 'your_app_password_here')

# Recipients can be multiple emails separated by commas
RECIPIENT_EMAILS_STR = os.getenv('RECIPIENT_EMAILS', 'matikopi@gmail.com')
RECIPIENT_EMAILS = [email.strip() for email in RECIPIENT_EMAILS_STR.split(',') if email.strip()]

# Print recipient configuration at startup
def print_email_config():
    """Print the current email configuration."""
    print(f"Sender: {SMTP_USER}")
    print(f"Recipients ({len(RECIPIENT_EMAILS)}): {', '.join(RECIPIENT_EMAILS)}")
    print()

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
    yesterday_str = yesterday.strftime('%Y-%m-%d')
    
    output_filename = f"NOGA Daily Report {yesterday_str}.xlsx"
    
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

def send_daily_summary(summary_filename):
    """Send the daily summary Excel file via Gmail SMTP."""
    current_date = datetime.now()
    yesterday = current_date - timedelta(days=1)
    yesterday_str = yesterday.strftime('%Y-%m-%d')
    
    if not os.path.exists(summary_filename):
        print(f"✗ Summary file not found: {summary_filename}")
        return False
    
    file_size = os.path.getsize(summary_filename) / (1024 * 1024)  # MB
    
    # Create email message
    msg = EmailMessage()
    msg["Subject"] = f"NOGA Daily Report {yesterday_str}"
    msg["From"] = SMTP_USER
    msg["To"] = ', '.join(RECIPIENT_EMAILS)
    
    # Email body
    generate_str = current_date.strftime('%B %-d, %Y at %H:%M') if os.name != 'nt' else current_date.strftime('%B %d, %Y at %H:%M').lstrip('0')
    email_body = f"""NOGA Israel - Daily Electricity Market Report

Generated on: {generate_str}

Attached is the daily electricity market data from Israel's Independent System Operator (Noga ISO), covering the previous day's activity (report for {yesterday_str}). The Excel file includes the following tabs:

• Production Mix – Energy generation by source (5-minute intervals)
• CO₂ Emissions – Carbon dioxide emissions (5-minute intervals)
• Demand – Electricity demand forecasts (30-minute intervals)
• SMP Pricing – System Marginal Prices (constrained & unconstrained, 30-minute intervals)

Please let me know if you'd like to adjust the format or add other data points in future reports."""
    
    # Append signature
    email_body += "\n\nThanks,\nMaitas Kopinsky"
    
    msg.set_content(email_body)
    
    # Attach the Excel file
    try:
        with open(summary_filename, 'rb') as f:
            file_data = f.read()
            msg.add_attachment(file_data, 
                             maintype='application', 
                             subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             filename=summary_filename)
    except Exception as e:
        print(f"✗ Error attaching file: {str(e)}")
        return False
    
    # Send email via Gmail SMTP
    try:
        print(f"Sending daily summary ({file_size:.1f} MB) via Gmail...")
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SMTP_USER, SMTP_PASS)
            smtp.send_message(msg)
        
        print(f"✓ Daily summary sent successfully to {', '.join(RECIPIENT_EMAILS)}!")
        return True
        
    except Exception as e:
        print(f"✗ Error sending email: {str(e)}")
        return False

def send_historical_file(filename, description, file_number, total_files):
    """Send a single historical Excel file via Gmail SMTP."""
    current_date = datetime.now()
    
    if not os.path.exists(filename):
        print(f"✗ File not found: {filename}")
        return False
    
    file_size = os.path.getsize(filename) / (1024 * 1024)  # MB
    
    # Create email message
    msg = EmailMessage()
    msg["Subject"] = f"Noga ISO Historical - {description} - {current_date.strftime('%Y-%m-%d')}"
    msg["From"] = SMTP_USER
    msg["To"] = ', '.join(RECIPIENT_EMAILS)
    
    # Email body
    email_body = f"""Noga ISO Historical Data - Part {file_number} of {total_files}

Generated on: {current_date.strftime('%Y-%m-%d %H:%M:%S')}

This email contains: {description} (Full Historical Data)
File size: {file_size:.1f} MB

Data from Israel's Independent System Operator (Noga ISO)
Source: https://apim-api.noga-iso.co.il/

This file contains complete historical data from the earliest available date.
This is part {file_number} of {total_files} historical data emails.
"""
    
    # Append signature
    email_body += "\n\nThanks,\nMaitas Kopinsky"
    
    msg.set_content(email_body)
    
    # Attach the Excel file
    try:
        with open(filename, 'rb') as f:
            file_data = f.read()
            msg.add_attachment(file_data, 
                             maintype='application', 
                             subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             filename=filename)
    except Exception as e:
        print(f"✗ Error attaching {filename}: {str(e)}")
        return False
    
    # Send email via Gmail SMTP
    try:
        print(f"Sending {filename} ({file_size:.1f} MB) via Gmail...")
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SMTP_USER, SMTP_PASS)
            smtp.send_message(msg)
        
        print(f"✓ {filename} sent successfully!")
        return True
        
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
    print("NOGA ISO API MAILER V2 (Gmail SMTP)")
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
    print("\nRequired environment variables:")
    print("  SENDER_EMAIL=your_gmail@gmail.com")
    print("  GMAIL_APP_PASSWORD=your_16_char_app_password")
    print("  RECIPIENT_EMAILS=recipient1@email.com,recipient2@email.com")
    print("\nUsage examples:")
    print("  py daily_api_mailer_v2.py                    # Daily summary only")
    print("  py daily_api_mailer_v2.py --historical       # Daily + historical")
    print("  py daily_api_mailer_v2.py --skip-fetch       # Use existing files")
    print("=" * 60)

if __name__ == "__main__":
    print_email_config()
    main() 