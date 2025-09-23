import pytesseract
from PIL import Image
import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
from datetime import datetime
import time

# Tesseract path set
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def test_ocr():
    """OCR test pannu"""
    print("🧪 Testing OCR...")
    
    from PIL import Image, ImageDraw
    img = Image.new('RGB', (500, 200), color='white')
    d = ImageDraw.Draw(img)
    d.text((20, 20), "BOARDING PASS", fill='black')
    d.text((20, 50), "Airline: Sky Airlines", fill='black')
    d.text((20, 80), "Flight: AA1234", fill='black')
    d.text((20, 110), "From: NYC to LAX", fill='black')
    d.text((20, 140), "Date: 2025-01-15 | Cost: $450.00", fill='black')
    img.save('test_boarding_pass.jpg')
    
    text = pytesseract.image_to_string(Image.open('test_boarding_pass.jpg'))
    print("📄 Extracted Text:")
    print("=" * 40)
    print(text)
    print("=" * 40)
    return text

def test_gmail_login(email, password):
    """Simple Gmail login test"""
    try:
        print("🔐 Testing Gmail login...")
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(email, password)
        print("✅ Login successful!")
        
        # Check inbox
        mail.select("inbox")
        status, messages = mail.search(None, "ALL")
        email_ids = messages[0].split()
        print(f"📧 Inbox has {len(email_ids)} emails")
        
        mail.close()
        mail.logout()
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {e}")
        return False

def extract_flight_info_from_emails(email, password):
    """Extract flight info from Gmail"""
    try:
        print("📧 Connecting to Gmail...")
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(email, password)
        mail.select("inbox")
        
        # Search for flight-related emails
        search_keywords = ['Flight', 'Booking', 'Itinerary', 'Airline', 'Ticket']
        all_flights = []
        
        for keyword in search_keywords:
            print(f"🔍 Searching for '{keyword}' emails...")
            status, messages = mail.search(None, f'SUBJECT "{keyword}"')
            email_ids = messages[0].split()
            
            for email_id in email_ids[:2]:  # Process first 2 of each type
                try:
                    status, msg_data = mail.fetch(email_id, "(RFC822)")
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # Get subject
                    subject = decode_header(msg["Subject"])[0][0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(errors='ignore')
                    
                    print(f"   ✉️ {subject[:60]}...")
                    
                    # Extract basic info
                    flight_data = {
                        'employee_name': extract_name_from_subject(subject),
                        'flight_number': extract_flight_number(subject),
                        'route': 'Extracted from email',
                        'date': datetime.now().strftime("%Y-%m-%d"),
                        'cost': 'Extracted from email',
                        'email_subject': subject[:80],
                        'keyword': keyword
                    }
                    all_flights.append(flight_data)
                    
                except Exception as e:
                    print(f"   ❌ Error processing email: {e}")
        
        mail.close()
        mail.logout()
        return all_flights
        
    except Exception as e:
        print(f"❌ Email processing error: {e}")
        return []

def extract_name_from_subject(subject):
    """Extract name from email subject"""
    patterns = [
        r'for\s+([A-Za-z]+\s+[A-Za-z]+)',
        r'-\s*([A-Za-z]+\s+[A-Za-z]+)',
        r'([A-Za-z]+\s+[A-Za-z]+)\s+-\s+Flight'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, subject, re.IGNORECASE)
        if match:
            return match.group(1).title()
    
    return "Unknown"

def extract_flight_number(subject):
    """Extract flight number from subject"""
    match = re.search(r'[A-Z]{2}\d{3,4}', subject.upper())
    return match.group() if match else "Not found"

def create_sample_flight_data():
    """Create realistic sample flight data"""
    return [
        {
            'employee_name': 'Venkatesh Kumar',
            'flight_number': 'AI101',
            'route': 'DEL (Delhi) to BOM (Mumbai)',
            'date': '2025-01-20',
            'cost': '₹8,500',
            'airline': 'Air India',
            'status': 'Confirmed',
            'class': 'Economy'
        },
        {
            'employee_name': 'Priya Sharma',
            'flight_number': '6E205',
            'route': 'BLR (Bangalore) to MAA (Chennai)',
            'date': '2025-01-25',
            'cost': '₹4,200',
            'airline': 'IndiGo',
            'status': 'Confirmed',
            'class': 'Economy'
        },
        {
            'employee_name': 'Rajesh Patel',
            'flight_number': 'UK815',
            'route': 'HYD (Hyderabad) to DEL (Delhi)',
            'date': '2025-02-01',
            'cost': '₹6,800',
            'airline': 'Vistara',
            'status': 'Pending',
            'class': 'Business'
        }
    ]

def main():
    print("🚀 FLIGHT EMAIL AUTOMATION - FINAL VERSION")
    print("=" * 55)
    
    # Test OCR
    test_ocr()
    
    # Get credentials
    print("\n🔐 Gmail Login")
    print("-" * 20)
    email = input("venkatf0618w@gmail.com").strip()
    password = input("tiwkksudqijxxxlt").strip()
    
    # Test login first
    if test_gmail_login(email, password):
        # Try to extract real emails
        print("\n📧 Processing emails...")
        real_flights = extract_flight_info_from_emails(email, password)
        
        if real_flights:
            df = pd.DataFrame(real_flights)
            df.to_excel('flight_records.xlsx', index=False)
            print(f"🎉 Saved {len(real_flights)} real flight records!")
        else:
            # Create sample data if no real emails found
            sample_data = create_sample_flight_data()
            df = pd.DataFrame(sample_data)
            df.to_excel('flight_records.xlsx', index=False)
            print("📊 Created sample flight data (no real emails found)")
    else:
        # Create sample data if login fails
        print("\n📝 Creating sample data for demonstration...")
        sample_data = create_sample_flight_data()
        df = pd.DataFrame(sample_data)
        df.to_excel('flight_records.xlsx', index=False)
        print("💾 Sample data saved to 'flight_records.xlsx'")
    
    print("\n✅ Process completed!")
    print("📁 Check 'flight_records.xlsx' file")

if __name__ == "__main__":
    main()