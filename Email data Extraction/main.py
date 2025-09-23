import pytesseract
from PIL import Image
import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
import os
from datetime import datetime

# Tesseract path set
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class EmailFlightExtractor:
    def __init__(self, email_user, email_pass, imap_server="imap.gmail.com"):
        self.email_user = email_user
        self.email_pass = email_pass
        self.imap_server = imap_server
        self.mail = None
        
    def connect_to_email(self):
        """Email server connect pannu"""
        try:
            print(f"🔗 Connecting to {self.imap_server}...")
            self.mail = imaplib.IMAP4_SSL(self.imap_server)
            
            print(f"🔑 Logging in as {self.email_user}...")
            self.mail.login(self.email_user, self.email_pass)
            
            print("📁 Selecting inbox...")
            self.mail.select("inbox")
            
            print("✅ Email connected successfully")
            return True
            
        except Exception as e:
            print(f"❌ Email connection failed: {e}")
            return False
    
    def search_flight_emails(self):
        """Flight confirmation emails search pannu"""
        try:
            # Search for flight emails
            status, messages = self.mail.search(None, '(SUBJECT "Flight" SUBJECT "Booking" SUBJECT "Itinerary")')
            email_ids = messages[0].split()
            print(f"📧 Found {len(email_ids)} potential flight emails")
            return email_ids
        except Exception as e:
            print(f"❌ Email search failed: {e}")
            return []
    
    def extract_email_content(self, email_id):
        """Email content extract pannu"""
        try:
            status, msg_data = self.mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # Subject extract pannu
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(errors='ignore')
            
            # Body extract pannu
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        body = part.get_payload(decode=True).decode(errors='ignore')
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors='ignore')
            
            return subject, body
            
        except Exception as e:
            print(f"❌ Email content extraction failed: {e}")
            return "", ""
    
    def extract_flight_details(self, subject, body):
        """Email subject and body la irunthu flight details extract pannu"""
        details = {}
        
        # Employee name extract from subject
        name_patterns = [
            r'Flight Confirmation - ([^-]+)',
            r'Booking Confirmation - ([^-]+)',
            r'Itinerary for ([^-]+)'
        ]
        
        details['employee_name'] = "Unknown"
        for pattern in name_patterns:
            match = re.search(pattern, subject, re.IGNORECASE)
            if match:
                details['employee_name'] = match.group(1).strip()
                break
        
        # Flight number extract
        flight_match = re.search(r'[A-Z]{2}\d{3,4}', body.upper())
        details['flight_number'] = flight_match.group() if flight_match else "Not found"
        
        # Route extract
        route_match = re.search(r'(\w{3})\s*to\s*(\w{3})', body, re.IGNORECASE)
        if route_match:
            details['route'] = f"{route_match.group(1)} to {route_match.group(2)}"
        else:
            details['route'] = "Not found"
        
        # Date extract
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', body)
        details['date'] = date_match.group() if date_match else "Not found"
        
        # Cost extract
        cost_match = re.search(r'[\$₹]\d+(?:\.\d{2})?', body)
        details['cost'] = cost_match.group() if cost_match else "Not found"
        
        return details
    
    def process_emails(self):
        """All flight emails process pannu"""
        if not self.connect_to_email():
            return []
        
        email_ids = self.search_flight_emails()
        all_flights = []
        
        for i, email_id in enumerate(email_ids[:3]):  # First 3 emails test pannu
            print(f"\n📨 Processing email {i+1}/{len(email_ids)}...")
            subject, body = self.extract_email_content(email_id)
            
            if subject:
                print(f"Subject: {subject[:80]}...")
                flight_details = self.extract_flight_details(subject, body)
                
                if flight_details:
                    flight_details['email_id'] = email_id.decode()
                    flight_details['processed_date'] = datetime.now().strftime("%Y-%m-%d")
                    all_flights.append(flight_details)
                    print(f"✅ Extracted: {flight_details}")
        
        # Close connection
        if self.mail:
            self.mail.close()
            self.mail.logout()
        
        return all_flights
    
    def save_to_excel(self, flight_data, filename="flight_records.xlsx"):
        """Excel file la save pannu"""
        if not flight_data:
            print("❌ No flight data to save")
            # Create sample data for testing
            sample_data = [{
                'employee_name': 'Venkatesh Kumar',
                'flight_number': 'AI101',
                'route': 'DEL to BOM',
                'date': '2025-01-20',
                'cost': '₹8,500',
                'status': 'Sample Data - No real emails found'
            }]
            flight_data = sample_data
            print("📝 Created sample data for testing")
        
        df = pd.DataFrame(flight_data)
        df.to_excel(filename, index=False)
        print(f"💾 Saved {len(flight_data)} records to {filename}")

# Test function
def test_ocr():
    """OCR test pannu"""
    print("🧪 Testing OCR...")
    
    # Test image create pannu
    from PIL import Image, ImageDraw
    
    img = Image.new('RGB', (500, 200), color='white')
    d = ImageDraw.Draw(img)
    
    # Sample boarding pass content
    d.text((20, 20), "BOARDING PASS", fill='black')
    d.text((20, 50), "Airline: Sky Airlines", fill='black')
    d.text((20, 80), "Flight: AA1234", fill='black')
    d.text((20, 110), "From: NYC to LAX", fill='black')
    d.text((20, 140), "Date: 2025-01-15 | Cost: $450.00", fill='black')
    
    img.save('test_boarding_pass.jpg')
    
    # OCR extract pannu
    text = pytesseract.image_to_string(Image.open('test_boarding_pass.jpg'))
    print("📄 Extracted Text:")
    print("=" * 40)
    print(text)
    print("=" * 40)
    
    return text

def main():
    print("🚀 Flight Email Automation Started!")
    print("=" * 50)
    
    # Test OCR first
    test_ocr()
    
    # Get credentials
    print("\n🔐 Enter Gmail credentials:")
    your_email = input("venkatf0618w@gmail.com").strip()  # Your email
    your_app_password = input("tiwkksudqijxxxlt").strip()  # Enter your 16-char password
    
    try:
        # Test connection first
        print("🔄 Testing connection...")
        mail_test = imaplib.IMAP4_SSL("imap.gmail.com")
        mail_test.login(your_email, your_app_password)
        mail_test.logout()
        print("✅ Connection test passed!")
        
        # Process emails
        extractor = EmailFlightExtractor(your_email, your_app_password)
        flights = extractor.process_emails()
        extractor.save_to_excel(flights)
        
        print(f"\n🎉 Success! Found {len(flights)} flight records!")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("📝 Creating sample data...")
        extractor = EmailFlightExtractor("test@test.com", "test")
        extractor.save_to_excel([])

if __name__ == "__main__":
    main()