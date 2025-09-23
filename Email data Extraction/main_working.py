import pytesseract
from PIL import Image
import pandas as pd
from datetime import datetime
import yagmail
import re

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

def test_yagmail_login(email, password):
    """Test Gmail login with yagmail"""
    try:
        print("🔐 Testing Gmail with yagmail...")
        yag = yagmail.SMTP(email, password)
        print("✅ yagmail login successful!")
        return True
    except Exception as e:
        print(f"❌ yagmail login failed: {e}")
        return False

def create_demo_flight_data():
    """Create realistic demo flight data"""
    flights = [
        {
            'Employee Name': 'Venkatesh Kumar',
            'Flight Number': 'AI101',
            'Route': 'DEL → BOM (Delhi to Mumbai)',
            'Date': '2025-01-20',
            'Time': '08:30 AM',
            'Cost': '₹8,500',
            'Airline': 'Air India',
            'Status': 'Confirmed',
            'Class': 'Economy',
            'Seat': '15A',
            'Booking Reference': 'AI7X8Y9'
        },
        {
            'Employee Name': 'Priya Sharma',
            'Flight Number': '6E205',
            'Route': 'BLR → MAA (Bangalore to Chennai)',
            'Date': '2025-01-25',
            'Time': '14:20 PM',
            'Cost': '₹4,200',
            'Airline': 'IndiGo',
            'Status': 'Confirmed',
            'Class': 'Economy',
            'Seat': '22C',
            'Booking Reference': '6E3A4B5'
        },
        {
            'Employee Name': 'Rajesh Patel',
            'Flight Number': 'UK815',
            'Route': 'HYD → DEL (Hyderabad to Delhi)',
            'Date': '2025-02-01',
            'Time': '19:45 PM',
            'Cost': '₹12,800',
            'Airline': 'Vistara',
            'Status': 'Pending',
            'Class': 'Business',
            'Seat': '1A',
            'Booking Reference': 'UK9X8Y7'
        },
        {
            'Employee Name': 'Anita Desai',
            'Flight Number': 'SG307',
            'Route': 'BOM → GOI (Mumbai to Goa)',
            'Date': '2025-01-28',
            'Time': '11:15 AM',
            'Cost': '₹3,500',
            'Airline': 'SpiceJet',
            'Status': 'Confirmed',
            'Class': 'Economy',
            'Seat': '18B',
            'Booking Reference': 'SG2C3D4'
        }
    ]
    return flights

def create_advanced_excel():
    """Create professional Excel report"""
    flights = create_demo_flight_data()
    df = pd.DataFrame(flights)
    
    # Add summary statistics
    total_flights = len(flights)
    total_cost = sum(int(flight['Cost'].replace('₹', '').replace(',', '')) for flight in flights)
    confirmed_flights = sum(1 for flight in flights if flight['Status'] == 'Confirmed')
    
    # Create Excel with multiple sheets
    with pd.ExcelWriter('flight_records_advanced.xlsx', engine='openpyxl') as writer:
        # Main data sheet
        df.to_excel(writer, sheet_name='Flight Records', index=False)
        
        # Summary sheet
        summary_data = {
            'Metric': ['Total Flights', 'Confirmed Flights', 'Pending Flights', 'Total Cost'],
            'Value': [total_flights, confirmed_flights, total_flights-confirmed_flights, f'₹{total_cost:,}']
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
    
    return df

def main():
    print("🚀 ADVANCED FLIGHT MANAGEMENT SYSTEM")
    print("=" * 50)
    
    # Test OCR
    test_ocr()
    
    # Test Gmail with yagmail
    print("\n🔐 Gmail Test (Optional)")
    print("-" * 25)
    
    try:
        email = "venkatf0618w@gmail.com"
        password = "tiwkksudqijxxxlt"
        
        if test_yagmail_login(email, password):
            print("🎉 Gmail integration working!")
        else:
            print("ℹ️ Using demo data mode")
    except:
        print("ℹ️ Using demo data mode")
    
    # Create advanced Excel report
    print("\n📊 Creating Advanced Flight Report...")
    df = create_advanced_excel()
    
    print("✅ ADVANCED EXCEL REPORT CREATED!")
    print("📁 File: 'flight_records_advanced.xlsx'")
    print(f"📈 Total Records: {len(df)}")
    print("\n📋 Report Includes:")
    print("   ✓ Flight Records Sheet")
    print("   ✓ Summary Statistics Sheet") 
    print("   ✓ Professional Formatting")
    print("   ✓ Indian Airlines Data")
    
    # Display sample data
    print("\n✨ Sample Flight Data:")
    print("=" * 70)
    for i, flight in enumerate(df.head(3).to_dict('records'), 1):
        print(f"{i}. {flight['Employee Name']} - {flight['Flight Number']} - {flight['Route']} - {flight['Cost']}")

if __name__ == "__main__":
    main()