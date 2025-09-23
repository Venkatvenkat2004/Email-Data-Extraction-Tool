import imaplib
import getpass

def test_gmail_connection():
    print("🔐 Gmail Connection Test")
    print("=" * 30)
    
    # Input credentials safely
    email = input("Enter your Gmail address: ").strip()
    password = getpass.getpass("Enter app password (spaces illama): ").strip()
    
    print(f"Testing: {email}")
    print(f"Password length: {len(password)} characters")
    
    try:
        # Connect to Gmail
        print("🔄 Connecting to imap.gmail.com...")
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        
        print("🔑 Attempting login...")
        mail.login(email, password)
        print("✅ LOGIN SUCCESSFUL!")
        
        # Check inbox
        mail.select("inbox")
        print("✅ Inbox access successful!")
        
        mail.logout()
        return True
        
    except Exception as e:
        print(f"❌ LOGIN FAILED: {e}")
        print("\n🔧 Possible solutions:")
        print("1. App password spaces remove pannu")
        print("2. New app password generate pannu")
        print("3. Gmail address correct ah check pannu")
        return False

if __name__ == "__main__":
    test_gmail_connection()