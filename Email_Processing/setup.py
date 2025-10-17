#!/usr/bin/env python3
"""
Email Processing System Setup Script
Installs required dependencies and sets up the system.
"""

import subprocess
import sys
import os

def install_dependencies():
    """Install required Python packages."""
    print("🔧 Installing required dependencies...")
    
    try:
        # Install mbox-to-json
        subprocess.run([sys.executable, "-m", "pip", "install", "mbox-to-json"], 
                      check=True, capture_output=True, text=True)
        print("✅ mbox-to-json installed successfully")
        
        # Install any other dependencies if needed
        print("✅ All dependencies installed")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error installing dependencies: {e}")
        return False

def check_system():
    """Check if the system is ready."""
    print("\n🔍 Checking system requirements...")
    
    # Check Python version
    if sys.version_info < (3, 7):
        print("❌ Python 3.7+ required")
        return False
    else:
        print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    
    # Check if mbox-to-json is available
    try:
        subprocess.run(["mbox-to-json", "-h"], 
                      check=True, capture_output=True, text=True)
        print("✅ mbox-to-json is available")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("⚠️  mbox-to-json not found - will install it")
        return install_dependencies()
    
    return True

def show_usage():
    """Show usage instructions."""
    print("\n" + "="*60)
    print("📧 EMAIL PROCESSING SYSTEM - READY TO USE!")
    print("="*60)
    print("\n🚀 Quick Start:")
    print("1. Place your .mbox file in the 'raw_data/' folder")
    print("2. Run: python3 extract_user_emails.py raw_data/your_file.mbox user@example.com")
    print("\n📋 Available Commands:")
    print("• Extract from MBOX: python3 extract_user_emails.py input.mbox user@email.com")
    print("• Extract from JSON: python3 extract_from_json.py input.json user@email.com")
    print("\n📁 Folder Structure:")
    print("• raw_data/     - Place your .mbox files here")
    print("• processed_data/ - Extracted email datasets appear here")
    print("\n📖 For detailed instructions, see README.md")

def main():
    """Main setup function."""
    print("🔧 Email Processing System Setup")
    print("="*40)
    
    if check_system():
        show_usage()
        print("\n✅ System is ready to use!")
    else:
        print("\n❌ Setup failed. Please check the errors above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
