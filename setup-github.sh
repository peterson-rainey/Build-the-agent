#!/bin/bash

# Google Ads Error Monitor - GitHub Setup Script
# This script helps you set up the GitHub repository with auto-save functionality

echo "🚀 Setting up GitHub repository for Google Ads Error Monitor"
echo "=========================================================="

# Check if GitHub CLI is installed
if ! command -v gh &> /dev/null; then
    echo "❌ GitHub CLI (gh) is not installed."
    echo "Please install it first:"
    echo "  macOS: brew install gh"
    echo "  Or visit: https://cli.github.com/"
    exit 1
fi

# Check if user is authenticated
if ! gh auth status &> /dev/null; then
    echo "🔐 Please authenticate with GitHub first:"
    gh auth login
fi

# Get repository name
read -p "Enter repository name (default: google-ads-error-monitor): " repo_name
repo_name=${repo_name:-google-ads-error-monitor}

# Get repository description
read -p "Enter repository description (default: Google Ads Error Monitoring Script with Auto-Save): " repo_description
repo_description=${repo_description:-Google Ads Error Monitoring Script with Auto-Save}

# Ask for visibility
echo "Choose repository visibility:"
echo "1) Public (recommended for open source)"
echo "2) Private"
read -p "Enter choice (1 or 2): " visibility_choice

if [ "$visibility_choice" = "2" ]; then
    visibility="--private"
else
    visibility="--public"
fi

echo "📦 Creating GitHub repository..."
gh repo create "$repo_name" --description "$repo_description" $visibility --source=. --remote=origin --push

if [ $? -eq 0 ]; then
    echo "✅ Repository created successfully!"
    echo "🔗 Repository URL: https://github.com/$(gh api user --jq .login)/$repo_name"
    
    echo ""
    echo "🎉 Setup complete! Your repository now includes:"
    echo "   • Auto-save workflow (daily at 2 AM UTC)"
    echo "   • Comprehensive README documentation"
    echo "   • MIT License"
    echo "   • .gitignore for clean version control"
    echo "   • All your Google Ads scripts"
    
    echo ""
    echo "📋 Next steps:"
    echo "   1. Visit your repository on GitHub"
    echo "   2. Check the Actions tab to see the auto-save workflow"
    echo "   3. Customize the script configuration as needed"
    echo "   4. Set up your Google Ads account mappings"
    
    echo ""
    echo "🔄 The auto-save workflow will:"
    echo "   • Run daily at 2 AM UTC"
    echo "   • Commit any changes automatically"
    echo "   • Create backup tags for version history"
    echo "   • Clean up old backup tags (keeps last 30)"
    
else
    echo "❌ Failed to create repository. Please check your GitHub authentication and try again."
    exit 1
fi
