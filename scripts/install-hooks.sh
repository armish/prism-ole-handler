#!/bin/bash

# Script to install Git hooks for development
# Run this script from the project root directory

set -e

echo "🔧 Installing Git hooks..."

# Ensure we're in the project root
if [ ! -f "pyproject.toml" ]; then
    echo "❌ Error: Please run this script from the project root directory"
    exit 1
fi

# Create hooks directory if it doesn't exist
mkdir -p .git/hooks

# Install pre-push hook
cat > .git/hooks/pre-push << 'EOF'
#!/bin/bash

# Pre-push hook to run code quality checks
# This prevents pushing code that would fail CI

set -e

echo "🔍 Running pre-push checks..."

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}✓${NC} $1"
}

print_error() {
    echo -e "${RED}✗${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

# Check if we're in the right directory
if [ ! -f "pyproject.toml" ]; then
    print_error "Not in project root directory"
    exit 1
fi

# 1. Check Black formatting
echo ""
echo "📝 Checking code formatting with Black..."
if black --check prism_ole_handler/ tests/; then
    print_status "Code formatting is correct"
else
    print_error "Code formatting issues found"
    echo ""
    print_warning "Run 'black prism_ole_handler/ tests/' to fix formatting"
    exit 1
fi

# 2. Run unit tests
echo ""
echo "🧪 Running unit tests..."
if python -m pytest tests/ -v --tb=short; then
    print_status "All tests passed"
else
    print_error "Tests failed"
    exit 1
fi

# 3. Check that we can import the package
echo ""
echo "📦 Checking package imports..."
if python -c "import prism_ole_handler; print('Package imports successfully')"; then
    print_status "Package imports correctly"
else
    print_error "Package import failed"
    exit 1
fi

echo ""
print_status "All pre-push checks passed! 🚀"
echo "Pushing to remote..."
EOF

# Make the hook executable
chmod +x .git/hooks/pre-push

echo "✅ Pre-push hook installed successfully!"
echo ""
echo "ℹ️  The hook will now run automatically before each 'git push'"
echo "   It will check:"
echo "   • Code formatting (Black)"
echo "   • Unit tests"
echo "   • Package imports"
echo ""
echo "   To bypass the hook (not recommended): git push --no-verify"