#!/bin/bash

echo ""
echo "Restoring frontend npm packages"
echo ""
cd frontend
npm install
if [ $? -ne 0 ]; then

echo ""
npm run build
python -m pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "Failed to restore backend python packages"
    exit $?
fi

cd frontend
# Clean up node_modules and lock file to ensure a fresh install
if [ -d node_modules ]; then
    rm -rf node_modules
fi
if [ -f package-lock.json ]; then
    rm package-lock.json
fi

# Install npm-force-resolutions if not present
npm install npm-force-resolutions --save-dev
if [ $? -ne 0 ]; then
    echo "Failed to install npm-force-resolutions"
    exit $?
fi

# Run npm-force-resolutions to enforce resolutions in package.json
npx npm-force-resolutions
if [ $? -ne 0 ]; then
    echo "Failed to run npm-force-resolutions"
    exit $?
fi

# Install all frontend dependencies
npm install
if [ $? -ne 0 ]; then
    echo "Failed to restore frontend npm packages"
    exit $?
fi
if [ $? -ne 0 ]; then
    echo "Failed to build frontend"
    exit $?
fi

cd ..
. ./scripts/loadenv.sh

echo ""
echo "Starting backend"
echo ""
./.venv/bin/python -m quart run --port=50505 --host=127.0.0.1 --reload
if [ $? -ne 0 ]; then
    echo "Failed to start backend"
    exit $?
fi
# Open browser to backend URL
if command -v xdg-open > /dev/null; then
    xdg-open http://127.0.0.1:50505
elif command -v open > /dev/null; then
    open http://127.0.0.1:50505
fi

# Start backend (using uvicorn to match .cmd)
python -m uvicorn app:app --port 50505 --reload
if [ $? -ne 0 ]; then
    echo "Failed to start backend"
    exit $?
fi
