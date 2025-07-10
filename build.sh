#!/bin/bash

# Build script for Gold Transaction Manager

echo "Building Gold Transaction Manager..."

# Create build directory
mkdir -p build
cd build

# Copy files
cp ../gold_app.py .
cp ../requirements.txt .
cp ../Dockerfile .

# Build Docker image
echo "Building Docker image..."
docker build -t gold-app-builder .

# Run container and extract the EXE
echo "Extracting EXE file..."
docker run --name gold-app-container gold-app-builder
docker cp gold-app-container:/app/dist/GoldTransactionManager.exe ./GoldTransactionManager.exe
docker rm gold-app-container

echo "Build completed! EXE file is ready: GoldTransactionManager.exe"
echo "You can now copy this file to any Windows machine and run it."

# Clean up
cd ..
echo "Cleaning up..."
docker rmi gold-app-builder