# Use Windows Server Core with Python
FROM mcr.microsoft.com/windows/servercore:ltsc2019

# Install Python
RUN powershell -Command \
    Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.11.0/python-3.11.0-amd64.exe -OutFile python-installer.exe; \
    Start-Process python-installer.exe -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1' -Wait; \
    Remove-Item python-installer.exe

# Set working directory
WORKDIR /app

# Copy requirements and install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY gold_app.py .

# Build the executable
RUN pyinstaller --onefile --windowed --name "GoldTransactionManager" gold_app.py

# The output will be in dist/ directory