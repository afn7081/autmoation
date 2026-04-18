FROM mcr.microsoft.com/azure-functions/python:4-python3.11

ENV AzureWebJobsScriptRoot=/home/site/wwwroot \
    AzureFunctionsJobHost__Logging__Console__IsEnabled=true

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-java-common \
    default-jre \
    fonts-dejavu \
    fonts-liberation && \
    apt-get clean && rm -rf /var/lib/apt/lists/* && \
    mkdir -p /tmp/.libreoffice && chmod 777 /tmp/.libreoffice

COPY requirements.txt /home/site/wwwroot/
RUN pip install --no-cache-dir -r /home/site/wwwroot/requirements.txt

COPY . /home/site/wwwroot
