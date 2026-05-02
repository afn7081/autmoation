FROM mcr.microsoft.com/azure-functions/python:4-python3.11

ENV AzureWebJobsScriptRoot=/home/site/wwwroot \
    AzureFunctionsJobHost__Logging__Console__IsEnabled=true

# MS EULA auto-accept for ttf-mscorefonts-installer (Calibri/Arial/Times).
RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" \
    | debconf-set-selections

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-java-common \
    default-jre \
    fonts-dejavu \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    curl \
    ca-certificates && \
    # Enable contrib for ttf-mscorefonts-installer
    sed -i 's/ main$/ main contrib/' /etc/apt/sources.list || true && \
    apt-get update && \
    apt-get install -y --no-install-recommends ttf-mscorefonts-installer || true && \
    # Install Google Fonts used by the certificate (direct TTF downloads from
    # google/fonts repo). Montserrat is a variable font; brackets must be URL-encoded.
    mkdir -p /usr/share/fonts/truetype/google && \
    cd /usr/share/fonts/truetype/google && \
    RAW="https://raw.githubusercontent.com/google/fonts/main/ofl" && \
    curl -fsSL -o "Montserrat[wght].ttf"        "$RAW/montserrat/Montserrat%5Bwght%5D.ttf"        && \
    curl -fsSL -o "Montserrat-Italic[wght].ttf" "$RAW/montserrat/Montserrat-Italic%5Bwght%5D.ttf" && \
    curl -fsSL -o "GreatVibes-Regular.ttf"      "$RAW/greatvibes/GreatVibes-Regular.ttf"          && \
    curl -fsSL -o "Allura-Regular.ttf"          "$RAW/allura/Allura-Regular.ttf"                  && \
    # "Signature" font isn't on Google Fonts; alias it to Allura at LibreOffice level
    # via a fontconfig substitution so the PPTX "Signature" runs still render as script.
    mkdir -p /etc/fonts/conf.d && \
    printf '%s\n' \
      '<?xml version="1.0"?>' \
      '<!DOCTYPE fontconfig SYSTEM "fonts.dtd">' \
      '<fontconfig>' \
      '  <alias><family>Signature</family><prefer><family>Allura</family></prefer></alias>' \
      '  <alias><family>Calibri</family><prefer><family>Carlito</family></prefer></alias>' \
      '</fontconfig>' > /etc/fonts/conf.d/99-cert-fonts.conf && \
    fc-cache -f && \
    apt-get clean && rm -rf /var/lib/apt/lists/* /tmp/* && \
    mkdir -p /tmp/.libreoffice && chmod 777 /tmp/.libreoffice

COPY requirements.txt /home/site/wwwroot/
RUN pip install --no-cache-dir -r /home/site/wwwroot/requirements.txt

COPY . /home/site/wwwroot
