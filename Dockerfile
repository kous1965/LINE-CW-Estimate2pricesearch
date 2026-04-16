# ベースイメージを「Bullseye（Debian 11）」に固定して、パッケージ消失エラーを防ぐ
FROM python:3.9-slim-bullseye

# 設定
ENV PYTHONUNBUFFERED=1
ENV PIP_BREAK_SYSTEM_PACKAGES=1
WORKDIR /app

# 1. 基本ツールと依存パッケージのインストール
# Bullseyeを使うことで libgconf-2-4 や fonts-kacst がインストール可能になります
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    curl \
    gnupg \
    unzip \
    ca-certificates \
    fonts-ipafont-gothic \
    fonts-wqy-zenhei \
    fonts-thai-tlwg \
    fonts-kacst \
    libglib2.0-0 \
    libnss3 \
    libgconf-2-4 \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# 2. Google Chrome (Stable) のインストール
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /usr/share/keyrings/google-chrome-keyring.gpg \
    && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome-keyring.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

# 3. requirements.txt をコピーしてインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. ソースコードをコピー
COPY . .

# ポート8080を開放
EXPOSE 8080

# アプリ起動
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]