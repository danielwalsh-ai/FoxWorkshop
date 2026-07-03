# Playwright's official Python image — browsers + system deps preinstalled,
# so the Autovolt scraper works in the container out of the box.
FROM mcr.microsoft.com/playwright/python:v1.61.0-jammy

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# The web dashboard (the tile). The daily 6pm job runs as a Coolify
# scheduled task: `python run_daily.py` inside this same image.
EXPOSE 8080
CMD ["uvicorn", "webapp:app", "--host", "0.0.0.0", "--port", "8080"]
