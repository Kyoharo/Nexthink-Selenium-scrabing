
FROM python:3.9

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

# Optional: Copy additional files (like geckodriver or chromedriver)
COPY geckodriver /app/geckodriver 
COPY chromedriver /app/chromedriver 


ENTRYPOINT ["python", "main.py"]
