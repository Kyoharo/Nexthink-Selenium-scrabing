FROM public.ecr.aws/docker/library/python:3.12.1-slim

WORKDIR /app

# Copy all files to the container's /app directory
COPY . .

# Install dependencies
RUN pip install -r requirements.txt


ENTRYPOINT ["python", "main.py"]
