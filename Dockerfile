# Use a slim Python base image
FROM python:3.11-slim-bookworm

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file and install dependencies
# This step is optimized for Docker caching:
# If requirements.txt doesn't change, Docker won't re-run pip install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code
# This includes app.py, config.py, excel_processor.py, ppt_updater.py, utils.py
# and your template PPT (FINAL_PowerPoint_Template.pptx)
COPY . .

# Expose the port that Gunicorn will listen on
EXPOSE 8000

# Command to run the application using Gunicorn
# Gunicorn is a WSGI HTTP Server for UNIX. CAE environments typically expect this.
# -w 4: Run with 4 worker processes (adjust based on your app's needs and CAE resources)
# app:app: Points to the 'app' Flask instance within the 'app.py' file
# -b 0.0.0.0:8000: Binds to all network interfaces on port 8000
CMD ["gunicorn", "-w", "4", "app:app", "-b", "0.0.0.0:8000"]