# ============================================================
# vmiq-rvtools-etl — Dockerfile
# Compatible with OpenShift restricted-v2 SCC (non-root).
# Final image is ~120MB (python:3.12-slim base).
# ============================================================

FROM python:3.12-slim

# Install only what we need — no dev tools in the final image
RUN apt-get update && apt-get install -y --no-install-recommends \
        libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Create a non-root user — OCP runs as arbitrary UID in range
# 1000–65535. We set 1001 to match our SCC runAsUser.
RUN groupadd -g 1001 vmiq && \
    useradd -u 1001 -g vmiq -s /bin/sh -m vmiq

WORKDIR /app

# Install Python dependencies first (layer caching — only
# reinstalls when requirements.txt changes)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY rvtools_etl/ ./rvtools_etl/

# OCP may run as an arbitrary UID that is not in /etc/passwd.
# Making the home dir group-writable satisfies that requirement.
RUN chown -R 1001:1001 /app && chmod -R g=u /app

USER 1001

# The CronJob calls:  python -m rvtools_etl.job
CMD ["python", "-m", "rvtools_etl.job"]
