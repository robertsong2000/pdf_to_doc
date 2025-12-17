#!/bin/bash
set -e

# Create necessary directories
mkdir -p "${UPLOAD_FOLDER:-/app/uploads}"
mkdir -p "${OUTPUT_FOLDER:-/app/outputs}"

# Set proper permissions
chmod 755 "${UPLOAD_FOLDER:-/app/uploads}"
chmod 755 "${OUTPUT_FOLDER:-/app/outputs}"

# Generate a random SECRET_KEY if not provided
if [ "$FLASK_ENV" = "production" ] && [ -z "$SECRET_KEY" ]; then
    export SECRET_KEY=$(python -c 'import secrets; print(secrets.token_hex(32))')
    echo "Generated SECRET_KEY for production"
fi

# Check if we need to wait for dependencies (like a database)
# In this case, we don't have external dependencies

# Run the application
exec "$@"