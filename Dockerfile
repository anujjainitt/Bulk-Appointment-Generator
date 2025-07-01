# Use official Node.js image
FROM node:18

# Set working directory
WORKDIR /app

# Copy backend package files and install dependencies
COPY backend/package*.json ./backend/
RUN cd backend && npm install

# Copy frontend and build it
COPY frontend ./frontend
RUN cd frontend && npm install && npm run build

# Copy backend source
COPY backend ./backend

# Expose port
EXPOSE 5000

# Start the backend (which serves the frontend build)
CMD ["node", "backend/server.js"] 