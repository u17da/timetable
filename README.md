# Timetable Parser

A web application that allows users to upload photos of printed schedules or Excel files and converts them into interactive, responsive weekly timetables using OpenAI's API.

## Features

- **File Upload Support**: Upload both image files (JPG, PNG) and Excel files (.xlsx, .xls)
- **AI-Powered Parsing**: Uses OpenAI's GPT-4 Vision API to extract timetable data from images and text processing for Excel files
- **Interactive Timetable Display**: Responsive weekly view showing all days with time slots, subjects, and room information
- **Modern UI**: Built with Next.js and Tailwind CSS for a clean, professional interface

## Technology Stack

### Backend
- **FastAPI**: Python web framework for the API
- **OpenAI API**: GPT-4 Vision for image processing and GPT-4 for Excel data structuring
- **Python Libraries**: 
  - `pillow` for image processing
  - `openpyxl` for Excel file handling
  - `python-multipart` for file uploads

### Frontend
- **Next.js 15**: React framework with TypeScript
- **Tailwind CSS**: Utility-first CSS framework
- **Lucide React**: Icon library

## Setup Instructions

### Prerequisites
- Python 3.12+
- Node.js 18+
- OpenAI API key

### Backend Setup

1. Navigate to the backend directory:
   ```bash
   cd backend
   ```

2. Install dependencies using Poetry:
   ```bash
   poetry install
   ```

3. Create a `.env` file and add your OpenAI API key:
   ```
   OPENAI_API_KEY=your_actual_openai_api_key_here
   ```

4. Start the development server:
   ```bash
   poetry run fastapi dev app/main.py
   ```

The backend will be available at `http://localhost:8000`

### Frontend Setup

1. Navigate to the frontend directory:
   ```bash
   cd frontend
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env.local` file:
   ```
   NEXT_PUBLIC_API_URL=http://localhost:8000
   ```

4. Start the development server:
   ```bash
   npm run dev
   ```

The frontend will be available at `http://localhost:3000`

## API Endpoints

### POST /upload
Upload and parse a timetable file (image or Excel)

**Request**: Multipart form data with file
**Response**: JSON with parsed timetable data

### GET /timetable/{file_id}
Retrieve a stored timetable by ID

### GET /timetables
List all stored timetables

### GET /healthz
Health check endpoint

## Data Structure

The application converts uploaded files into a standardized JSON format:

```json
{
  "title": "Schedule Title",
  "schedule": {
    "Monday": [
      {
        "time": "09:00-10:00",
        "subject": "Mathematics",
        "room": "A101"
      }
    ],
    "Tuesday": [...],
    "Wednesday": [...],
    "Thursday": [...],
    "Friday": [...],
    "Saturday": [...],
    "Sunday": [...]
  }
}
```

## Usage

1. Open the application in your browser at `http://localhost:3000`
2. Click "Choose File" to upload either:
   - A photo of a printed timetable (JPG, PNG)
   - An Excel file containing schedule data (.xlsx, .xls)
3. The application will process the file using OpenAI's API
4. View your interactive timetable with a responsive weekly layout
5. Click "Upload a different file" to process another schedule

## Important Notes

- **OpenAI API Key Required**: The application requires a valid OpenAI API key to function. Without it, file uploads will fail.
- **In-Memory Storage**: Timetable data is stored in memory and will be lost when the backend server restarts. This is suitable for development and proof-of-concept usage.
- **File Size Limits**: Large files may take longer to process due to OpenAI API processing time.

## Development

The application uses modern development practices:
- TypeScript for type safety
- Responsive design with Tailwind CSS
- Error handling and loading states
- Clean, modular code structure

## Deployment

For production deployment:
1. Set up environment variables for both frontend and backend
2. Build the frontend: `npm run build`
3. Deploy the backend with proper environment configuration
4. Ensure CORS settings are appropriate for your domain

## License

This project is licensed under the MIT License - see the LICENSE file for details.
