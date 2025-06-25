'use client';

import { useState } from 'react';
import { Upload, Calendar, FileText, Image as ImageIcon } from 'lucide-react';

interface TimetableEntry {
  time: string;
  subject: string;
  room: string;
}

interface TimetableData {
  title: string;
  schedule: {
    [key: string]: TimetableEntry[];
  };
}

export default function Home() {
  const [timetableData, setTimetableData] = useState<TimetableData | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setError(null);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const apiUrl = typeof window !== 'undefined' && window.location.hostname !== 'localhost'
        ? 'https://app-cnignukt.fly.dev' 
        : 'http://localhost:8000';

      const response = await fetch(`${apiUrl}/upload`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('Failed to upload file');
      }

      const result = await response.json();
      setTimetableData(result.data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setIsUploading(false);
    }
  };

  const weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  return (
    <div className="min-h-screen bg-gray-50 py-8 px-4">
      <div className="max-w-6xl mx-auto">
        <header className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-2 flex items-center justify-center gap-3">
            <Calendar className="text-blue-600" size={40} />
            Timetable Parser
          </h1>
          <p className="text-gray-600 text-lg">
            Upload a photo of your schedule or an Excel file to generate an interactive timetable
          </p>
        </header>

        {!timetableData && (
          <div className="bg-white rounded-lg shadow-lg p-8 mb-8">
            <div className="text-center">
              <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 hover:border-blue-400 transition-colors">
                <Upload className="mx-auto text-gray-400 mb-4" size={48} />
                <h3 className="text-xl font-semibold text-gray-700 mb-2">
                  Upload Your Timetable
                </h3>
                <p className="text-gray-500 mb-6">
                  Drag and drop or click to select a photo or Excel file
                </p>
                
                <div className="flex justify-center gap-4 mb-6">
                  <div className="flex items-center gap-2 text-sm text-gray-600">
                    <ImageIcon size={16} />
                    Photos (JPG, PNG)
                  </div>
                  <div className="flex items-center gap-2 text-sm text-gray-600">
                    <FileText size={16} />
                    Excel Files (.xlsx, .xls)
                  </div>
                </div>

                <input
                  type="file"
                  accept="image/*,.xlsx,.xls"
                  onChange={handleFileUpload}
                  disabled={isUploading}
                  className="hidden"
                  id="file-upload"
                />
                <label
                  htmlFor="file-upload"
                  className={`inline-flex items-center px-6 py-3 border border-transparent text-base font-medium rounded-md text-white bg-blue-600 hover:bg-blue-700 cursor-pointer transition-colors ${
                    isUploading ? 'opacity-50 cursor-not-allowed' : ''
                  }`}
                >
                  {isUploading ? 'Processing...' : 'Choose File'}
                </label>
              </div>
            </div>

            {error && (
              <div className="mt-4 p-4 bg-red-50 border border-red-200 rounded-md">
                <p className="text-red-600">{error}</p>
              </div>
            )}
          </div>
        )}

        {timetableData && (
          <div className="bg-white rounded-lg shadow-lg overflow-hidden">
            <div className="bg-blue-600 text-white p-6">
              <h2 className="text-2xl font-bold">{timetableData.title || 'Your Timetable'}</h2>
              <button
                onClick={() => setTimetableData(null)}
                className="mt-2 text-blue-100 hover:text-white underline"
              >
                Upload a different file
              </button>
            </div>

            <div className="p-6">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                {weekdays.map((day) => (
                  <div key={day} className="border border-gray-200 rounded-lg overflow-hidden">
                    <div className="bg-gray-100 px-4 py-3 border-b">
                      <h3 className="font-semibold text-gray-800">{day}</h3>
                    </div>
                    <div className="p-4">
                      {timetableData.schedule[day]?.length > 0 ? (
                        <div className="space-y-3">
                          {timetableData.schedule[day].map((entry, index) => (
                            <div
                              key={index}
                              className="bg-blue-50 border border-blue-200 rounded-md p-3"
                            >
                              <div className="text-sm font-medium text-blue-800">
                                {entry.time}
                              </div>
                              <div className="text-gray-800 font-semibold">
                                {entry.subject}
                              </div>
                              {entry.room && (
                                <div className="text-sm text-gray-600">
                                  Room: {entry.room}
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="text-gray-400 text-center py-8">
                          No classes scheduled
                        </div>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
