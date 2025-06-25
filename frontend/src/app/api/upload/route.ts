import { NextRequest, NextResponse } from 'next/server';
import OpenAI from 'openai';
import * as XLSX from 'xlsx';
import sharp from 'sharp';

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const timetablesStorage: Record<string, unknown> = {};

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    
    if (!file) {
      return NextResponse.json(
        { error: 'No file provided' },
        { status: 400 }
      );
    }

    if (file.type.startsWith('image/')) {
      return await processImageFile(file);
    } else if (
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel'
    ) {
      return await processExcelFile(file);
    } else {
      return NextResponse.json(
        { error: 'Unsupported file type. Please upload an image or Excel file.' },
        { status: 400 }
      );
    }
  } catch (error) {
    console.error('Upload error:', error);
    return NextResponse.json(
      { error: 'Error processing file' },
      { status: 500 }
    );
  }
}

async function processImageFile(file: File) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    
    const pngBuffer = await sharp(buffer).png().toBuffer();
    const base64Image = pngBuffer.toString('base64');
    
    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "user",
          content: [
            {
              type: "text",
              text: `Please analyze this timetable image and extract the schedule information. Return a JSON object with the following structure:
              {
                "title": "Schedule title if visible",
                "schedule": {
                  "Monday": [{"time": "09:00-10:00", "subject": "Math", "room": "A101"}],
                  "Tuesday": [{"time": "09:00-10:00", "subject": "English", "room": "B202"}],
                  "Wednesday": [],
                  "Thursday": [],
                  "Friday": [],
                  "Saturday": [],
                  "Sunday": []
                }
              }
              Extract all visible time slots, subjects, and room numbers. If information is unclear, use your best judgment.`
            },
            {
              type: "image_url",
              image_url: {
                url: `data:image/png;base64,${base64Image}`
              }
            }
          ]
        }
      ],
      max_tokens: 1000
    });

    const content = response.choices[0].message.content;
    let timetableData;
    
    try {
      timetableData = JSON.parse(content || '{}');
    } catch {
      const jsonMatch = content?.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
      if (jsonMatch) {
        timetableData = JSON.parse(jsonMatch[1]);
      } else {
        const simpleJsonMatch = content?.match(/\{[\s\S]*\}/);
        if (simpleJsonMatch) {
          timetableData = JSON.parse(simpleJsonMatch[0]);
        } else {
          throw new Error('Failed to parse OpenAI response as JSON');
        }
      }
    }

    const fileId = `img_${Object.keys(timetablesStorage).length}`;
    timetablesStorage[fileId] = timetableData;

    return NextResponse.json({
      id: fileId,
      data: timetableData
    });

  } catch (error) {
    console.error('Image processing error:', error);
    return NextResponse.json(
      { error: 'Error processing image file' },
      { status: 500 }
    );
  }
}

async function processExcelFile(file: File) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    const excelText = (jsonData as unknown[][])
      .filter((row) => row.some((cell) => cell !== null && cell !== undefined))
      .map((row) => row.map((cell) => cell?.toString() || '').join('\t'))
      .join('\n');

    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "user",
          content: `Please analyze this Excel timetable data and convert it to a structured JSON format. The data is:

${excelText}

Return a JSON object with the following structure:
{
  "title": "Schedule title if identifiable",
  "schedule": {
    "Monday": [{"time": "09:00-10:00", "subject": "Math", "room": "A101"}],
    "Tuesday": [{"time": "09:00-10:00", "subject": "English", "room": "B202"}],
    "Wednesday": [],
    "Thursday": [],
    "Friday": [],
    "Saturday": [],
    "Sunday": []
  }
}

Extract all time slots, subjects, and room information. Organize by weekdays. If the format is unclear, use your best judgment to structure the data appropriately.`
        }
      ],
      max_tokens: 1000
    });

    const content = response.choices[0].message.content;
    let timetableData;
    
    try {
      timetableData = JSON.parse(content || '{}');
    } catch {
      const jsonMatch = content?.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
      if (jsonMatch) {
        timetableData = JSON.parse(jsonMatch[1]);
      } else {
        const simpleJsonMatch = content?.match(/\{[\s\S]*\}/);
        if (simpleJsonMatch) {
          timetableData = JSON.parse(simpleJsonMatch[0]);
        } else {
          throw new Error('Failed to parse OpenAI response as JSON');
        }
      }
    }

    const fileId = `excel_${Object.keys(timetablesStorage).length}`;
    timetablesStorage[fileId] = timetableData;

    return NextResponse.json({
      id: fileId,
      data: timetableData
    });

  } catch (error) {
    console.error('Excel processing error:', error);
    return NextResponse.json(
      { error: 'Error processing Excel file' },
      { status: 500 }
    );
  }
}
