import { NextRequest, NextResponse } from 'next/server';
import OpenAI from 'openai';
import * as XLSX from 'xlsx';
import sharp from 'sharp';

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

interface SubjectMaster {
  [schoolLevel: string]: {
    [grade: string]: {
      [subject: string]: {
        aliases: string[];
        color: string;
      };
    };
  };
}

let subjectMasterCache: SubjectMaster | null = null;

async function loadSubjectMaster(): Promise<SubjectMaster> {
  if (subjectMasterCache) {
    return subjectMasterCache;
  }
  
  try {
    const fs = await import('fs');
    const path = await import('path');
    
    let filePath = path.join(process.cwd(), 'public', 'subject_master_full.json');
    
    if (!fs.existsSync(filePath)) {
      filePath = path.join(process.cwd(), '..', '..', '..', 'public', 'subject_master_full.json');
    }
    
    if (!fs.existsSync(filePath)) {
      filePath = path.join(__dirname, '..', '..', '..', '..', 'public', 'subject_master_full.json');
    }
    
    if (!fs.existsSync(filePath)) {
      try {
        const response = await fetch('/subject_master_full.json');
        if (response.ok) {
          const jsonData = await response.json();
          subjectMasterCache = jsonData;
          return subjectMasterCache!;
        }
      } catch (fetchError) {
        console.error('Fetch fallback failed:', fetchError);
      }
    }
    
    const fileContent = fs.readFileSync(filePath, 'utf-8');
    subjectMasterCache = JSON.parse(fileContent);
    return subjectMasterCache!;
  } catch (error) {
    console.error('Error loading subject master:', error);
    subjectMasterCache = { elementary: {}, junior: {} };
    return subjectMasterCache;
  }
}

function extractColorHex(colorString: string): string {
  const match = colorString.match(/^#[0-9A-Fa-f]{6}/);
  return match ? match[0] : '#E5E7EB';
}

function normalizeSubject(
  subject: string, 
  schoolLevel: string, 
  grade: string, 
  subjectMaster: SubjectMaster
): { normalizedSubject: string; color: string; isUnmatched: boolean } {
  const gradeData = subjectMaster[schoolLevel]?.[grade];
  if (!gradeData) {
    return { normalizedSubject: subject, color: '#EF4444', isUnmatched: true };
  }

  for (const [canonicalSubject, data] of Object.entries(gradeData)) {
    if (data.aliases.includes(subject)) {
      return {
        normalizedSubject: canonicalSubject,
        color: extractColorHex(data.color),
        isUnmatched: false
      };
    }
  }

  for (const [canonicalSubject, data] of Object.entries(gradeData)) {
    if (canonicalSubject.includes(subject) || subject.includes(canonicalSubject)) {
      return {
        normalizedSubject: canonicalSubject,
        color: extractColorHex(data.color),
        isUnmatched: false
      };
    }
  }

  return { normalizedSubject: subject, color: '#EF4444', isUnmatched: true };
}

const timetablesStorage: Record<string, unknown> = {};

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const schoolLevel = formData.get('schoolLevel') as string || 'elementary';
    const grade = formData.get('grade') as string || '1';
    
    if (!file) {
      return NextResponse.json(
        { error: 'No file provided' },
        { status: 400 }
      );
    }

    if (file.type.startsWith('image/')) {
      return await processImageFile(file, schoolLevel, grade);
    } else if (
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel'
    ) {
      return await processExcelFile(file, schoolLevel, grade);
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

async function processImageFile(file: File, schoolLevel: string, grade: string) {
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

    const subjectMaster = await loadSubjectMaster();
    if (timetableData.schedule) {
      for (const [, entries] of Object.entries(timetableData.schedule)) {
        if (Array.isArray(entries)) {
          for (const entry of entries) {
            if (entry.subject) {
              const normalized = normalizeSubject(entry.subject, schoolLevel, grade, subjectMaster);
              entry.normalizedSubject = normalized.normalizedSubject;
              entry.subjectColor = normalized.color;
              entry.isUnmatched = normalized.isUnmatched;
            }
          }
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

async function processExcelFile(file: File, schoolLevel: string, grade: string) {
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

    const subjectMaster = await loadSubjectMaster();
    if (timetableData.schedule) {
      for (const [, entries] of Object.entries(timetableData.schedule)) {
        if (Array.isArray(entries)) {
          for (const entry of entries) {
            if (entry.subject) {
              const normalized = normalizeSubject(entry.subject, schoolLevel, grade, subjectMaster);
              entry.normalizedSubject = normalized.normalizedSubject;
              entry.subjectColor = normalized.color;
              entry.isUnmatched = normalized.isUnmatched;
            }
          }
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
