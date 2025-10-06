import { useState, useEffect } from 'react';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';

interface FileData {
  type: 'docx' | 'xlsx';
  fileName: string;
  sheets?: Array<{ name: string; html: string }>;
}

interface FlutterMessage {
  type: string;
  fileType: 'docx' | 'xlsx';
  fileName: string;
  fileData: string;
}

export default function FileViewer() {
  const [fileData, setFileData] = useState<FileData | null>(null);
  const [content, setContent] = useState<string>('');
  const [excelSheets, setExcelSheets] = useState<Array<{ name: string; html: string }>>([]);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    // Notify parent that viewer is ready
    window.parent.postMessage({ type: 'VIEWER_READY' }, '*');

    // Listen for file data from Flutter WebView
    const handleMessage = async (event: MessageEvent<FlutterMessage>) => {
      if (event.data.type === 'FILE_DATA') {
        const { fileType, fileName, fileData } = event.data;
        
        try {
          setLoading(true);
          setError(null);
          
          // Convert base64 to ArrayBuffer
          const byteCharacters = atob(fileData);
          const byteNumbers = new Array(byteCharacters.length);
          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
          }
          const byteArray = new Uint8Array(byteNumbers);
          const arrayBuffer = byteArray.buffer;

          if (fileType === 'docx') {
            await renderDocx(arrayBuffer, fileName);
          } else if (fileType === 'xlsx') {
            await renderExcel(arrayBuffer, fileName);
          }
          
          setLoading(false);
        } catch (err) {
          console.error('Error processing file:', err);
          setError(err instanceof Error ? err.message : 'Unknown error occurred');
          setLoading(false);
        }
      }
    };

    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const renderDocx = async (arrayBuffer: ArrayBuffer, fileName: string) => {
    const options = {
      convertImage: mammoth.images.imgElement((image: any) => {
        return image.read("base64").then((imageBuffer: string) => {
          return {
            src: `data:${image.contentType};base64,${imageBuffer}`
          };
        });
      })
    };
    
    const htmlResult = await mammoth.convertToHtml({ arrayBuffer }, options);
    const textResult = await mammoth.extractRawText({ arrayBuffer });
    const hasUnderscores = /_+/.test(textResult.value);
    
    let processedHtml = htmlResult.value;
    
    if (hasUnderscores) {
      processedHtml = processedHtml.replace(/<p>(.*?)<\/p>/g, (content) => {
        const processedContent = content.replace(/_{3,}/g, (underscores: string) => {
          return `<span class="underscore-field">${'_'.repeat(underscores.length)}</span>`;
        });
        return `<p>${processedContent}</p>`;
      });
    }
    
    setFileData({ type: 'docx', fileName });
    setContent(processedHtml);
  };

  const renderExcel = async (arrayBuffer: ArrayBuffer, fileName: string) => {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    const sheets = workbook.SheetNames.map(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const htmlTable = XLSX.utils.sheet_to_html(worksheet);
      return { name: sheetName, html: htmlTable };
    });

    setFileData({ type: 'xlsx', fileName, sheets });
    setExcelSheets(sheets);
    setSelectedSheetIndex(0);
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-screen flex-col gap-4 bg-gray-50">
        <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
        <p className="text-gray-600 font-medium">Loading document...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-screen flex-col gap-4 p-6 bg-gray-50">
        <div className="text-6xl">⚠️</div>
        <h2 className="text-2xl font-bold text-red-600">Error Loading File</h2>
        <p className="text-gray-600 text-center max-w-md">{error}</p>
        <button
          onClick={() => window.location.reload()}
          className="px-6 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 transition-colors"
        >
          Try Again
        </button>
      </div>
    );
  }

  if (!fileData) {
    return (
      <div className="flex items-center justify-center min-h-screen flex-col gap-4 bg-gray-50">
        <div className="text-6xl">📄</div>
        <p className="text-gray-600 font-medium">Waiting for file...</p>
      </div>
    );
  }

  // Render DOCX
  if (fileData.type === 'docx') {
    return (
      <div className="min-h-screen bg-gray-50 py-8">
        <div className="max-w-4xl mx-auto px-2">
          <div className="bg-white rounded-lg shadow-lg overflow-hidden">
          <div 
                className="prose prose-sm sm:prose lg:prose-lg max-w-none"
                dangerouslySetInnerHTML={{ __html: content }}
              />
          </div>
        </div>

        <style>{`
          .underscore-field {
            font-family: monospace;
            letter-spacing: 0.05em;
            white-space: pre;
            color: #000;
          }
          
          .prose p {
            margin: 0.75em 0;
            white-space: pre-wrap;
            word-wrap: break-word;
          }
          
          .prose img {
            display: block;
            margin: 1.5em auto;
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          }
          
          .prose table {
            border-collapse: collapse;
            width: 100%;
            margin: 1.5em 0;
            font-size: 0.9em;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
          }
          
          .prose th,
          .prose td {
            border: 1px solid #e5e7eb;
            padding: 10px 12px;
            text-align: left;
          }
          
          .prose th {
            background-color: #f3f4f6;
            font-weight: 600;
            color: #374151;
          }
          
          .prose tr:nth-child(even) {
            background-color: #f9fafb;
          }
        `}</style>
      </div>
    );
  }

  // Render XLSX
  if (fileData.type === 'xlsx' && excelSheets.length > 0) {
    const currentSheet = excelSheets[selectedSheetIndex];

    return (
      <div className="min-h-screen bg-gray-50">
        <div className="bg-white shadow-md sticky top-0 z-10">
          <div className="max-w-7xl mx-auto px-4 py-4">
            
            {excelSheets.length > 1 && (
              <div className="flex gap-2 overflow-x-auto pb-2">
                {excelSheets.map((sheet, idx) => (
                  <button
                    key={idx}
                    onClick={() => setSelectedSheetIndex(idx)}
                    className={`px-4 py-2 rounded-lg font-medium text-sm whitespace-nowrap transition-all ${
                      selectedSheetIndex === idx
                        ? 'bg-green-500 text-white shadow-md'
                        : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                    }`}
                  >
                    {sheet.name}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        
        <div className="max-w-7xl mx-auto p-4">
          <div className="bg-white rounded-lg shadow-lg overflow-hidden">
            <div className="overflow-x-auto">
              <div 
                className="excel-table-container"
                dangerouslySetInnerHTML={{ __html: currentSheet.html }}
              />
            </div>
          </div>
        </div>

        <style>{`
          .excel-table-container table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
          }
          
          .excel-table-container td,
          .excel-table-container th {
            border: 1px solid #d1d5db;
            padding: 10px 12px;
            text-align: left;
            min-width: 100px;
          }
          
          .excel-table-container th {
            background-color: #10b981;
            color: white;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
          }
          
          .excel-table-container tr:nth-child(even) {
            background-color: #f9fafb;
          }
          
          .excel-table-container tr:hover {
            background-color: #f3f4f6;
          }
          
          @media (max-width: 640px) {
            .excel-table-container table {
              font-size: 11px;
            }
            
            .excel-table-container td,
            .excel-table-container th {
              padding: 6px 8px;
              min-width: 80px;
            }
          }
        `}</style>
      </div>
    );
  }

  return null;
}

