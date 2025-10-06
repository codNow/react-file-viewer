import { useState, type ChangeEvent } from 'react';
import { Upload, X, FileText, Table, FileSpreadsheet } from 'lucide-react';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';

function App() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [fileContent, setFileContent] = useState<string>('');
  const [fileType, setFileType] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [excelData, setExcelData] = useState<any[]>([]);

  const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    
    if (!file) return;

    setError('');
    setIsLoading(true);
    setSelectedFile(file);
    setFileContent('');
    setExcelData([]);

    try {
      const fileExtension = file.name.split('.').pop()?.toLowerCase();

      if (fileExtension === 'docx') {
        // Handle Word documents with Mammoth.js
        setFileType('word');
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        setFileContent(result.value);
      } 
      else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
        // Handle Excel files with SheetJS
        setFileType('excel');
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Convert all sheets to HTML tables
        const sheets = workbook.SheetNames.map(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const htmlTable = XLSX.utils.sheet_to_html(worksheet);
          return { name: sheetName, html: htmlTable };
        });
        
        setExcelData(sheets);
      }
      else if (fileExtension === 'pdf') {
        setFileType('pdf');
        const url = URL.createObjectURL(file);
        setFileContent(url);
      }
      else if (fileExtension === 'doc' || fileExtension === 'ppt' || fileExtension === 'pptx') {
        setError(`${fileExtension.toUpperCase()} files require conversion. Please save as DOCX (for Word) or use Google Docs viewer.`);
      }
      else {
        setError('Unsupported file type. Please upload .docx, .xlsx, .xls, or .pdf files.');
      }
    } catch (err) {
      setError(`Error loading file: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleClearFile = () => {
    setSelectedFile(null);
    setFileContent('');
    setExcelData([]);
    setFileType('');
    setError('');
  };

  return (
    <div className="bg-gray-50 min-h-screen">
      <div className="container mx-auto px-4 py-8">
        <div className="max-w-6xl mx-auto">
          <h1 className="text-3xl font-bold text-gray-800 mb-2 text-center">
            Office Document Viewer
          </h1>
          <p className="text-center text-gray-600 mb-8">
            Upload Word (.docx), Excel (.xlsx, .xls), or PDF files
          </p>

          {!selectedFile ? (
            <div className="max-w-2xl mx-auto mb-8">
              <div className="bg-white rounded-xl shadow-lg p-8 border-2 border-dashed border-gray-300 hover:border-blue-400 transition-colors">
                <div className="text-center">
                  <Upload className="w-16 h-16 text-gray-400 mx-auto mb-4" />
                  <label htmlFor="file-upload" className="cursor-pointer">
                    <span className="text-lg font-medium text-gray-700 hover:text-blue-600 transition-colors inline-block">
                      Choose a document to upload
                    </span>
                    <input
                      id="file-upload"
                      type="file"
                      className="hidden"
                      accept=".docx,.xlsx,.xls,.pdf"
                      onChange={handleFileUpload}
                      disabled={isLoading}
                    />
                  </label>
                  
                  <div className="mt-6 flex justify-center gap-4">
                    <div className="flex items-center gap-2 text-sm text-gray-600">
                      <FileText className="w-4 h-4 text-blue-500" />
                      <span>Word</span>
                    </div>
                    <div className="flex items-center gap-2 text-sm text-gray-600">
                      <Table className="w-4 h-4 text-green-500" />
                      <span>Excel</span>
                    </div>
                    <div className="flex items-center gap-2 text-sm text-gray-600">
                      <FileSpreadsheet className="w-4 h-4 text-red-500" />
                      <span>PDF</span>
                    </div>
                  </div>

                  {error && (
                    <p className="text-red-600 text-sm mt-4 bg-red-50 p-3 rounded-lg">
                      {error}
                    </p>
                  )}
                  {isLoading && (
                    <div className="mt-4">
                      <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-blue-500 border-t-transparent"></div>
                      <p className="text-blue-600 text-sm mt-2">Loading document...</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="bg-white rounded-lg shadow p-4 flex items-center justify-between">
                <div className="flex items-center space-x-3">
                  <div className="bg-blue-100 p-2 rounded">
                    {fileType === 'word' && <FileText className="w-5 h-5 text-blue-600" />}
                    {fileType === 'excel' && <Table className="w-5 h-5 text-green-600" />}
                    {fileType === 'pdf' && <FileSpreadsheet className="w-5 h-5 text-red-600" />}
                  </div>
                  <div>
                    <p className="font-medium text-gray-800">{selectedFile.name}</p>
                    <p className="text-sm text-gray-500">
                      {(selectedFile.size / 1024).toFixed(2)} KB
                    </p>
                  </div>
                </div>
                <button
                  onClick={handleClearFile}
                  className="bg-red-50 hover:bg-red-100 text-red-600 p-2 rounded-lg transition-colors"
                  title="Remove document"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {error && (
                <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-lg">
                  {error}
                </div>
              )}

              {/* Word Document Display */}
              {fileType === 'word' && fileContent && (
                <div className="bg-white rounded-lg shadow-lg p-8 overflow-auto" style={{ maxHeight: '80vh' }}>
                  <div 
                    className="prose max-w-none"
                    dangerouslySetInnerHTML={{ __html: fileContent }}
                  />
                </div>
              )}

              {/* Excel Document Display */}
              {fileType === 'excel' && excelData.length > 0 && (
                <div className="bg-white rounded-lg shadow-lg overflow-auto" style={{ maxHeight: '80vh' }}>
                  {excelData.map((sheet, idx) => (
                    <div key={idx} className="p-6 border-b last:border-b-0">
                      <h3 className="text-lg font-semibold text-gray-800 mb-4 bg-gray-100 px-4 py-2 rounded">
                        Sheet: {sheet.name}
                      </h3>
                      <div 
                        className="overflow-x-auto"
                        dangerouslySetInnerHTML={{ __html: sheet.html }}
                        style={{
                          fontSize: '14px'
                        }}
                      />
                    </div>
                  ))}
                </div>
              )}

              {/* PDF Display */}
              {fileType === 'pdf' && fileContent && (
                <div className="bg-white rounded-lg shadow-lg overflow-hidden" style={{ height: '80vh' }}>
                  <iframe
                    src={fileContent}
                    className="w-full h-full"
                    title="PDF Viewer"
                  />
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <style>{`
        /* Word document styles */
        .docx-wrapper {
          background: white;
          padding: 20px;
          box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        
        @media (min-width: 640px) {
          .docx-wrapper {
            padding: 40px;
          }
        }
        
        .docx-wrapper section.docx {
          background: white;
          margin-bottom: 20px;
          padding: 10px;
          max-width: 100%;
          overflow-x: auto;
        }
        
        @media (min-width: 640px) {
          .docx-wrapper section.docx {
            padding: 20px;
          }
        }
        
        /* Make docx content responsive */
        .docx-wrapper * {
          max-width: 100% !important;
        }
        
        .docx-wrapper img {
          height: auto !important;
        }
        
        .docx-wrapper table {
          font-size: 12px;
        }
        
        @media (min-width: 640px) {
          .docx-wrapper table {
            font-size: 14px;
          }
        }
        
        /* Excel table styles */
        .prose table {
          border-collapse: collapse;
          width: 100%;
          margin: 1em 0;
          font-size: 11px;
        }
        
        @media (min-width: 640px) {
          .prose table {
            font-size: 13px;
          }
        }
        
        .prose th,
        .prose td {
          border: 1px solid #e5e7eb;
          padding: 4px 6px;
          text-align: left;
        }
        
        @media (min-width: 640px) {
          .prose th,
          .prose td {
            padding: 8px 12px;
          }
        }
        
        .prose th {
          background-color: #f9fafb;
          font-weight: 600;
        }
        table {
          border-collapse: collapse;
          width: 100%;
          font-size: 11px;
        }
        
        @media (min-width: 640px) {
          table {
            font-size: 13px;
          }
        }
        
        table td,
        table th {
          border: 1px solid #ddd;
          padding: 4px 6px;
        }
        
        @media (min-width: 640px) {
          table td,
          table th {
            padding: 8px;
          }
        }
        
        table tr:nth-child(even) {
          background-color: #f9fafb;
        }
        table th {
          background-color: #3b82f6;
          color: white;
          font-weight: bold;
        }
      `}</style>
    </div>
  );
}

export default App;