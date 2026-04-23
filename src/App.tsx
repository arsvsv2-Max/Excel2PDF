import React, { useState, useCallback, useRef, useEffect } from 'react';
import { 
  FileUp, 
  FileText, 
  Table, 
  ArrowRight, 
  Sparkles, 
  CheckCircle2, 
  Download, 
  RefreshCw, 
  Trash2, 
  ChevronRight,
  Loader2,
  AlertCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  parseExcel, 
  parsePDFFields, 
  parseMappingExcel,
  generateBatch, 
  ExcelData, 
  PDFFieldInfo 
} from './services/fileUtils';
import { getSmartMapping, MappingResult } from './services/geminiService';

type AppState = 'upload' | 'mapping' | 'generating' | 'finished';

export default function App() {
  const [state, setState] = useState<AppState>('upload');
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [pdfFields, setPdfFields] = useState<PDFFieldInfo[]>([]);
  const [mapping, setMapping] = useState<Record<string, string[]>>({});
  const [filenamePattern, setFilenamePattern] = useState('[Efternamn], [Förnamn]');
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [zipBlob, setZipBlob] = useState<Blob | null>(null);
  const [generatedFiles, setGeneratedFiles] = useState<{ name: string; blob: Blob }[]>([]);
  const [isAiMapping, setIsAiMapping] = useState(false);
  const [isMappingLoading, setIsMappingLoading] = useState(false);
  const [messageBox, setMessageBox] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [filterColumn, setFilterColumn] = useState<string>('');
  const [filterValue, setFilterValue] = useState<string>('');

  const fileInputPdf = useRef<HTMLInputElement>(null);
  const fileInputExcel = useRef<HTMLInputElement>(null);
  const fileInputMapping = useRef<HTMLInputElement>(null);

  // Load configuration on mount
  useEffect(() => {
    const savedPattern = localStorage.getItem('excel2pdf_pattern');
    if (savedPattern) {
      setFilenamePattern(savedPattern);
    }
  }, []);

  // Save filename pattern on change
  useEffect(() => {
    localStorage.setItem('excel2pdf_pattern', filenamePattern);
  }, [filenamePattern]);

  const handlePdfUpload = async (file: File) => {
    if (!file.name.endsWith('.pdf')) {
      setError('Please upload a valid PDF file.');
      return;
    }
    try {
      const fields = await parsePDFFields(file);
      setPdfFields(fields);
      setPdfFile(file);
      setError(null);
    } catch (err) {
      setError('Failed to parse PDF fields.');
    }
  };

  const handleExcelUpload = async (file: File) => {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setError('Please upload a valid Excel file.');
      return;
    }
    try {
      const data = await parseExcel(file);
      setExcelData(data);
      setExcelFile(file);
      
      // Auto-set filter for Anstnr if it exists (as requested)
      const anstnrHeader = data.headers.find(h => h.toLowerCase() === 'anstnr');
      if (anstnrHeader) {
        setFilterColumn(anstnrHeader);
        setFilterValue('10103450');
      }
      
      setError(null);
    } catch (err: any) {
      setError(err.message || 'Failed to parse Excel data.');
    }
  };

  const handleSheetChange = async (sheetName: string) => {
    if (!excelFile) return;
    try {
      const data = await parseExcel(excelFile, sheetName);
      setExcelData(data);
      setError(null);
    } catch (err: any) {
      setError(err.message || 'Failed to parse selected sheet.');
    }
  };

  const handleSmartMap = async () => {
    if (!excelData || pdfFields.length === 0) return;
    setIsAiMapping(true);
    try {
      const results: MappingResult[] = await getSmartMapping(
        pdfFields.map(f => f.name),
        excelData.headers
      );
      
      const newMapping = { ...mapping };
      results.forEach(res => {
        if (!newMapping[res.pdfField]) {
          newMapping[res.pdfField] = [res.excelHeader];
        } else if (!newMapping[res.pdfField].includes(res.excelHeader)) {
          newMapping[res.pdfField].push(res.excelHeader);
        }
      });
      setMapping(newMapping);
    } catch (err) {
      setError('AI mapping failed. Please try manual mapping.');
    } finally {
      setIsAiMapping(false);
    }
  };

  const handleMappingUpload = async (file: File) => {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setError('Please upload a valid Excel file for mapping.');
      return;
    }
    setIsMappingLoading(true);
    try {
      const { mapping: newMapping, firstDocField } = await parseMappingExcel(file);
      setMapping(prev => ({ ...prev, ...newMapping }));
      if (firstDocField) {
        setMessageBox(`First Docfield found: ${firstDocField}`);
      }
      setError(null);
    } catch (err: any) {
      setError(err.message || 'Failed to parse mapping file.');
    } finally {
      setIsMappingLoading(false);
    }
  };

  const handleGenerate = async () => {
    if (!pdfFile || !excelData) return;
    setState('generating');
    try {
      let filteredRows = excelData.rows;
      if (filterColumn && filterValue) {
        filteredRows = excelData.rows.filter(row => 
          String(row[filterColumn]).trim() === filterValue.trim()
        );
      }

      if (filteredRows.length === 0) {
        setError(filterColumn ? `No records found where ${filterColumn} = ${filterValue}` : 'No records to process');
        setState('mapping');
        return;
      }

      const filteredData = { ...excelData, rows: filteredRows };

      const result = await generateBatch(
        pdfFile,
        filteredData,
        mapping,
        filenamePattern,
        (current, total) => setProgress({ current, total })
      );
      setZipBlob(result.zip);
      setGeneratedFiles(result.files);
      setState('finished');
    } catch (err) {
      setError('Generation failed.');
      setState('mapping');
    }
  };

  const resetApp = () => {
    setPdfFile(null);
    setExcelFile(null);
    setExcelData(null);
    setPdfFields([]);
    setMapping({});
    setFilenamePattern('[Efternamn], [Förnamn]');
    setZipBlob(null);
    setGeneratedFiles([]);
    setState('upload');
    setError(null);
  };

  const downloadZip = () => {
    if (!zipBlob) return;
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Excel2PDF_Batch_${new Date().getTime()}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const downloadPdf = (file: { name: string; blob: Blob }) => {
    const url = URL.createObjectURL(file.blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = file.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-50 font-sans selection:bg-indigo-500/30">
      <header className="border-b border-slate-800/50 bg-slate-950/50 backdrop-blur-xl sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shadow-lg shadow-indigo-500/20">
              <FileText className="text-white w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">Excel2PDF AI</h1>
              <p className="text-xs text-slate-400 font-medium uppercase tracking-widest">Bulk PDF Automation</p>
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            {state !== 'upload' && (
              <button 
                onClick={resetApp}
                className="p-2 hover:bg-slate-800 rounded-lg transition-colors text-slate-400 hover:text-white"
                title="Start New Batch"
              >
                <RefreshCw className="w-5 h-5" />
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-12">
        <AnimatePresence mode="wait">
          {state === 'upload' && (
            <motion.div 
              key="upload"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8"
            >
              <div className="text-center space-y-2">
                <h2 className="text-3xl font-bold">Upload your files</h2>
                <p className="text-slate-400">Start by providing your PDF template and Excel data source.</p>
              </div>

              <div className="grid md:grid-cols-2 gap-6">
                {/* PDF Upload */}
                <div 
                  onClick={() => fileInputPdf.current?.click()}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    const file = e.dataTransfer.files[0];
                    if (file) handlePdfUpload(file);
                  }}
                  className={`group relative border-2 border-dashed rounded-3xl p-10 flex flex-col items-center justify-center gap-4 transition-all cursor-pointer
                    ${pdfFile ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-slate-800 hover:border-indigo-500/50 hover:bg-indigo-500/5'}`}
                >
                  <input 
                    type="file" 
                    ref={fileInputPdf} 
                    className="hidden" 
                    accept=".pdf" 
                    onChange={(e) => e.target.files?.[0] && handlePdfUpload(e.target.files[0])} 
                  />
                  <div className={`w-16 h-16 rounded-2xl flex items-center justify-center transition-transform group-hover:scale-110
                    ${pdfFile ? 'bg-emerald-500 text-white' : 'bg-slate-900 text-slate-400 group-hover:text-indigo-400'}`}>
                    <FileText className="w-8 h-8" />
                  </div>
                  <div className="text-center">
                    <p className="font-semibold">{pdfFile ? pdfFile.name : 'Select PDF Form'}</p>
                    <p className="text-sm text-slate-500">{pdfFile ? `${pdfFields.length} fields detected` : 'Drag & drop fillable PDF'}</p>
                  </div>
                  {pdfFile && (
                    <div className="absolute top-4 right-4 text-emerald-500">
                      <CheckCircle2 className="w-6 h-6" />
                    </div>
                  )}
                </div>

                {/* Excel Upload */}
                <div 
                  onClick={() => fileInputExcel.current?.click()}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    const file = e.dataTransfer.files[0];
                    if (file) handleExcelUpload(file);
                  }}
                  className={`group relative border-2 border-dashed rounded-3xl p-10 flex flex-col items-center justify-center gap-4 transition-all cursor-pointer
                    ${excelFile ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-slate-800 hover:border-indigo-500/50 hover:bg-indigo-500/5'}`}
                >
                  <input 
                    type="file" 
                    ref={fileInputExcel} 
                    className="hidden" 
                    accept=".xlsx,.xls" 
                    onChange={(e) => e.target.files?.[0] && handleExcelUpload(e.target.files[0])} 
                  />
                  <div className={`w-16 h-16 rounded-2xl flex items-center justify-center transition-transform group-hover:scale-110
                    ${excelFile ? 'bg-emerald-500 text-white' : 'bg-slate-900 text-slate-400 group-hover:text-indigo-400'}`}>
                    <Table className="w-8 h-8" />
                  </div>
                  <div className="text-center">
                    <p className="font-semibold">{excelFile ? excelFile.name : 'Select Excel Data'}</p>
                    <p className="text-sm text-slate-500">{excelFile ? `${excelData?.rows.length} rows found` : 'Drag & drop .xlsx or .xls'}</p>
                  </div>
                  {excelFile && excelData && excelData.sheetNames.length > 1 && (
                    <div className="mt-4 w-full px-4" onClick={(e) => e.stopPropagation()}>
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-1">Select Worksheet</label>
                      <select 
                        value={excelData.selectedSheet}
                        onChange={(e) => handleSheetChange(e.target.value)}
                        className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/50"
                      >
                        {excelData.sheetNames.map(name => (
                          <option key={name} value={name}>{name}</option>
                        ))}
                      </select>
                    </div>
                  )}
                  {excelFile && (
                    <div className="absolute top-4 right-4 text-emerald-500">
                      <CheckCircle2 className="w-6 h-6" />
                    </div>
                  )}
                </div>
              </div>

              {error && (
                <div className="bg-red-500/10 border border-red-500/20 text-red-400 px-4 py-3 rounded-xl flex items-center gap-3">
                  <AlertCircle className="w-5 h-5" />
                  <p className="text-sm">{error}</p>
                </div>
              )}

              <div className="flex justify-center pt-4">
                <button
                  disabled={!pdfFile || !excelFile}
                  onClick={() => setState('mapping')}
                  className="px-8 py-4 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold rounded-2xl flex items-center gap-2 transition-all shadow-xl shadow-indigo-500/20 active:scale-95"
                >
                  Continue to Mapping
                  <ArrowRight className="w-5 h-5" />
                </button>
              </div>
            </motion.div>
          )}

          {state === 'mapping' && (
            <motion.div 
              key="mapping"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8"
            >
                <div className="flex flex-wrap items-end justify-between gap-4">
                  <div className="space-y-2">
                    <h2 className="text-3xl font-bold">Configure Mapping</h2>
                    <p className="text-slate-400">Map Excel columns to your PDF form fields.</p>
                  </div>
                  <div className="flex gap-3">
                    <input 
                      type="file" 
                      ref={fileInputMapping} 
                      className="hidden" 
                      accept=".xlsx,.xls" 
                      onChange={(e) => e.target.files?.[0] && handleMappingUpload(e.target.files[0])} 
                    />
                    <button
                      onClick={() => fileInputMapping.current?.click()}
                      disabled={isMappingLoading}
                      className="px-6 py-3 bg-slate-800 hover:bg-slate-700 text-emerald-400 font-semibold rounded-xl flex items-center gap-2 transition-all border border-slate-700"
                    >
                      {isMappingLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : <FileUp className="w-5 h-5" />}
                      Mappning körtillstånd
                    </button>
                    <button
                      onClick={handleSmartMap}
                      disabled={isAiMapping}
                      className="px-6 py-3 bg-slate-800 hover:bg-slate-700 text-indigo-400 font-semibold rounded-xl flex items-center gap-2 transition-all border border-slate-700"
                    >
                      {isAiMapping ? <Loader2 className="w-5 h-5 animate-spin" /> : <Sparkles className="w-5 h-5" />}
                      Smart Auto-Map
                    </button>
                  </div>
                </div>

              {/* Filename Pattern */}
              <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-8 space-y-6">
                <div className="grid md:grid-cols-2 gap-8">
                  <div className="space-y-2">
                    <label className="text-sm font-bold text-slate-400 uppercase tracking-wider">Output Filename Pattern</label>
                    <div className="relative">
                      <input 
                        type="text"
                        value={filenamePattern}
                        onChange={(e) => setFilenamePattern(e.target.value)}
                        className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 focus:outline-none focus:ring-2 focus:ring-indigo-500/50 transition-all font-mono text-sm"
                        placeholder="e.g. Invoice_[ID]_[Name]"
                      />
                    </div>
                    <div className="flex flex-wrap gap-2 mt-3">
                      {excelData?.headers.slice(0, 8).map(header => (
                        <button
                          key={header}
                          onClick={() => setFilenamePattern(prev => `${prev}[${header}]`)}
                          className="px-3 py-1 bg-slate-800 hover:bg-slate-700 text-xs rounded-full border border-slate-700 transition-colors"
                        >
                          +{header}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-sm font-bold text-slate-400 uppercase tracking-wider">Filter Row (Optional)</label>
                    <div className="flex gap-2">
                      <select 
                        value={filterColumn}
                        onChange={(e) => setFilterColumn(e.target.value)}
                        className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-3 focus:outline-none focus:ring-2 focus:ring-indigo-500/50 transition-all text-sm w-1/3"
                      >
                        <option value="">No Filter</option>
                        {excelData?.headers.map(h => <option key={h} value={h}>{h}</option>)}
                      </select>
                      <input 
                        type="text"
                        value={filterValue}
                        onChange={(e) => setFilterValue(e.target.value)}
                        className="flex-1 bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 focus:outline-none focus:ring-2 focus:ring-indigo-500/50 transition-all text-sm"
                        placeholder="Value to match..."
                      />
                    </div>
                    <p className="text-[10px] text-slate-500 italic">Only rows matching this criteria will be processed.</p>
                  </div>
                </div>
              </div>

              {/* Field Mapping List */}
              <div className="space-y-4">
                <h3 className="text-lg font-bold flex items-center gap-2">
                  <FileUp className="w-5 h-5 text-indigo-500" />
                  Field Assignments
                </h3>
                <div className="grid gap-3">
                  {pdfFields.map((field) => (
                    <div key={field.name} className="bg-slate-900/50 border border-slate-800 rounded-2xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-slate-400">
                          {field.type === 'checkbox' ? <CheckCircle2 className="w-4 h-4" /> : <FileText className="w-4 h-4" />}
                        </div>
                        <div>
                          <p className="font-semibold text-sm">{field.name}</p>
                          <p className="text-xs text-slate-500 capitalize">{field.type} Field</p>
                        </div>
                      </div>
                      
                      <div className="flex flex-wrap items-center gap-2 md:w-2/3 justify-end">
                        {mapping[field.name]?.map((mappedHeader, idx) => (
                          <div key={idx} className="flex items-center gap-1 bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 px-3 py-1.5 rounded-lg text-sm">
                            {mappedHeader}
                            <button 
                              onClick={() => {
                                const newMapping = { ...mapping };
                                newMapping[field.name] = newMapping[field.name].filter((_, i) => i !== idx);
                                setMapping(newMapping);
                              }}
                              className="hover:text-red-400 transition-colors"
                            >
                              <Trash2 className="w-3 h-3" />
                            </button>
                          </div>
                        ))}
                        <select
                          className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/50"
                          value=""
                          onChange={(e) => {
                            if (!e.target.value) return;
                            const newMapping = { ...mapping };
                            if (!newMapping[field.name]) newMapping[field.name] = [];
                            if (!newMapping[field.name].includes(e.target.value)) {
                              newMapping[field.name].push(e.target.value);
                            }
                            setMapping(newMapping);
                          }}
                        >
                          <option value="">Add Header...</option>
                          {excelData?.headers.map(header => (
                            <option key={header} value={header}>{header}</option>
                          ))}
                        </select>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex justify-between pt-8">
                <button
                  onClick={() => setState('upload')}
                  className="px-6 py-3 text-slate-400 hover:text-white font-semibold transition-colors"
                >
                  Back to Upload
                </button>
                <button
                  onClick={handleGenerate}
                  className="px-10 py-4 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-2xl flex items-center gap-2 transition-all shadow-xl shadow-emerald-500/20 active:scale-95"
                >
                  Generate Files
                  <ChevronRight className="w-5 h-5" />
                </button>
              </div>
            </motion.div>
          )}

          {state === 'generating' && (
            <motion.div 
              key="generating"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className="max-w-md mx-auto space-y-8 py-20 text-center"
            >
              <div className="relative inline-block">
                <div className="w-24 h-24 border-4 border-slate-800 border-t-indigo-500 rounded-full animate-spin"></div>
                <div className="absolute inset-0 flex items-center justify-center">
                  <FileText className="w-8 h-8 text-indigo-500" />
                </div>
              </div>
              <div className="space-y-4">
                <h2 className="text-2xl font-bold">Processing Batch</h2>
                <p className="text-slate-400">Please wait while we generate your documents...</p>
                <div className="space-y-2">
                  <div className="h-3 w-full bg-slate-900 rounded-full overflow-hidden border border-slate-800">
                    <motion.div 
                      className="h-full bg-indigo-500"
                      initial={{ width: 0 }}
                      animate={{ width: `${(progress.current / progress.total) * 100}%` }}
                    />
                  </div>
                  <p className="text-sm font-mono text-slate-500">
                    {progress.current} of {progress.total} processed
                  </p>
                </div>
              </div>
            </motion.div>
          )}

          {state === 'finished' && (
            <motion.div 
              key="finished"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="max-w-2xl mx-auto space-y-8 py-12"
            >
              <div className="text-center space-y-4">
                <div className="w-20 h-20 bg-emerald-500/10 text-emerald-500 rounded-full flex items-center justify-center mx-auto border border-emerald-500/20">
                  <CheckCircle2 className="w-10 h-10" />
                </div>
                <div className="space-y-2">
                  <h2 className="text-3xl font-bold text-center">Batch Complete!</h2>
                  <p className="text-slate-400 text-center">Successfully generated {progress.total} PDF documents.</p>
                </div>
              </div>

              <div className="grid md:grid-cols-2 gap-8">
                {/* ZIP Download */}
                <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-8 flex flex-col items-center gap-6">
                  <div className="w-16 h-16 bg-indigo-500/10 text-indigo-500 rounded-2xl flex items-center justify-center">
                    <Download className="w-8 h-8" />
                  </div>
                  <div className="text-center">
                    <h3 className="font-bold text-lg">Download All</h3>
                    <p className="text-sm text-slate-500">Get all documents in a single ZIP archive.</p>
                  </div>
                  <button
                    onClick={downloadZip}
                    className="w-full py-4 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-2xl flex items-center justify-center gap-3 transition-all shadow-xl shadow-indigo-500/20 active:scale-95"
                  >
                    Download ZIP
                  </button>
                </div>

                {/* Individual PDF Download */}
                <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-8 flex flex-col gap-6">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-emerald-500/10 text-emerald-500 rounded-xl flex items-center justify-center">
                      <FileText className="w-6 h-6" />
                    </div>
                    <div>
                      <h3 className="font-bold">Individual PDFs</h3>
                      <p className="text-xs text-slate-500">{generatedFiles.length} files generated</p>
                    </div>
                  </div>
                  
                  <div className="max-h-[200px] overflow-y-auto pr-2 space-y-2 custom-scrollbar">
                    {generatedFiles.map((file, idx) => (
                      <div key={idx} className="flex items-center justify-between bg-slate-950 border border-slate-800 p-3 rounded-xl group">
                        <span className="text-xs font-mono truncate max-w-[180px] text-slate-400 group-hover:text-slate-200 transition-colors">{file.name}</span>
                        <button 
                          onClick={() => downloadPdf(file)}
                          className="p-1.5 hover:bg-emerald-500/10 text-emerald-500 rounded-lg transition-colors"
                          title="Download PDF"
                        >
                          <Download className="w-4 h-4" />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
              
              <div className="flex justify-center gap-4">
                <button
                  onClick={() => setState('mapping')}
                  className="px-8 py-4 bg-slate-900 hover:bg-slate-800 text-slate-300 font-semibold rounded-2xl border border-slate-800 transition-all"
                >
                  Back to Mapping
                </button>
                <button
                  onClick={resetApp}
                  className="px-8 py-4 bg-slate-900 hover:bg-slate-800 text-slate-300 font-semibold rounded-2xl border border-slate-800 transition-all"
                >
                  New Batch
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
      
      <footer className="max-w-7xl mx-auto px-6 py-12 border-t border-slate-900 text-center">
        <p className="text-slate-600 text-sm">Powered by Gemini AI & PDF-Lib</p>
      </footer>

      {/* MessageBox Modal */}
      <AnimatePresence>
        {messageBox && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex items-center justify-center p-6 bg-slate-950/80 backdrop-blur-sm"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.9, y: 20 }}
              className="bg-slate-900 border border-slate-800 rounded-3xl p-8 max-w-sm w-full shadow-2xl space-y-6"
            >
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 bg-indigo-500/10 text-indigo-500 rounded-2xl flex items-center justify-center">
                  <FileText className="w-6 h-6" />
                </div>
                <h3 className="text-xl font-bold">Mapping Info</h3>
              </div>
              <p className="text-slate-300">{messageBox}</p>
              <button 
                onClick={() => setMessageBox(null)}
                className="w-full py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-xl transition-all active:scale-95"
              >
                Close
              </button>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
