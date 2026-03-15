import React, { useState, useCallback, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import * as CFB from 'cfb';
import { 
  FileSpreadsheet, 
  ShieldAlert, 
  ShieldCheck, 
  Trash2, 
  Download, 
  FileSearch, 
  FolderOpen,
  ChevronRight,
  AlertCircle,
  Info,
  CheckCircle2,
  X,
  Play,
  Code2,
  Terminal,
  Search,
  FileText,
  Cpu,
  Loader2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { GoogleGenAI } from "@google/genai";
import Markdown from 'react-markdown';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface MacroInfo {
  name: string;
  type: 'Module' | 'Class' | 'Sheet' | 'Unknown';
  content?: string;
  selected?: boolean;
}

interface ExcelFile {
  id: string;
  name: string;
  path: string;
  size: number;
  type: string;
  lastModified: number;
  data: ArrayBuffer;
  hasMacros: boolean;
  macros: MacroInfo[];
  structure: string[];
  status: 'pending' | 'analyzing' | 'ready' | 'cleaned' | 'error';
  errorMessage?: string;
  sandboxResult?: string;
  isSimulating?: boolean;
  uploadId: string;
}

export default function App() {
  const [files, setFiles] = useState<ExcelFile[]>([]);
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [checkedFileIds, setCheckedFileIds] = useState<string[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [scanResultModal, setScanResultModal] = useState<{ isOpen: boolean; message: string; found: boolean } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  const selectedFile = files.find(f => f.id === selectedFileId);

  const rescanAllFiles = async () => {
    // Reset status for all files
    setFiles(prev => prev.map(f => ({ ...f, status: 'analyzing' })));
    
    // Process sequentially to avoid overwhelming memory
    for (const file of files) {
      const analyzed = await analyzeFile(file);
      setFiles(prev => prev.map(f => f.id === file.id ? analyzed : f));
    }
  };

  const scanKangatangVirus = () => {
    const virusFileName = 'mypersonnel.xls';
    const found = files.some(f => f.name.toLowerCase() === virusFileName.toLowerCase());
    
    if (found) {
      setFiles(prev => prev.filter(f => f.name.toLowerCase() !== virusFileName.toLowerCase()));
      setScanResultModal({
        isOpen: true,
        message: `Kangatang Virus detected! Removed ${virusFileName} from the list.`,
        found: true
      });
    } else {
      setScanResultModal({
        isOpen: true,
        message: 'No Kangatang Virus detected. Your workspace is clean.',
        found: false
      });
    }
  };

  const toggleFileCheck = (id: string) => {
    setCheckedFileIds(prev => 
      prev.includes(id) ? prev.filter(fid => fid !== id) : [...prev, id]
    );
  };

  const toggleAllFilesCheck = (checked: boolean) => {
    if (checked) {
      setCheckedFileIds(files.map(f => f.id));
    } else {
      setCheckedFileIds([]);
    }
  };

  const downloadSelectedFiles = async () => {
    const selected = files.filter(f => checkedFileIds.includes(f.id));
    if (selected.length === 0) return;

    if (selected.length === 1) {
      downloadFile(selected[0]);
      return;
    }

    // Create a zip for multiple files
    const zip = new JSZip();
    selected.forEach(f => {
      zip.file(f.name, f.data);
    });

    const content = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = `sentinel_export_${new Date().getTime()}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const removeSelectedFilesFromList = () => {
    setFiles(prev => prev.filter(f => !checkedFileIds.includes(f.id)));
    setCheckedFileIds([]);
    if (selectedFileId && checkedFileIds.includes(selectedFileId)) {
      setSelectedFileId(null);
    }
  };

  const toggleMacroSelection = (fileId: string, macroName: string) => {
    setFiles(prev => prev.map(f => {
      if (f.id !== fileId) return f;
      return {
        ...f,
        macros: f.macros.map(m => m.name === macroName ? { ...m, selected: !m.selected } : m)
      };
    }));
  };

  const toggleAllMacros = (fileId: string, selected: boolean) => {
    setFiles(prev => prev.map(f => {
      if (f.id !== fileId) return f;
      return {
        ...f,
        macros: f.macros.map(m => ({ ...m, selected }))
      };
    }));
  };

  const deleteSelectedMacros = async (fileId: string) => {
    const file = files.find(f => f.id === fileId);
    if (!file) return;

    const selectedMacros = file.macros.filter(m => m.selected);
    if (selectedMacros.length === 0) return;

    setFiles(prev => prev.map(f => f.id === fileId ? { ...f, status: 'analyzing' } : f));

    try {
      const zip = new JSZip();
      const contents = await zip.loadAsync(file.data);
      
      // Find the VBA project file
      const vbaFile = contents.file(/vbaProject\.bin$/i)[0];
      if (!vbaFile) throw new Error("No VBA project found in this file structure.");

      const vbaBuffer = await vbaFile.async('arraybuffer');
      const cfb = CFB.read(new Uint8Array(vbaBuffer), { type: 'array' });

      // Surgically remove selected macro streams from the CFB structure
      selectedMacros.forEach(macro => {
        // Find the exact path in the CFB (usually /VBA/ModuleName)
        const cfbPath = cfb.FullPaths.find(p => p.endsWith('/' + macro.name));
        if (cfbPath) {
          // CFB.utils.file_del might be missing in some versions, 
          // we can filter the FileIndex directly as a fallback
          const index = cfb.FullPaths.indexOf(cfbPath);
          if (index > -1) {
            cfb.FileIndex.splice(index, 1);
            cfb.FullPaths.splice(index, 1);
          }
        }
      });

      // Write the modified CFB back to a buffer
      const newVbaBuffer = CFB.write(cfb, { type: 'array' });
      
      // Update the zip content with the modified VBA project
      contents.file(vbaFile.name, newVbaBuffer);

      // Generate the new file data
      const newData = await contents.generateAsync({ type: 'arraybuffer' });
      
      const remainingMacros = file.macros.filter(m => !m.selected);

      setFiles(prev => prev.map(f => f.id === fileId ? { 
        ...f, 
        data: newData,
        macros: remainingMacros,
        hasMacros: remainingMacros.length > 0,
        status: 'ready',
        sandboxResult: undefined // Reset sandbox as code has changed
      } : f));
    } catch (err) {
      console.error("Surgical deletion error:", err);
      setFiles(prev => prev.map(f => f.id === fileId ? { 
        ...f, 
        status: 'error', 
        errorMessage: 'Failed to delete specific macros: ' + (err as Error).message 
      } : f));
    }
  };

  const extractMacros = async (data: ArrayBuffer): Promise<MacroInfo[]> => {
    const macros: MacroInfo[] = [];
    try {
      const zip = new JSZip();
      const contents = await zip.loadAsync(data);
      
      // Look for vbaProject.bin
      const vbaFile = contents.file(/vbaProject\.bin$/i)[0];
      if (vbaFile) {
        const vbaBuffer = await vbaFile.async('arraybuffer');
        const cfb = CFB.read(new Uint8Array(vbaBuffer), { type: 'array' });
        
        // CFB contains streams. We look for modules.
        // This is a simplified extraction of names. 
        // Actual source code decompression is complex, so we'll use AI for deep analysis if needed.
        cfb.FullPaths.forEach(path => {
          if (path.includes('VBA/')) {
            const name = path.split('/').pop() || '';
            if (name && !name.startsWith('_') && name !== 'dir') {
              macros.push({
                name,
                type: 'Module',
                content: 'Source code extraction in progress...'
              });
            }
          }
        });
      }
    } catch (err) {
      console.error("Macro extraction error:", err);
    }
    return macros;
  };

  const analyzeFile = async (file: ExcelFile): Promise<ExcelFile> => {
    try {
      const zip = new JSZip();
      const contents = await zip.loadAsync(file.data);
      const structure = Object.keys(contents.files);
      
      const vbaFiles = structure.filter(path => 
        path.toLowerCase().includes('vbaproject.bin')
      );

      const macros = await extractMacros(file.data);

      return {
        ...file,
        hasMacros: vbaFiles.length > 0 || macros.length > 0,
        macros,
        structure,
        status: 'ready'
      };
    } catch (err) {
      try {
        const workbook = XLSX.read(file.data, { type: 'array', bookVBA: true });
        const hasVBA = !!workbook.vbaraw;
        return {
          ...file,
          hasMacros: hasVBA,
          macros: hasVBA ? [{ name: 'Legacy VBA Project', type: 'Unknown' }] : [],
          structure: ['Binary Format (Legacy .xls)'],
          status: 'ready'
        };
      } catch (innerErr) {
        return {
          ...file,
          status: 'error',
          errorMessage: 'Unsupported or corrupted file format'
        };
      }
    }
  };

  const handleFileSelection = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(e.target.files || []);
    if (selectedFiles.length === 0) return;

    const uploadId = Math.random().toString(36).substring(7);

    const newFiles: ExcelFile[] = await Promise.all(
      selectedFiles.map(async (f) => {
        const file = f as File;
        const buffer = await file.arrayBuffer();
        return {
          id: Math.random().toString(36).substring(7),
          name: file.name,
          path: (file as any).webkitRelativePath || file.name,
          size: file.size,
          type: file.type,
          lastModified: file.lastModified,
          data: buffer,
          hasMacros: false,
          macros: [],
          structure: [],
          status: 'analyzing',
          uploadId
        };
      })
    );

    setFiles(prev => [...prev, ...newFiles]);

    for (const file of newFiles) {
      const analyzed = await analyzeFile(file);
      setFiles(prev => prev.map(f => f.id === file.id ? analyzed : f));
    }
  };

  const runSandboxSimulation = async (fileId: string, macroName?: string) => {
    const file = files.find(f => f.id === fileId);
    if (!file) return;

    setFiles(prev => prev.map(f => f.id === fileId ? { ...f, isSimulating: true } : f));

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || '' });
      
      // We send the file structure and macro names to the AI for simulation
      const prompt = `
        Analyze this Excel file structure and its VBA macros for security risks and simulate their behavior.
        File Name: ${file.name}
        File Path: ${file.path}
        Macros Found: ${file.macros.map(m => `${m.name} (${m.type})`).join(', ')}
        Internal Structure: ${file.structure.slice(0, 50).join(', ')}
        
        Task: 
        1. Simulate the execution of ${macroName ? `the macro "${macroName}"` : "all macros"} in a virtual sandbox.
        2. Identify potential errors, malicious behaviors (file system access, network calls), or suspicious logic.
        3. Provide a detailed report on what this macro likely does.
        4. If the macro name suggests common patterns (e.g., Auto_Open, Workbook_Open), explain the risks.
        
        Format the output in Markdown.
      `;

      const result = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: prompt,
      });
      
      const text = result.text || "No analysis generated.";

      setFiles(prev => prev.map(f => f.id === fileId ? { 
        ...f, 
        sandboxResult: text, 
        isSimulating: false 
      } : f));
    } catch (err) {
      setFiles(prev => prev.map(f => f.id === fileId ? { 
        ...f, 
        sandboxResult: "Error during simulation: " + (err as Error).message, 
        isSimulating: false 
      } : f));
    }
  };

  const removeMacros = async (fileId: string) => {
    const file = files.find(f => f.id === fileId);
    if (!file) return;

    setFiles(prev => prev.map(f => f.id === fileId ? { ...f, status: 'analyzing' } : f));

    try {
      const workbook = XLSX.read(file.data, { type: 'array' });
      const outData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const newName = file.name.replace(/\.(xlsm|xlsb|xls)$/i, '.xlsx');
      
      setFiles(prev => prev.map(f => f.id === fileId ? { 
        ...f, 
        name: newName,
        data: outData,
        hasMacros: false,
        macros: [],
        status: 'cleaned',
        sandboxResult: undefined
      } : f));
    } catch (err) {
      setFiles(prev => prev.map(f => f.id === fileId ? { 
        ...f, 
        status: 'error', 
        errorMessage: 'Failed to strip macros' 
      } : f));
    }
  };

  const downloadFile = (file: ExcelFile) => {
    const blob = new Blob([file.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = file.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const filteredFiles = files.filter(f => 
    f.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    f.path.toLowerCase().includes(searchTerm.toLowerCase())
  );

  // Tree structure building
  interface FileNode {
    name: string;
    path: string;
    isFolder: boolean;
    children: { [key: string]: FileNode };
    file?: ExcelFile;
  }

  const buildFileTree = (fileList: ExcelFile[]): FileNode => {
    const root: FileNode = { name: 'root', path: '', isFolder: true, children: {} };
    
    fileList.forEach(file => {
      const parts = file.path.split('/');
      let current = root;
      
      for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        const isLast = i === parts.length - 1;
        
        // Use uploadId for top-level nodes to prevent merging folders with same name from different uploads
        const nodeKey = i === 0 ? `${part}_${file.uploadId}` : part;
        
        if (!current.children[nodeKey]) {
          current.children[nodeKey] = {
            name: part,
            path: parts.slice(0, i + 1).join('/') + (i === 0 ? `_${file.uploadId}` : ''),
            isFolder: !isLast,
            children: {},
            file: isLast ? file : undefined
          };
        }
        current = current.children[nodeKey];
      }
    });
    
    return root;
  };

  const fileTree = buildFileTree(filteredFiles);

  const [expandedFolders, setExpandedFolders] = useState<Set<string>>(new Set(['root']));

  const toggleFolder = (path: string) => {
    setExpandedFolders(prev => {
      const next = new Set(prev);
      if (next.has(path)) next.delete(path);
      else next.add(path);
      return next;
    });
  };

  const getFolderFileIds = (node: FileNode): string[] => {
    const ids: string[] = [];
    const collect = (n: FileNode) => {
      if (n.file) ids.push(n.file.id);
      Object.values(n.children).forEach(collect);
    };
    collect(node);
    return ids;
  };

  const toggleFolderCheck = (node: FileNode, checked: boolean) => {
    const ids = getFolderFileIds(node);
    setCheckedFileIds(prev => {
      if (checked) {
        const next = new Set(prev);
        ids.forEach(id => next.add(id));
        return Array.from(next);
      } else {
        return prev.filter(id => !ids.includes(id));
      }
    });
  };

  const isFolderChecked = (node: FileNode): boolean => {
    const ids = getFolderFileIds(node);
    if (ids.length === 0) return false;
    return ids.every(id => checkedFileIds.includes(id));
  };

  const isFolderIndeterminate = (node: FileNode): boolean => {
    const ids = getFolderFileIds(node);
    if (ids.length === 0) return false;
    const checkedCount = ids.filter(id => checkedFileIds.includes(id)).length;
    return checkedCount > 0 && checkedCount < ids.length;
  };

  const renderTree = (node: FileNode, level: number = 0) => {
    const sortedChildren = Object.values(node.children).sort((a, b) => {
      if (a.isFolder && !b.isFolder) return -1;
      if (!a.isFolder && b.isFolder) return 1;
      return a.name.localeCompare(b.name);
    });

    return sortedChildren.map((child) => {
      const isExpanded = expandedFolders.has(child.path);
      
      if (child.isFolder) {
        const checked = isFolderChecked(child);
        const indeterminate = isFolderIndeterminate(child);

        return (
          <div key={child.path}>
            <div 
              className="flex items-center gap-2 p-2 hover:bg-[#141414]/5 cursor-pointer group"
              style={{ paddingLeft: `${level * 12 + 12}px` }}
              onClick={() => toggleFolder(child.path)}
            >
              <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
                <input 
                  type="checkbox"
                  checked={checked}
                  ref={el => {
                    if (el) el.indeterminate = indeterminate;
                  }}
                  onChange={(e) => toggleFolderCheck(child, e.target.checked)}
                  className="accent-[#141414] w-3 h-3"
                />
              </div>
              <ChevronRight 
                size={14} 
                className={cn("transition-transform opacity-30", isExpanded && "rotate-90")} 
              />
              <FolderOpen size={14} className="opacity-50" />
              <span className="font-mono text-[10px] uppercase truncate flex-1">{child.name}</span>
            </div>
            {isExpanded && renderTree(child, level + 1)}
          </div>
        );
      }

      const file = child.file!;
      return (
        <div 
          key={file.id}
          onClick={() => setSelectedFileId(file.id)}
          className={cn(
            "p-3 cursor-pointer transition-all group relative flex items-start gap-3",
            selectedFileId === file.id ? "bg-[#141414] text-[#E4E3E0]" : "hover:bg-[#141414]/5"
          )}
          style={{ paddingLeft: `${level * 12 + 24}px` }}
        >
          <input 
            type="checkbox"
            checked={checkedFileIds.includes(file.id)}
            onClick={(e) => e.stopPropagation()}
            onChange={() => toggleFileCheck(file.id)}
            className={cn(
              "mt-1 accent-[#141414]",
              selectedFileId === file.id && "accent-white"
            )}
          />
          <div className="flex-1 min-w-0">
            <div className="flex items-start justify-between gap-2">
              <div className="flex-1 min-w-0">
                <h3 className="font-mono text-[10px] truncate uppercase tracking-tighter">{file.name}</h3>
                <div className="flex items-center gap-2 mt-1">
                  <span className={cn(
                    "text-[8px] font-mono px-1 py-0 border",
                    selectedFileId === file.id ? "border-white/30" : "border-[#141414]/20"
                  )}>
                    {file.status.toUpperCase()}
                  </span>
                  {file.macros.length > 0 && (
                    <span className="text-[8px] font-mono bg-orange-500 text-white px-1 py-0">
                      {file.macros.length} M
                    </span>
                  )}
                </div>
              </div>
              {file.hasMacros ? (
                <ShieldAlert className="text-orange-500 shrink-0" size={14} />
              ) : file.status === 'ready' || file.status === 'cleaned' ? (
                <ShieldCheck className="text-emerald-500 shrink-0" size={14} />
              ) : null}
            </div>
          </div>
          <button 
            onClick={(e) => { e.stopPropagation(); setFiles(prev => prev.filter(f => f.id !== file.id)); }}
            className="absolute right-2 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 p-1 hover:text-red-500 transition-opacity"
          >
            <X size={12} />
          </button>
        </div>
      );
    });
  };

  return (
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] font-sans selection:bg-[#141414] selection:text-[#E4E3E0]">
      {/* Header */}
      <header className="border-bottom border-[#141414] px-6 py-4 flex items-center justify-between bg-white/50 backdrop-blur-sm sticky top-0 z-10">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-[#141414] flex items-center justify-center rounded-sm">
            <ShieldAlert className="text-[#E4E3E0] w-6 h-6" />
          </div>
          <div>
            <h1 className="font-serif italic text-xl leading-none">Excel Macro Sentinel</h1>
            <p className="text-[10px] uppercase tracking-widest opacity-50 mt-1 font-mono">Security & Recursive Analysis Utility</p>
          </div>
        </div>
        <div className="flex gap-2">
          <button 
            onClick={scanKangatangVirus}
            className="px-4 py-2 border border-red-600 text-red-600 text-xs font-mono uppercase tracking-tighter hover:bg-red-600 hover:text-white transition-colors flex items-center gap-2"
          >
            <ShieldAlert size={14} />
            Kangatang Scan
          </button>
          <button 
            onClick={rescanAllFiles}
            disabled={files.length === 0}
            className="px-4 py-2 border border-[#141414] text-xs font-mono uppercase tracking-tighter hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors flex items-center gap-2 disabled:opacity-30"
          >
            <Loader2 size={14} className={files.some(f => f.status === 'analyzing') ? 'animate-spin' : ''} />
            Rescan All
          </button>
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="px-4 py-2 border border-[#141414] text-xs font-mono uppercase tracking-tighter hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors flex items-center gap-2"
          >
            <FileSpreadsheet size={14} />
            Add Files
          </button>
          <button 
            onClick={() => folderInputRef.current?.click()}
            className="px-4 py-2 border border-[#141414] text-xs font-mono uppercase tracking-tighter hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors flex items-center gap-2"
          >
            <FolderOpen size={14} />
            Recursive Scan
          </button>
          <input 
            type="file" 
            ref={fileInputRef} 
            className="hidden" 
            multiple 
            accept=".xls,.xlsx,.xlsm,.xlsb" 
            onChange={handleFileSelection} 
          />
          <input 
            type="file" 
            ref={folderInputRef} 
            className="hidden" 
            // @ts-ignore
            webkitdirectory="" 
            directory="" 
            onChange={handleFileSelection} 
          />
        </div>
      </header>

      <main className="flex h-[calc(100vh-73px)]">
        {/* Sidebar: File List */}
        <aside className="w-96 border-right border-[#141414] flex flex-col bg-white/30">
          <div className="p-4 border-bottom border-[#141414] bg-[#141414]/5 space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <input 
                  type="checkbox" 
                  checked={files.length > 0 && checkedFileIds.length === files.length}
                  onChange={(e) => toggleAllFilesCheck(e.target.checked)}
                  className="accent-[#141414]"
                />
                <span className="text-[10px] font-mono uppercase opacity-50">Select All</span>
              </div>
              {checkedFileIds.length > 0 && (
                <div className="flex gap-1">
                  <button 
                    onClick={removeSelectedFilesFromList}
                    className="text-[10px] font-mono uppercase bg-red-600 text-white px-2 py-1 flex items-center gap-2 hover:bg-red-700 transition-colors"
                  >
                    <Trash2 size={12} />
                    Remove ({checkedFileIds.length})
                  </button>
                  <button 
                    onClick={downloadSelectedFiles}
                    className="text-[10px] font-mono uppercase bg-[#141414] text-[#E4E3E0] px-2 py-1 flex items-center gap-2 hover:bg-[#141414]/80 transition-colors"
                  >
                    <Download size={12} />
                    Download ({checkedFileIds.length})
                  </button>
                </div>
              )}
            </div>
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 opacity-30" size={14} />
              <input 
                type="text"
                placeholder="SEARCH FILES OR PATHS..."
                className="w-full bg-transparent border border-[#141414]/20 pl-9 pr-4 py-2 text-[10px] font-mono uppercase tracking-widest focus:outline-none focus:border-[#141414]"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>
          
          <div className="flex-1 overflow-y-auto custom-scrollbar">
            {filteredFiles.length === 0 ? (
              <div className="p-12 text-center opacity-30 flex flex-col items-center gap-4">
                <FileSearch size={48} strokeWidth={1} />
                <p className="font-serif italic text-sm">No matching files found</p>
              </div>
            ) : (
              <div className="py-2">
                {renderTree(fileTree)}
              </div>
            )}
          </div>
        </aside>

        {/* Main Content: Analysis */}
        <section className="flex-1 overflow-y-auto p-8 bg-white/20">
          <AnimatePresence mode="wait">
            {selectedFile ? (
              <motion.div 
                key={selectedFile.id}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="max-w-5xl mx-auto"
              >
                {/* File Header */}
                <div className="flex items-end justify-between border-bottom border-[#141414] pb-6 mb-8">
                  <div className="flex-1">
                    <div className="flex items-center gap-2 mb-2">
                      <span className="font-mono text-[10px] bg-[#141414] text-[#E4E3E0] px-2 py-0.5 uppercase">Recursive Analysis</span>
                      <span className="font-mono text-[10px] opacity-50 truncate max-w-md">{selectedFile.path}</span>
                    </div>
                    <h2 className="font-serif italic text-4xl">{selectedFile.name}</h2>
                  </div>
                  <div className="flex gap-3">
                    {selectedFile.hasMacros && (
                      <button 
                        onClick={() => removeMacros(selectedFile.id)}
                        className="px-6 py-3 bg-orange-500 text-white font-mono text-xs uppercase tracking-widest hover:bg-orange-600 transition-colors flex items-center gap-2"
                      >
                        <Trash2 size={16} />
                        Strip Macros
                      </button>
                    )}
                    <button 
                      onClick={() => downloadFile(selectedFile)}
                      className="px-6 py-3 border border-[#141414] font-mono text-xs uppercase tracking-widest hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors flex items-center gap-2"
                    >
                      <Download size={16} />
                      Download
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-12 gap-8">
                  {/* Left Column: Stats & Macros */}
                  <div className="col-span-4 space-y-6">
                    <div className="p-6 bg-white/50 border border-[#141414] space-y-4">
                      <div className="flex items-center justify-between">
                        <span className="font-serif italic text-[10px] uppercase opacity-50">Security Profile</span>
                        {selectedFile.hasMacros ? <ShieldAlert className="text-orange-500" size={16} /> : <ShieldCheck className="text-emerald-500" size={16} />}
                      </div>
                      <div className="space-y-2">
                        <div className="flex justify-between font-mono text-[10px]">
                          <span className="opacity-50">VBA DETECTED:</span>
                          <span className={selectedFile.hasMacros ? "text-orange-600 font-bold" : "text-emerald-600"}>{selectedFile.hasMacros ? "YES" : "NO"}</span>
                        </div>
                        <div className="flex justify-between font-mono text-[10px]">
                          <span className="opacity-50">MACRO COUNT:</span>
                          <span>{selectedFile.macros.length}</span>
                        </div>
                        <div className="flex justify-between font-mono text-[10px]">
                          <span className="opacity-50">FILE SIZE:</span>
                          <span>{(selectedFile.size / 1024).toFixed(2)} KB</span>
                        </div>
                      </div>
                    </div>

                    <div className="bg-white/50 border border-[#141414]">
                      <div className="p-4 border-bottom border-[#141414] bg-[#141414]/5 flex items-center justify-between">
                        <h3 className="font-serif italic text-sm flex items-center gap-2">
                          <Code2 size={14} />
                          Macro Components
                        </h3>
                        {selectedFile.macros.length > 0 && (
                          <div className="flex items-center gap-4">
                            <button 
                              onClick={() => {
                                const allSelected = selectedFile.macros.every(m => m.selected);
                                toggleAllMacros(selectedFile.id, !allSelected);
                              }}
                              className="text-[9px] font-mono uppercase underline hover:no-underline"
                            >
                              {selectedFile.macros.every(m => m.selected) ? 'Deselect All' : 'Select All'}
                            </button>
                            <button 
                              onClick={() => deleteSelectedMacros(selectedFile.id)}
                              disabled={!selectedFile.macros.some(m => m.selected)}
                              className="text-[9px] font-mono uppercase text-red-600 font-bold disabled:opacity-30 flex items-center gap-1"
                            >
                              <Trash2 size={10} />
                              Delete Selected
                            </button>
                          </div>
                        )}
                      </div>
                      <div className="divide-y divide-[#141414]">
                        {selectedFile.macros.length === 0 ? (
                          <div className="p-8 text-center opacity-30 font-mono text-[10px]">NO MACROS FOUND</div>
                        ) : (
                          selectedFile.macros.map((macro, i) => (
                            <div key={i} className={cn(
                              "p-4 hover:bg-[#141414]/5 group transition-colors flex items-start gap-3",
                              macro.selected && "bg-[#141414]/5"
                            )}>
                              <input 
                                type="checkbox" 
                                checked={!!macro.selected}
                                onChange={() => toggleMacroSelection(selectedFile.id, macro.name)}
                                className="mt-1 accent-[#141414]"
                              />
                              <div className="flex-1">
                                <div className="flex items-center justify-between mb-2">
                                  <span className="font-mono text-[11px] font-bold uppercase">{macro.name}</span>
                                  <span className="text-[9px] font-mono opacity-50 px-1.5 py-0.5 border border-[#141414]/20">{macro.type}</span>
                                </div>
                                <button 
                                  onClick={() => runSandboxSimulation(selectedFile.id, macro.name)}
                                  disabled={selectedFile.isSimulating}
                                  className="w-full py-2 border border-[#141414] font-mono text-[9px] uppercase tracking-widest hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors flex items-center justify-center gap-2 disabled:opacity-50"
                                >
                                  {selectedFile.isSimulating ? <Loader2 size={12} className="animate-spin" /> : <Play size={12} />}
                                  Run Virtual Sandbox
                                </button>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Right Column: AI Sandbox & Reports */}
                  <div className="col-span-8 space-y-6">
                    <div className="bg-[#141414] text-[#E4E3E0] border border-[#141414] min-h-[500px] flex flex-col">
                      <div className="p-4 border-bottom border-white/10 flex items-center justify-between bg-white/5">
                        <div className="flex items-center gap-3">
                          <Terminal size={16} className="text-emerald-400" />
                          <h3 className="font-mono text-xs uppercase tracking-widest">AI Virtual Sandbox Environment</h3>
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                          <span className="font-mono text-[9px] opacity-50 uppercase">Ready for Simulation</span>
                        </div>
                      </div>
                      
                      <div className="flex-1 p-6 font-mono text-[11px] overflow-y-auto custom-scrollbar">
                        {selectedFile.isSimulating ? (
                          <div className="h-full flex flex-col items-center justify-center gap-4 opacity-50">
                            <Cpu size={48} className="animate-pulse" />
                            <p className="animate-bounce">INITIALIZING VIRTUAL ENVIRONMENT & ANALYZING BYTECODE...</p>
                          </div>
                        ) : selectedFile.sandboxResult ? (
                          <div className="prose prose-invert prose-xs max-w-none">
                            <Markdown>{selectedFile.sandboxResult}</Markdown>
                          </div>
                        ) : (
                          <div className="h-full flex flex-col items-center justify-center gap-4 opacity-20 text-center">
                            <Terminal size={64} strokeWidth={1} />
                            <div>
                              <p className="text-sm italic font-serif">Waiting for simulation trigger...</p>
                              <p className="mt-2 text-[9px] uppercase tracking-widest">Select a macro to simulate behavior in a safe environment</p>
                            </div>
                          </div>
                        )}
                      </div>

                      <div className="p-3 bg-white/5 border-top border-white/10 flex items-center justify-between font-mono text-[9px] opacity-40">
                        <span>SANDBOX_ID: {selectedFile.id.toUpperCase()}</span>
                        <span>ENGINE: GEMINI-3.1-PRO</span>
                      </div>
                    </div>

                    {/* Structure Explorer */}
                    <div className="bg-white/50 border border-[#141414]">
                      <div className="p-4 border-bottom border-[#141414] bg-[#141414]/5 flex items-center justify-between">
                        <h3 className="font-serif italic text-sm flex items-center gap-2">
                          <FileText size={14} />
                          Internal File Nodes
                        </h3>
                        <span className="font-mono text-[10px] opacity-50">{selectedFile.structure.length} Nodes</span>
                      </div>
                      <div className="max-h-[200px] overflow-y-auto p-4 font-mono text-[10px] grid grid-cols-2 gap-x-8 gap-y-1">
                        {selectedFile.structure.map((node, i) => (
                          <div key={i} className="flex items-center gap-2 opacity-60 hover:opacity-100 transition-opacity truncate">
                            <ChevronRight size={10} className="shrink-0" />
                            <span className="truncate">{node}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              </motion.div>
            ) : (
              <div className="h-full flex flex-col items-center justify-center text-center opacity-20">
                <ShieldCheck size={120} strokeWidth={0.5} />
                <h2 className="font-serif italic text-2xl mt-6">Select a file from the recursive scan</h2>
                <p className="font-mono text-xs uppercase tracking-widest mt-2">Ready for deep inspection</p>
              </div>
            )}
          </AnimatePresence>
        </section>
      </main>

      {/* Footer / Status Bar */}
      <footer className="border-top border-[#141414] px-6 py-2 bg-white/50 flex items-center justify-between font-mono text-[10px] uppercase tracking-widest opacity-50">
        <div className="flex gap-6">
          <span>Status: Secure</span>
          <span>Recursive Mode: Active</span>
          <span>AI Engine: Gemini 3.1 Pro</span>
        </div>
        <div>
          © 2026 Sentinel Security Labs • Multi-Layer Analysis
        </div>
      </footer>

      {/* Scan Result Modal */}
      <AnimatePresence>
        {scanResultModal?.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#141414]/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 max-w-md w-full shadow-2xl"
            >
              <div className="flex items-center gap-4 mb-6">
                <div className={cn(
                  "w-12 h-12 flex items-center justify-center rounded-sm",
                  scanResultModal.found ? "bg-red-600 text-white" : "bg-emerald-600 text-white"
                )}>
                  {scanResultModal.found ? <ShieldAlert size={24} /> : <ShieldCheck size={24} />}
                </div>
                <div>
                  <h3 className="font-serif italic text-2xl leading-none">Scan Complete</h3>
                  <p className="text-[10px] uppercase tracking-widest opacity-50 mt-1 font-mono">Kangatang Engine v1.0</p>
                </div>
              </div>
              
              <p className="font-mono text-xs leading-relaxed mb-8 opacity-80">
                {scanResultModal.message}
              </p>
              
              <button 
                onClick={() => setScanResultModal(null)}
                className="w-full py-3 bg-[#141414] text-[#E4E3E0] font-mono text-xs uppercase tracking-widest hover:bg-[#141414]/80 transition-colors"
              >
                Acknowledge
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: rgba(255, 255, 255, 0.05);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: rgba(255, 255, 255, 0.2);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: rgba(255, 255, 255, 0.3);
        }
      `}} />
    </div>
  );
}
