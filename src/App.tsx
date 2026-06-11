import { useState, useEffect, useMemo, useRef } from 'react';
import {
  Layers,
  Plus,
  Trash2,
  Settings,
  Upload,
  Play,
  Image as ImageIcon,
  Printer,
  AlertTriangle,
  RefreshCw,
  Info,
  ChevronDown,
  ChevronUp,
  FileSpreadsheet,
  Copy,
  Check
} from 'lucide-react';
import { solveNesting } from './utils/solver';
import type { CutItem, SolveResult, SolverSettings, NestedBar } from './utils/solver';
import { parseTableText } from './utils/parser';
import type { ParsedAssembly } from './utils/parser';
import { analyzeBOMImage } from './utils/gemini';
import * as XLSX from 'xlsx';

// Interface representing a construction / assembly
interface Construction {
  id: string;
  name: string;
  position: string;
  qty: number;
  items: {
    id: string;
    position: string;
    profile: string; // e.g. "50 x 50 x 3.2"
    length: number; // mm
    qty: number;
    isSheetMetal?: boolean;
    width?: number; // only for sheet metal
  }[];
}

interface GroupedBar {
  cutsKey: string;
  sampleBar: NestedBar;
  count: number;
  barIndices: number[];
  originalBars: NestedBar[];
}

// Initial preloaded data matching the user's screenshot exactly
const INITIAL_CONSTRUCTIONS: Construction[] = [
  {
    id: 'c1',
    name: 'მოაჯირი - ორივე მხარეს',
    position: '1',
    qty: 1,
    items: [
      { id: 'c1-i1', position: '1.1', profile: '50 x 50 x 3.2', length: 1090, qty: 2 },
      { id: 'c1-i2', position: '1.2', profile: '50 x 30 x 2.6', length: 2000, qty: 1 },
      { id: 'c1-i3', position: '1.3', profile: '20 x 20 x 2.0', length: 960, qty: 15 },
      { id: 'c1-i4', position: '1.4', profile: '50 x 30 x 2.6', length: 1900, qty: 1 }
    ]
  },
  {
    id: 'c2',
    name: 'მოაჯირი - ცალმხრივი გადაბმით',
    position: '2',
    qty: 31,
    items: [
      { id: 'c2-i1', position: '2.1', profile: '50 x 50 x 3.2', length: 1090, qty: 1 },
      { id: 'c2-i2', position: '2.2', profile: '50 x 30 x 2.6', length: 2000, qty: 1 },
      { id: 'c2-i3', position: '2.3', profile: '20 x 20 x 2.0', length: 960, qty: 15 },
      { id: 'c2-i4', position: '2.4', profile: '50 x 30 x 2.6', length: 1950, qty: 1 }
    ]
  }
];

// Aesthetic profile colors for segments
const PROFILE_COLORS: Record<string, string> = {
  '50 x 50 x 3.2': '#6366f1', // Indigo
  '50 x 30 x 2.6': '#10b981', // Emerald
  '20 x 20 x 2.0': '#8b5cf6', // Violet
  'default1': '#3b82f6',
  'default2': '#ec4899',
  'default3': '#f59e0b',
  'default4': '#14b8a6',
  'default5': '#f43f5e'
};

export default function App() {
  const [constructions, setConstructions] = useState<Construction[]>(() => {
    try {
      const saved = localStorage.getItem('material_nesting_constructions');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          // Sanitize the loaded data to ensure no NaNs or undefined items exist
          return parsed.map((c: any, cIndex: number) => ({
            id: c.id || `c-saved-${Date.now()}-${cIndex}`,
            name: c.name || `კონსტრუქცია ${cIndex + 1}`,
            position: c.position || String(cIndex + 1),
            qty: Math.max(1, parseInt(c.qty, 10) || 1),
            items: Array.isArray(c.items)
              ? c.items.map((item: any, iIndex: number) => ({
                  id: item.id || `i-saved-${Date.now()}-${cIndex}-${iIndex}`,
                  position: item.position || `${c.position || cIndex + 1}.${iIndex + 1}`,
                  profile: item.profile || '50 x 50 x 3.2',
                  length: Math.max(0, parseFloat(item.length) || 0),
                  width: item.width !== undefined ? Math.max(0, parseFloat(item.width) || 0) : undefined,
                  qty: Math.max(1, parseInt(item.qty, 10) || 1),
                  isSheetMetal: !!item.isSheetMetal
                }))
              : []
          }));
        }
      }
    } catch (e) {
      console.error('Error loading constructions from cache:', e);
    }
    return [];
  });

  const [settings, setSettings] = useState<SolverSettings>(() => {
    try {
      const saved = localStorage.getItem('nesting_settings');
      if (saved) {
        const parsed = JSON.parse(saved);
        return {
          stockLength: Math.max(1, parseInt(parsed.stockLength, 10) || 6000),
          kerf: Math.max(0, parseInt(parsed.kerf, 10) || 3),
          trim: Math.max(0, parseInt(parsed.trim, 10) || 10),
          customProfileSettings: parsed.customProfileSettings || {}
        };
      }
    } catch (e) {
      console.error('Error loading settings from cache:', e);
    }
    return {
      stockLength: 6000,
      kerf: 3,
      trim: 10,
      customProfileSettings: {}
    };
  });

  const [apiKey, setApiKey] = useState<string>(() => {
    return localStorage.getItem('gemini_api_key') || '';
  });

  const [activeTab, setActiveTab] = useState<'editor' | 'paste' | 'ai' | 'results'>('editor');
  
  // Text area for Excel paste
  const [rawPasteText, setRawPasteText] = useState('');
  
  // AI upload state
  const [dragActive, setDragActive] = useState(false);
  const [imagePreview, setImagePreview] = useState<string | null>(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);
  const [aiLogs, setAiLogs] = useState<string[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Sheet Metal waste settings
  const [sheetMetalWasteFactor, setSheetMetalWasteFactor] = useState(1.15); // +15% waste

  // Copy cut list feedback
  const [copiedStates, setCopiedStates] = useState<Record<string, boolean>>({});

  // Collapsed assemblies in editor
  const [collapsedAss, setCollapsedAss] = useState<Record<string, boolean>>({});

  // Printing state
  const [isPrinting, setIsPrinting] = useState(false);

  // Trigger print and reset
  useEffect(() => {
    if (isPrinting) {
      const timer = setTimeout(() => {
        window.print();
      }, 300);
      return () => clearTimeout(timer);
    }
  }, [isPrinting]);

  useEffect(() => {
    const handleAfterPrint = () => {
      setIsPrinting(false);
    };
    window.addEventListener('afterprint', handleAfterPrint);
    return () => {
      window.removeEventListener('afterprint', handleAfterPrint);
    };
  }, []);

  // Save data to localStorage
  useEffect(() => {
    localStorage.setItem('material_nesting_constructions', JSON.stringify(constructions));
  }, [constructions]);

  useEffect(() => {
    localStorage.setItem('nesting_settings', JSON.stringify(settings));
  }, [settings]);

  useEffect(() => {
    localStorage.setItem('gemini_api_key', apiKey);
  }, [apiKey]);

  // Helpers to get profile color
  const getProfileColor = (profile: string, index: number) => {
    if (PROFILE_COLORS[profile]) return PROFILE_COLORS[profile];
    const keys = Object.keys(PROFILE_COLORS).filter(k => !k.startsWith('default'));
    const colorIndex = (index + keys.length) % 5 + 1;
    return PROFILE_COLORS[`default${colorIndex}`];
  };

  // Compile all cut items and sheet metals from all constructions multiplied by parent qty
  const compiledData = useMemo(() => {
    const profileCuts: Record<string, CutItem[]> = {};
    const sheetMetals: Record<
      string,
      { id: string; constructionName: string; position: string; length: number; width: number; qty: number }[]
    > = {};

    const activeProfiles = new Set<string>();

    constructions.forEach((c) => {
      c.items.forEach((item) => {
        const totalQty = item.qty * c.qty;
        if (totalQty <= 0) return;

        if (item.isSheetMetal) {
          const key = item.profile; // e.g. "2.0mm Plate"
          if (!sheetMetals[key]) sheetMetals[key] = [];
          sheetMetals[key].push({
            id: item.id,
            constructionName: c.name,
            position: item.position,
            length: item.length,
            width: item.width || 1000,
            qty: totalQty
          });
        } else {
          const key = item.profile; // e.g. "50 x 50 x 3.2"
          activeProfiles.add(key);
          if (!profileCuts[key]) profileCuts[key] = [];
          profileCuts[key].push({
            id: item.id,
            constructionId: c.id,
            constructionName: c.name,
            profile: item.profile,
            length: item.length,
            qty: totalQty,
            position: item.position
          });
        }
      });
    });

    return {
      profileCuts,
      sheetMetals,
      profilesList: Array.from(activeProfiles)
    };
  }, [constructions]);

  // Nesting solver results
  const solveResults = useMemo(() => {
    const results: Record<string, SolveResult> = {};
    Object.keys(compiledData.profileCuts).forEach((profile) => {
      results[profile] = solveNesting(profile, compiledData.profileCuts[profile], settings);
    });
    return results;
  }, [compiledData, settings]);

  // Total bars summary
  const totalsSummary = useMemo(() => {
    let totalBars = 0;
    let totalCuts = 0;
    let totalScrapLength = 0;
    let averageYield = 0;
    const profileCount = Object.keys(solveResults).length;

    Object.values(solveResults).forEach((res) => {
      totalBars += res.totalBars;
      totalCuts += res.totalCuts;
      totalScrapLength += res.totalWaste;
      averageYield += res.yieldPercent;
    });

    return {
      totalBars,
      totalCuts,
      totalScrapLength,
      averageYield: profileCount > 0 ? averageYield / profileCount : 100
    };
  }, [solveResults]);

  // Construction CRUD operations
  const addConstruction = () => {
    const newId = `c-${Date.now()}`;
    const newPos = String(constructions.length + 1);
    const newC: Construction = {
      id: newId,
      name: `კონსტრუქცია ${newPos}`,
      position: newPos,
      qty: 1,
      items: [
        { id: `${newId}-i1`, position: `${newPos}.1`, profile: '50 x 50 x 3.2', length: 1000, qty: 1 }
      ]
    };
    setConstructions([...constructions, newC]);
  };

  const removeConstruction = (id: string) => {
    setConstructions(constructions.filter((c) => c.id !== id));
  };

  const updateConstructionHeader = (id: string, field: 'name' | 'position' | 'qty', value: any) => {
    setConstructions(
      constructions.map((c) => {
        if (c.id !== id) return c;
        if (field === 'qty') {
          return { ...c, qty: Math.max(0, parseInt(value, 10) || 0) };
        }
        return { ...c, [field]: value };
      })
    );
  };

  // Add Item to construction
  const addItemToConstruction = (constructionId: string, isSheet: boolean = false) => {
    setConstructions(
      constructions.map((c) => {
        if (c.id !== constructionId) return c;
        const newId = `i-${Date.now()}`;
        const newPos = `${c.position}.${c.items.length + 1}`;
        const newItem = isSheet
          ? {
              id: newId,
              position: newPos,
              profile: 'ფურცელი 3მმ',
              length: 1000,
              width: 1000,
              qty: 1,
              isSheetMetal: true
            }
          : {
              id: newId,
              position: newPos,
              profile: c.items[c.items.length - 1]?.profile || '50 x 50 x 3.2',
              length: 1000,
              qty: 1
            };
        return { ...c, items: [...c.items, newItem] };
      })
    );
  };

  const removeItemFromConstruction = (constructionId: string, itemId: string) => {
    setConstructions(
      constructions.map((c) => {
        if (c.id !== constructionId) return c;
        return { ...c, items: c.items.filter((item) => item.id !== itemId) };
      })
    );
  };

  const updateItemInConstruction = (
    constructionId: string,
    itemId: string,
    field: string,
    value: any
  ) => {
    setConstructions(
      constructions.map((c) => {
        if (c.id !== constructionId) return c;
        return {
          ...c,
          items: c.items.map((item) => {
            if (item.id !== itemId) return item;
            if (field === 'length' || field === 'qty' || field === 'width') {
              return { ...item, [field]: Math.max(0, parseFloat(value) || 0) };
            }
            return { ...item, [field]: value };
          })
        };
      })
    );
  };

  // Import Parsed assemblies
  const importParsedAssemblies = (assemblies: ParsedAssembly[]) => {
    const formatted = assemblies.map((ass, aIndex) => {
      const parentId = `c-import-${Date.now()}-${aIndex}`;
      return {
        id: parentId,
        name: ass.name,
        position: ass.position || String(aIndex + 1),
        qty: ass.qty || 1,
        items: ass.items.map((item, iIndex) => {
          // Detect sheet metal from description or parser flag
          const isSheet = item.isSheetMetal !== undefined
            ? item.isSheetMetal
            : (item.description.toLowerCase().includes('plate') ||
               item.description.toLowerCase().includes('sheet') ||
               item.description.toLowerCase().includes('ფურცელი'));
          
          return {
            id: `${parentId}-i-${iIndex}`,
            position: item.position || `${ass.position || aIndex + 1}.${iIndex + 1}`,
            profile: item.description,
            length: item.length || 1000,
            qty: item.qty || 1,
            isSheetMetal: isSheet,
            width: item.width !== undefined ? item.width : (isSheet ? 1000 : undefined)
          };
        })
      };
    });

    setConstructions(formatted);
    setActiveTab('results');
  };

  // Text paste import handler
  const handlePasteImport = () => {
    if (!rawPasteText.trim()) return;
    try {
      const parsed = parseTableText(rawPasteText);
      if (parsed.length === 0) {
        alert('სვეტების ამოკითხვა ვერ მოხერხდა. დარწმუნდით, რომ მონაცემები შეიცავს პროფილის ზომას და სიგრძეს მაინც.');
        return;
      }
      importParsedAssemblies(parsed);
      setRawPasteText('');
    } catch (e) {
      alert('ბუფერიდან ჩაწერილი ტექსტის გაანალიზების შეცდომა. დარწმუნდით, რომ იგი კოპირებულია Excel-ის ან Solidworks-ის ცხრილის ფორმატიდან.');
    }
  };

  // Excel File upload handler
  const handleExcelUpload = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Convert sheet to 2D array
        const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 });
        
        // Convert rows to TSV string
        const tsvText = rows
          .map((row) =>
            row
              .map((cell) => {
                if (cell === null || cell === undefined) return '';
                return String(cell).replace(/\r?\n/g, ' ').replace(/\t/g, ' ');
              })
              .join('\t')
          )
          .join('\n');
          
        const parsed = parseTableText(tsvText);
        if (parsed.length === 0) {
          alert('სვეტების ამოკითხვა ელექტრონული ცხრილიდან ვერ მოხერხდა. შეამოწმეთ სვეტების სათაურები, როგორიცაა: პოზიცია, აღწერა, სიგრძე, რაოდენობა.');
          return;
        }
        importParsedAssemblies(parsed);
      } catch (err) {
        alert('Excel ფაილის გაანალიზება ვერ მოხერხდა. დარწმუნდით, რომ ფაილი არის ვალიდური .xlsx, .xls ან .csv ფორმატის.');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // File drag & drop handlers
  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  };

  const processFile = (file: File) => {
    if (!file.type.startsWith('image/')) {
      alert('გთხოვთ ატვირთოთ სურათის ფაილი (PNG, JPG).');
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      if (e.target?.result) {
        setImagePreview(e.target.result as string);
        setAiError(null);
      }
    };
    reader.readAsDataURL(file);
  };

  // Run AI BOM extraction
  const runAIExtraction = async () => {
    if (!imagePreview) return;
    if (!apiKey) {
      setAiError('გთხოვთ შეიყვანოთ Gemini API გასაღები მარცხენა პარამეტრების პანელში.');
      return;
    }

    setAiLoading(true);
    setAiError(null);
    setAiLogs(['Gemini API მოთხოვნის ინიციალიზაცია...', 'სურათის base64-ის ატვირთვა...', 'ცხრილის სტრუქტურისა და მონაცემების დამუშავება...']);

    try {
      const timer = setInterval(() => {
        setAiLogs((prev) => [
          ...prev,
          ['სათაურების გაანალიზება...', 'რაოდენობების წაკითხვა...', 'სიგრძეების გამოთვლა...', 'JSON სტრუქტურის შექმნა...'][Math.floor(Math.random() * 4)]
        ]);
      }, 1500);

      const data = await analyzeBOMImage(imagePreview, apiKey);
      clearInterval(timer);

      setAiLogs((prev) => [...prev, 'BOM წარმატებით ამოიკითხა!', 'კონსტრუქციებისა და პროფილების დაკავშირება...']);
      
      if (data.length === 0) {
        throw new Error('მონაცემები ვერ ამოიკითხა. დარწმუნდით, რომ სკრინშოტში მოცემული ცხრილი მკაფიოდ იკითხება.');
      }
      
      importParsedAssemblies(data);
    } catch (err: any) {
      setAiError(err.message || 'ამოკითხვისას დაფიქსირდა მოულოდნელი შეცდომა.');
    } finally {
      setAiLoading(false);
    }
  };

  const clearImage = () => {
    setImagePreview(null);
    setAiError(null);
    setAiLogs([]);
  };

  // Copy bar cutting sheet to clipboard
  const copyCutInstructions = (profile: string, res: SolveResult) => {
    let txt = `=== გადანაჭრების ინსტრუქცია პროფილისთვის: ${profile} ===\n`;
    txt += `სულ საწყისი ღერო (6მ): ${res.totalBars}\n`;
    txt += `კიდის ჩამონაჭერი: ${res.trim}მმ | ხერხის სისქე: ${res.kerf}მმ\n\n`;
    
    const grouped = getGroupedBars(res.bars);
    grouped.forEach((gBar) => {
      const bar = gBar.sampleBar;
      const barTitle = gBar.count > 1
        ? `ღეროები #${gBar.barIndices[0]} - #${gBar.barIndices[gBar.barIndices.length - 1]} (${gBar.count} ცალი - იდენტური გადანაჭრებით)`
        : `ღერო #${gBar.barIndices[0]}`;
        
      txt += `${barTitle}:\n`;
      bar.cuts.forEach((cut, idx) => {
        txt += `  - გადანაჭერი ${idx + 1}: ${cut.length}მმ [კოორდინატი: ${cut.startOffset}მმ - ${cut.endOffset}მმ]\n`;
      });
      
      const groupedCuts = getGroupedCutsInfo(gBar);
      txt += `  ჯამური დეტალები ამ ჯგუფისთვის:\n`;
      groupedCuts.forEach((item) => {
        const positionsText = item.positionsList.map(p => `პოზ ${p.pos} (${p.qty}ც)`).join(', ');
        txt += `    * ${item.length}მმ x ${item.totalQty}ცალი (${positionsText})\n`;
      });
      
      txt += `  - ნარჩენი: ${bar.waste}მმ (${bar.wastePercent.toFixed(1)}%)\n\n`;
    });

    navigator.clipboard.writeText(txt);
    setCopiedStates({ ...copiedStates, [profile]: true });
    setTimeout(() => {
      setCopiedStates({ ...copiedStates, [profile]: false });
    }, 2000);
  };

  // Toggle Collapse
  const toggleCollapse = (id: string) => {
    setCollapsedAss(prev => ({ ...prev, [id]: !prev[id] }));
  };

  // Total Sheet Metal calculation
  const totalSheetMetalArea = useMemo(() => {
    const areas: Record<string, { totalArea: number; wasteArea: number; itemsCount: number }> = {};
    Object.keys(compiledData.sheetMetals).forEach((key) => {
      let areaSum = 0;
      compiledData.sheetMetals[key].forEach((item) => {
        // Area in m^2: (length * width * qty) / 1,000,000
        areaSum += (item.length * item.width * item.qty) / 1000000;
      });

      areas[key] = {
        totalArea: areaSum,
        wasteArea: areaSum * sheetMetalWasteFactor,
        itemsCount: compiledData.sheetMetals[key].length
      };
    });
    return areas;
  }, [compiledData.sheetMetals, sheetMetalWasteFactor]);

  // Export nesting details to a clean Excel file using XLSX
  const exportToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Summary Table
      const summaryData = [];
      summaryData.push(['პროფილი / მასალა', 'რაოდენობა', 'ერთეული', 'დეტალური აღწერა']);

      Object.entries(solveResults).forEach(([profile, res]) => {
        summaryData.push([profile, res.totalBars, '6მ ღერო', `${res.yieldPercent.toFixed(1)}% სასარგებლო გამოყენება`]);
      });

      Object.entries(totalSheetMetalArea).forEach(([plate, stats]) => {
        summaryData.push([plate, stats.wasteArea.toFixed(3), 'კვ.მ (მ²)', `${(sheetMetalWasteFactor * 100).toFixed(0)}% კოეფიციენტით`]);
      });

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
      XLSX.utils.book_append_sheet(wb, wsSummary, 'ჯამური მასალები');

      // Sheet 2: Detailed Cut Instructions
      const cutsData = [];
      cutsData.push(['პროფილი', 'ღეროს #', 'გადანაჭრის #', 'სიგრძე (მმ)', 'პოზიცია', 'კონსტრუქცია', 'კოორდინატი (მმ)']);

      Object.entries(solveResults).forEach(([profile, res]) => {
        res.bars.forEach((bar, barIdx) => {
          bar.cuts.forEach((cut, cutIdx) => {
            cutsData.push([
              profile,
              `ღერო #${barIdx + 1}`,
              `გადანაჭერი #${cutIdx + 1}`,
              cut.length,
              cut.position,
              cut.constructionName,
              `${cut.startOffset} - ${cut.endOffset}`
            ]);
          });
          // Add a row for the waste of this bar
          cutsData.push([
            profile,
            `ღერო #${barIdx + 1}`,
            'ნარჩენი',
            bar.waste,
            '-',
            '-',
            `${res.stockLength - bar.waste} - ${res.stockLength}`
          ]);
        });
      });

      const wsCuts = XLSX.utils.aoa_to_sheet(cutsData);
      XLSX.utils.book_append_sheet(wb, wsCuts, 'პროფილების გადანაჭრები');

      // Sheet 3: Sheet Metal details
      if (Object.keys(compiledData.sheetMetals).length > 0) {
        const sheetData = [];
        sheetData.push(['ფურცელი', 'სიგრძე (მმ)', 'სიგანე (მმ)', 'რაოდენობა (ცალი)', 'კონსტრუქცია', 'სუფთა ფართობი (მ²)']);

        Object.keys(compiledData.sheetMetals).forEach((plate) => {
          const items = compiledData.sheetMetals[plate] || [];
          items.forEach((item) => {
            const area = (item.length * item.width * item.qty) / 1000000;
            sheetData.push([plate, item.length, item.width, item.qty, item.constructionName, area]);
          });
          
          const stats = totalSheetMetalArea[plate];
          sheetData.push([`ჯამი (${plate})`, '-', '-', '-', 'სუფთა ფართობი', stats.totalArea]);
          sheetData.push([`ჯამი კოეფიციენტით (${plate})`, '-', '-', '-', `ბრუტო (+${((sheetMetalWasteFactor-1)*100).toFixed(0)}%)`, stats.wasteArea]);
          sheetData.push([]); // empty spacer row
        });

        const wsSheets = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(wb, wsSheets, 'ფურცლოვანი ლითონი');
      }

      XLSX.writeFile(wb, 'მასალების_გაანგარიშება.xlsx');
    } catch (e) {
      alert('ექსპორტისას მოხდა შეცდომა: ' + e);
    }
  };

  // Group identical bars for visual rendering
  const getGroupedBars = (bars: NestedBar[]): GroupedBar[] => {
    const groups: GroupedBar[] = [];

    bars.forEach((bar, idx) => {
      // Use the sequence of cut lengths as the grouping key
      const key = bar.cuts.map((c) => c.length).join('-');
      const existing = groups.find((g) => g.cutsKey === key);
      if (existing) {
        existing.count += 1;
        existing.barIndices.push(idx + 1);
        existing.originalBars.push(bar);
      } else {
        groups.push({
          cutsKey: key,
          sampleBar: bar,
          count: 1,
          barIndices: [idx + 1],
          originalBars: [bar]
        });
      }
    });
    return groups;
  };

  // Get aggregated cut details for a grouped bar
  const getGroupedCutsInfo = (gBar: GroupedBar) => {
    const counts: Record<number, Record<string, { qty: number; constructionName: string }>> = {};
    
    gBar.originalBars.forEach((bar) => {
      bar.cuts.forEach((cut) => {
        if (!counts[cut.length]) counts[cut.length] = {};
        const posKey = `${cut.position} (${cut.constructionName})`;
        if (!counts[cut.length][posKey]) {
          counts[cut.length][posKey] = { qty: 0, constructionName: cut.constructionName };
        }
        counts[cut.length][posKey].qty += 1;
      });
    });

    return Object.entries(counts).map(([lenStr, posMap]) => {
      const length = parseFloat(lenStr);
      const positionsList = Object.entries(posMap).map(([posKey, info]) => {
        const pos = posKey.split(' (')[0];
        return {
          pos,
          constructionName: info.constructionName,
          qty: info.qty
        };
      });
      
      positionsList.sort((a, b) => a.pos.localeCompare(b.pos, undefined, { numeric: true }));
      const totalQty = positionsList.reduce((sum, p) => sum + p.qty, 0);

      return {
        length,
        totalQty,
        positionsList
      };
    }).sort((a, b) => b.length - a.length);
  };

  // Helper to group and sum cut items by length for a specific profile
  const getAggregatedCutsForProfile = (profile: string) => {
    const cuts = compiledData.profileCuts[profile] || [];
    const groups: Record<number, number> = {};
    cuts.forEach((c) => {
      groups[c.length] = (groups[c.length] || 0) + c.qty;
    });
    return Object.entries(groups)
      .map(([length, qty]) => ({ length: parseFloat(length), qty }))
      .sort((a, b) => b.length - a.length);
  };

  if (isPrinting) {
    return (
      <div className="print-report-container" style={{ padding: '2rem', backgroundColor: 'white', color: 'black', fontFamily: 'sans-serif' }}>
        {activeTab === 'editor' ? (
          <div>
            <h1 style={{ textAlign: 'center', marginBottom: '1.5rem', fontSize: '1.5rem', fontWeight: 'bold' }}>
              კონსტრუქციებისა და გადანაჭრების რედაქტორი (სამუშაო ფურცელი)
            </h1>
            {constructions.map((c) => (
              <div key={c.id} style={{ marginBottom: '2rem', borderBottom: '2px solid black', paddingBottom: '1rem' }}>
                <h3 style={{ fontSize: '1.1rem', marginBottom: '0.5rem', fontWeight: 'bold' }}>
                  პოზ. {c.position} - {c.name} (რაოდენობა: {c.qty}ც)
                </h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '0.875rem' }}>
                  <thead>
                    <tr>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600' }}>პოზ</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600' }}>ტიპი</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600' }}>აღწერილობა / პროფილი</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600', textAlign: 'right' }}>სიგრძე (მმ)</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600', textAlign: 'right' }}>სიგანე (მმ)</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600', textAlign: 'right' }}>რაოდ.</th>
                    </tr>
                  </thead>
                  <tbody>
                    {c.items.map((item) => (
                      <tr key={item.id}>
                        <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem' }}>{item.position}</td>
                        <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem' }}>
                          {item.isSheetMetal ? 'ფურცლოვანი' : 'პროფილი'}
                        </td>
                        <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem' }}>{item.profile}</td>
                        <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', textAlign: 'right' }}>{item.length}</td>
                        <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', textAlign: 'right' }}>
                          {item.isSheetMetal ? item.width : '-'}
                        </td>
                        <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', textAlign: 'right' }}>{item.qty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ))}
          </div>
        ) : (
          <div>
            <h1 style={{ textAlign: 'center', marginBottom: '1.5rem', fontSize: '1.5rem', fontWeight: 'bold' }}>
              მასალების გაანგარიშება და ჭრის რუკა (საამქრო რეპორტი)
            </h1>
            
            {/* Stats Summary */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '1rem', marginBottom: '2rem', border: '1px solid black', padding: '1rem', borderRadius: '4px' }}>
              <div>
                <div style={{ fontSize: '0.75rem', textTransform: 'uppercase', color: '#666' }}>ღეროების რაოდენობა</div>
                <div style={{ fontSize: '1.25rem', fontWeight: 'bold' }}>{totalsSummary.totalBars}ც ({settings.stockLength / 1000}მ)</div>
              </div>
              <div>
                <div style={{ fontSize: '0.75rem', textTransform: 'uppercase', color: '#666' }}>ჯამური გადანაჭრები</div>
                <div style={{ fontSize: '1.25rem', fontWeight: 'bold' }}>{totalsSummary.totalCuts}ც</div>
              </div>
              <div>
                <div style={{ fontSize: '0.75rem', textTransform: 'uppercase', color: '#666' }}>ნარჩენი სიგრძე</div>
                <div style={{ fontSize: '1.25rem', fontWeight: 'bold' }}>{(totalsSummary.totalScrapLength / 1000).toFixed(2)}მ</div>
              </div>
              <div>
                <div style={{ fontSize: '0.75rem', textTransform: 'uppercase', color: '#666' }}>გამოსავლიანობა (Yield)</div>
                <div style={{ fontSize: '1.25rem', fontWeight: 'bold' }}>{totalsSummary.averageYield.toFixed(1)}%</div>
              </div>
            </div>

            {/* Profiles Nesting Results */}
            {Object.entries(solveResults).map(([profile, res]) => (
              <div key={profile} style={{ marginBottom: '2.5rem', pageBreakInside: 'avoid' }}>
                <div style={{ background: '#f1f5f9', borderLeft: '4px solid black', padding: '0.5rem 1rem', marginBottom: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <h3 style={{ fontSize: '1.1rem', fontWeight: 'bold', margin: 0 }}>{profile}</h3>
                  <span style={{ fontSize: '0.9rem', fontWeight: '600' }}>
                    საჭიროა: {res.totalBars} ღერო | გამოსავლიანობა: {res.yieldPercent.toFixed(1)}%
                  </span>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                  {getGroupedBars(res.bars).map((gBar, gIdx) => {
                    const bar = gBar.sampleBar;
                    const trimPct = (bar.trim / bar.stockLength) * 100;
                    const groupedCuts = getGroupedCutsInfo(gBar);
                    const barTitle = gBar.count > 1
                      ? `ღეროები: #${gBar.barIndices[0]} - #${gBar.barIndices[gBar.barIndices.length - 1]} (${gBar.count} ცალი - იდენტური ჭრებით)`
                      : `ღერო #${gBar.barIndices[0]}`;

                    return (
                      <div key={gIdx} style={{ border: '1px solid black', padding: '0.75rem', borderRadius: '4px', pageBreakInside: 'avoid', marginBottom: '1rem' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.8rem', fontWeight: 'bold', marginBottom: '0.4rem' }}>
                          <span>{barTitle}</span>
                          <span>
                            გამოყენებული: {(bar.usedLength).toFixed(0)}მმ / {res.stockLength}მმ (ნარჩენი: {bar.waste}მმ)
                          </span>
                        </div>

                        {/* Visual bar layout for print */}
                        <div style={{ height: '22px', border: '1px solid black', display: 'flex', overflow: 'hidden', position: 'relative', marginBottom: '0.5rem', backgroundColor: 'white' }}>
                          {bar.trim > 0 && (
                            <div style={{ width: `${trimPct}%`, backgroundColor: '#ffcccc', height: '100%' }} />
                          )}
                          {bar.cuts.map((cut) => {
                            const cutPct = (cut.length / bar.stockLength) * 100;
                            return (
                              <div
                                key={cut.id}
                                style={{
                                  width: `${cutPct}%`,
                                  backgroundColor: '#f1f5f9',
                                  borderRight: '1px solid black',
                                  height: '100%',
                                  display: 'flex',
                                  alignItems: 'center',
                                  justifyContent: 'center',
                                  fontSize: '0.7rem',
                                  fontWeight: 'bold',
                                  color: 'black'
                                }}
                              >
                                {cut.length}
                              </div>
                            );
                          })}
                          {bar.waste > 0 && (
                            <div style={{ flexGrow: 1, backgroundColor: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.7rem', color: '#666', borderLeft: '1px dashed black' }}>
                              ნარჩენი ({bar.waste}მმ)
                            </div>
                          )}
                        </div>

                        {/* List cuts text */}
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem', fontSize: '0.75rem' }}>
                          {groupedCuts.map((item, idx) => {
                            const positionsText = item.positionsList.map(p => `პოზ ${p.pos} (${p.qty}ც)`).join(', ');
                            return (
                              <span key={idx} style={{ background: '#f1f5f9', padding: '0.15rem 0.4rem', borderRadius: '4px', border: '1px solid #ccc', color: 'black' }}>
                                გადანაჭერი: <strong>{item.length}მმ</strong> &times; {item.totalQty}ცალი ({positionsText})
                              </span>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}

            {/* Sheet Metals Output Panel */}
            <div style={{ marginTop: '2rem', pageBreakInside: 'avoid' }}>
              <h2 style={{ fontSize: '1.25rem', borderBottom: '2px solid black', paddingBottom: '0.5rem', marginBottom: '1rem', fontWeight: 'bold' }}>
                ფურცლოვანი ლითონის ჯამური ფართობი (მ²)
              </h2>
              {Object.keys(compiledData.sheetMetals).length === 0 ? (
                <p>კონსტრუქციებში ფურცლოვანი ლითონი ვერ მოიძებნა.</p>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '0.875rem' }}>
                  <thead>
                    <tr>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600' }}>ფურცლის სისქე/აღწერა</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600' }}>გამოყენებული ნაწილები</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600', textAlign: 'right' }}>ნაწილების ჯამური ფართობი</th>
                      <th style={{ borderBottom: '2px solid black', padding: '0.4rem', fontWeight: '600', textAlign: 'right' }}>საჭირო ფართობი (+{( (sheetMetalWasteFactor - 1) * 100 ).toFixed(0)}% დანაკარგი)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(compiledData.sheetMetals).map(([thickness, items]) => {
                      const netArea = items.reduce((sum, item) => sum + (item.length * item.width * item.qty) / 1000000, 0);
                      const grossArea = netArea * sheetMetalWasteFactor;
                      const partsList = items.map((item) => `${item.length}x${item.width} (პოზ ${item.position}) x${item.qty}`).join(', ');

                      return (
                        <tr key={thickness}>
                          <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', fontWeight: 'bold' }}>{thickness}</td>
                          <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', fontSize: '0.75rem', color: '#555' }}>{partsList}</td>
                          <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', textAlign: 'right' }}>{netArea.toFixed(2)} მ²</td>
                          <td style={{ borderBottom: '1px solid #ccc', padding: '0.4rem', textAlign: 'right', fontWeight: 'bold' }}>{grossArea.toFixed(2)} მ²</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="app-container">
      {/* Sidebar (Settings & Global API configurations) */}
      <aside className="sidebar no-print">
        <div className="brand">
          <div className="brand-icon">MC</div>
          <h1 className="brand-name">მასალის ნესტინგი</h1>
        </div>

        {/* Materials Summary widget */}
        <div className="glass-panel" style={{ borderColor: 'var(--border-color-glow)' }}>
          <h3 className="panel-title" style={{ marginBottom: '0.5rem', color: 'var(--primary)' }}>
            <Layers size={16} /> მასალის ჯამური შეკვეთა
          </h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.4rem', fontSize: '0.8rem' }}>
            {Object.entries(solveResults).map(([profile, res]) => {
              const aggregated = getAggregatedCutsForProfile(profile);
              return (
                <div key={profile} style={{ borderBottom: '1px solid var(--border-color)', paddingBottom: '0.4rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '600', marginBottom: '0.15rem' }}>
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '120px' }} title={profile}>{profile}</span>
                    <span style={{ color: 'var(--primary)' }}>{res.totalBars} ღერო <span style={{ fontWeight: 'normal', color: 'var(--text-muted)', fontSize: '0.75rem' }}>({res.stockLength / 1000}მ)</span></span>
                  </div>
                  <div style={{ paddingLeft: '0.5rem', display: 'flex', flexDirection: 'column', gap: '0.1rem', fontSize: '0.75rem', color: 'var(--text-secondary)' }}>
                    {aggregated.map((item, idx) => (
                      <div key={idx} style={{ display: 'flex', justifyContent: 'space-between' }}>
                        <span>↳ {item.length} მმ</span>
                        <span style={{ fontWeight: '600' }}>&times; {item.qty} ცალი</span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
            {Object.entries(totalSheetMetalArea).map(([plate, stats]) => {
              const items = compiledData.sheetMetals[plate] || [];
              return (
                <div key={plate} style={{ borderBottom: '1px solid var(--border-color)', paddingBottom: '0.4rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '600', marginBottom: '0.15rem' }}>
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '120px' }} title={plate}>{plate}</span>
                    <span style={{ color: 'var(--secondary)' }}>{stats.wasteArea.toFixed(2)} მ²</span>
                  </div>
                  <div style={{ paddingLeft: '0.5rem', display: 'flex', flexDirection: 'column', gap: '0.1rem', fontSize: '0.75rem', color: 'var(--text-secondary)' }}>
                    {items.map((item, idx) => (
                      <div key={idx} style={{ display: 'flex', justifyContent: 'space-between' }}>
                        <span>↳ {item.length} &times; {item.width} მმ</span>
                        <span style={{ fontWeight: '600' }}>&times; {item.qty} ცალი</span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
            {Object.keys(solveResults).length === 0 && Object.keys(totalSheetMetalArea).length === 0 && (
              <span style={{ color: 'var(--text-muted)', fontStyle: 'italic', textAlign: 'center', display: 'block', padding: '0.5rem 0' }}>მასალები არ არის დამატებული</span>
            )}
          </div>
        </div>

        <div className="glass-panel">
          <h3 className="panel-title" style={{ marginBottom: '1rem' }}>
            <Settings size={18} /> გლობალური პარამეტრები
          </h3>
          <div className="settings-group">
            <div className="input-field">
              <label className="input-label">საწესდებო სიგრძე (მმ)</label>
              <input
                type="number"
                className="form-control"
                value={settings.stockLength}
                onChange={(e) =>
                  setSettings({ ...settings, stockLength: Math.max(1, parseInt(e.target.value, 10) || 0) })
                }
              />
            </div>
            <div className="input-field">
              <label className="input-label">ხერხის სისქე (მმ)</label>
              <input
                type="number"
                className="form-control"
                value={settings.kerf}
                onChange={(e) =>
                  setSettings({ ...settings, kerf: Math.max(0, parseFloat(e.target.value) || 0) })
                }
              />
            </div>
            <div className="input-field">
              <label className="input-label">კიდის ჩამონაჭერი (მმ)</label>
              <input
                type="number"
                className="form-control"
                value={settings.trim}
                onChange={(e) =>
                  setSettings({ ...settings, trim: Math.max(0, parseInt(e.target.value, 10) || 0) })
                }
                placeholder="ორმხრივი ჩამონაჭერი"
              />
            </div>
          </div>
        </div>

        <div className="glass-panel">
          <h3 className="panel-title" style={{ marginBottom: '1rem' }}>
            <ImageIcon size={18} /> Gemini API გასაღები
          </h3>
          <div className="settings-group">
            <p className="input-label" style={{ textTransform: 'none', color: 'var(--text-secondary)' }}>
              საჭიროა სკრინშოტის გასაანალიზებლად. ინახება თქვენს ბრაუზერში.
            </p>
            <input
              type="password"
              className="form-control"
              placeholder="AIzaSy..."
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
            />
            <a
              href="https://aistudio.google.com/"
              target="_blank"
              rel="noopener noreferrer"
              style={{ color: 'var(--primary)', fontSize: '0.75rem', textDecoration: 'none' }}
            >
              მიიღეთ უფასო API გასაღები Google AI Studio-დან &rarr;
            </a>
          </div>
        </div>

        <div className="glass-panel">
          <h3 className="panel-title" style={{ marginBottom: '1rem' }}>
            <Layers size={18} /> ფურცლოვანი ლითონის კოეფიციენტი
          </h3>
          <div className="settings-group">
            <div className="input-field">
              <label className="input-label">გადანაჭრის კოეფიციენტი</label>
              <input
                type="number"
                step="0.05"
                className="form-control"
                value={sheetMetalWasteFactor}
                onChange={(e) => setSheetMetalWasteFactor(Math.max(1, parseFloat(e.target.value) || 1))}
              />
              <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>
                ამრავლებს ფართობს. მაგ. 1.15 ნიშნავს +15% დანაკარგს.
              </span>
            </div>
          </div>
        </div>

        <div style={{ marginTop: 'auto', display: 'flex', gap: '0.5rem' }}>
          <button
            className="btn btn-secondary"
            style={{ flex: 1 }}
            onClick={() => {
              if (confirm('გსურთ მონაცემების გასუფთავება და თავიდან დაწყება?')) {
                setConstructions([]);
                setActiveTab('editor');
              }
            }}
          >
            გასუფთავება
          </button>
          <button
            className="btn btn-primary"
            style={{ flex: 1 }}
            onClick={() => setIsPrinting(true)}
          >
            <Printer size={16} /> ბეჭდვა
          </button>
        </div>
      </aside>

      {/* Main Content Dashboard */}
      <main className="main-content">
        {/* Navigation Tabs */}
        <div className="tabs-container no-print">
          <button
            className={`tab-btn ${activeTab === 'editor' ? 'active' : ''}`}
            onClick={() => setActiveTab('editor')}
          >
            <Layers size={16} /> კონსტრუქციების რედაქტირება ({constructions.length})
          </button>
          <button
            className={`tab-btn ${activeTab === 'paste' ? 'active' : ''}`}
            onClick={() => setActiveTab('paste')}
          >
            <FileSpreadsheet size={16} /> ცხრილის ჩასმა ან ატვირთვა
          </button>
          <button
            className={`tab-btn ${activeTab === 'ai' ? 'active' : ''}`}
            onClick={() => setActiveTab('ai')}
          >
            <ImageIcon size={16} /> AI სკრინშოტის ატვირთვა
          </button>
          <button
            className={`tab-btn ${activeTab === 'results' ? 'active' : ''}`}
            onClick={() => setActiveTab('results')}
            style={{ marginLeft: 'auto', background: 'var(--primary-glow)', color: 'white' }}
          >
            <Play size={16} /> გაანგარიშება და ვიზუალიზაცია
          </button>
        </div>

        {/* Tab content 1: Visual Cut Lists / Editor */}
        <div className={`editor-tab tab-content-panel ${activeTab === 'editor' ? 'active' : ''}`}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
              <div>
                <h2>კონსტრუქციებისა და გადანაჭრების რედაქტორი</h2>
                <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>
                  განსაზღვრეთ კონსტრუქციები და პროფილების ზომები. ძირითადი რაოდენობა ამრავლებს შიდა გადანაჭრებს.
                </p>
              </div>
              <button className="btn btn-primary" onClick={addConstruction}>
                <Plus size={16} /> კონსტრუქციის დამატება
              </button>
            </div>

            {constructions.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '4rem 2rem', color: 'var(--text-secondary)' }} className="glass-panel">
                <Layers size={40} style={{ color: 'var(--primary)', marginBottom: '1rem', display: 'inline-block' }} />
                <h3 style={{ color: 'white', marginBottom: '0.5rem' }}>კონსტრუქციები არ არის დამატებული</h3>
                <p style={{ fontSize: '0.875rem', marginBottom: '1.5rem', maxWidth: '400px', margin: '0 auto 1.5rem auto' }}>
                  დაიწყეთ კონსტრუქციის ხელით დამატებით, Excel-ის ცხრილის ჩასმით ან BOM-ის სკრინშოტის ატვირთვით.
                </p>
                <div style={{ display: 'flex', gap: '0.75rem', justifyContent: 'center', flexWrap: 'wrap' }}>
                  <button className="btn btn-primary btn-sm" onClick={addConstruction}>
                    <Plus size={14} /> კონსტრუქციის დამატება
                  </button>
                  <button className="btn btn-secondary btn-sm" onClick={() => setActiveTab('paste')}>
                    ცხრილის ჩასმა
                  </button>
                  <button className="btn btn-secondary btn-sm" onClick={() => setActiveTab('ai')}>
                    AI სკრინშოტი
                  </button>
                  <button className="btn btn-secondary btn-sm" onClick={() => setConstructions(INITIAL_CONSTRUCTIONS)} style={{ borderStyle: 'dashed' }}>
                    მოაჯირის ნიმუშის ჩატვირთვა
                  </button>
                </div>
              </div>
            ) : (
              constructions.map((c) => {
                const isCollapsed = collapsedAss[c.id];
              return (
                <div key={c.id} className="glass-panel construction-card">
                  <div className="construction-header">
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', flexGrow: 1 }}>
                      <button 
                        onClick={() => toggleCollapse(c.id)}
                        className="btn btn-secondary btn-sm"
                        style={{ padding: '0.2rem' }}
                      >
                        {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
                      </button>
                      <input
                        type="text"
                        className="form-control"
                        style={{ width: '40px', textAlign: 'center', fontWeight: 'bold' }}
                        value={c.position}
                        onChange={(e) => updateConstructionHeader(c.id, 'position', e.target.value)}
                      />
                      <input
                        type="text"
                        className="form-control"
                        style={{ fontWeight: 'bold', fontSize: '1rem', background: 'transparent', border: 'none' }}
                        value={c.name}
                        onChange={(e) => updateConstructionHeader(c.id, 'name', e.target.value)}
                      />
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                      <div className="input-field" style={{ flexDirection: 'row', alignItems: 'center', gap: '0.5rem' }}>
                        <label className="input-label" style={{ whiteSpace: 'nowrap' }}>რაოდენობა</label>
                        <input
                          type="number"
                          className="form-control"
                          style={{ width: '80px' }}
                          value={c.qty}
                          onChange={(e) => updateConstructionHeader(c.id, 'qty', e.target.value)}
                        />
                      </div>
                      <button
                        className="btn btn-danger btn-sm"
                        onClick={() => removeConstruction(c.id)}
                        title="კონსტრუქციის წაშლა"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </div>

                  {!isCollapsed && (
                    <div className="table-container">
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th style={{ width: '80px' }}>პოზ</th>
                            <th style={{ width: '120px' }}>ტიპი</th>
                            <th>აღწერა / პროფილის ზომა</th>
                            <th style={{ width: '120px' }}>სიგრძე (მმ)</th>
                            <th style={{ width: '120px' }}>სიგანე (მმ)</th>
                            <th style={{ width: '100px' }}>რაოდ. (ცალი)</th>
                            <th style={{ width: '120px' }}>ჯამ. რაოდენობა</th>
                            <th style={{ width: '70px' }}>მოქმედებები</th>
                          </tr>
                        </thead>
                        <tbody>
                          {c.items.map((item) => (
                            <tr key={item.id}>
                              <td>
                                <input
                                  type="text"
                                  className="form-control"
                                  value={item.position}
                                  onChange={(e) =>
                                    updateItemInConstruction(c.id, item.id, 'position', e.target.value)
                                  }
                                />
                              </td>
                              <td>
                                <select
                                  className="form-control"
                                  value={item.isSheetMetal ? 'sheet' : 'profile'}
                                  onChange={(e) =>
                                    updateItemInConstruction(
                                      c.id,
                                      item.id,
                                      'isSheetMetal',
                                      e.target.value === 'sheet'
                                    )
                                  }
                                >
                                  <option value="profile">პროფილი / მილი</option>
                                  <option value="sheet">ფურცლოვანი ლითონი</option>
                                </select>
                              </td>
                              <td>
                                <input
                                  type="text"
                                  className="form-control"
                                  value={item.profile}
                                  onChange={(e) =>
                                    updateItemInConstruction(c.id, item.id, 'profile', e.target.value)
                                  }
                                  placeholder={item.isSheetMetal ? 'ფურცელი 3მმ' : '50 x 50 x 3.2'}
                                />
                              </td>
                              <td>
                                <input
                                  type="number"
                                  className="form-control"
                                  value={item.length}
                                  onChange={(e) =>
                                    updateItemInConstruction(c.id, item.id, 'length', e.target.value)
                                  }
                                />
                              </td>
                              <td>
                                {item.isSheetMetal ? (
                                  <input
                                    type="number"
                                    className="form-control"
                                    value={item.width || 0}
                                    onChange={(e) =>
                                      updateItemInConstruction(c.id, item.id, 'width', e.target.value)
                                    }
                                  />
                                ) : (
                                  <span style={{ color: 'var(--text-muted)' }}>-</span>
                                )}
                              </td>
                              <td>
                                <input
                                  type="number"
                                  className="form-control"
                                  value={item.qty}
                                  onChange={(e) =>
                                    updateItemInConstruction(c.id, item.id, 'qty', e.target.value)
                                  }
                                />
                              </td>
                              <td style={{ fontWeight: '600' }}>{item.qty * c.qty}</td>
                              <td>
                                <button
                                  className="btn btn-danger btn-sm"
                                  onClick={() => removeItemFromConstruction(c.id, item.id)}
                                >
                                  <Trash2 size={12} />
                                </button>
                              </td>
                            </tr>
                          ))}
                          <tr>
                            <td colSpan={8} style={{ padding: '0.75rem' }}>
                              <div style={{ display: 'flex', gap: '0.5rem' }}>
                                <button
                                  className="btn btn-secondary btn-sm"
                                  onClick={() => addItemToConstruction(c.id, false)}
                                >
                                  <Plus size={12} /> პროფილის/მილის გადანაჭრის დამატება
                                </button>
                                <button
                                  className="btn btn-secondary btn-sm"
                                  onClick={() => addItemToConstruction(c.id, true)}
                                >
                                  <Plus size={12} /> ფურცლოვანი ლითონის დამატება
                                </button>
                              </div>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            }))}
            {constructions.length > 0 && (
              <div style={{ display: 'flex', gap: '0.5rem', marginTop: '1rem' }} className="no-print">
                <button className="btn btn-primary" onClick={addConstruction}>
                  <Plus size={16} /> კონსტრუქციის დამატება
                </button>
              </div>
            )}
        </div>

        {/* Tab content 2: Quick Copy-Paste Excel */}
        <div className={`paste-tab tab-content-panel ${activeTab === 'paste' ? 'active' : ''}`}>
            <div className="glass-panel">
              <h2 style={{ marginBottom: '0.5rem' }}>ცხრილების იმპორტი (Excel / CSV) ან ჩასმა</h2>
              <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>
                შეგიძლიათ ატვირთოთ `.xlsx`, `.xls` ან `.csv` ფაილი, ან ჩაასვათ Excel-იდან ან Solidworks-იდან კოპირებული სვეტები.
              </p>
            </div>

            <div className="grid-2">
              {/* Document Upload panel */}
              <div className="glass-panel" style={{ display: 'flex', flexDirection: 'column', minHeight: '350px' }}>
                <h3 className="panel-title" style={{ marginBottom: '1rem' }}>
                  <FileSpreadsheet size={18} style={{ color: 'var(--secondary)' }} /> ვარიანტი 1: ცხრილის ფაილის ატვირთვა
                </h3>
                
                <div
                  className={`dropzone ${dragActive ? 'active' : ''}`}
                  onDragEnter={handleDrag}
                  onDragOver={handleDrag}
                  onDragLeave={handleDrag}
                  onDrop={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    setDragActive(false);
                    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                      handleExcelUpload(e.dataTransfer.files[0]);
                    }
                  }}
                  onClick={() => {
                    const el = document.createElement('input');
                    el.type = 'file';
                    el.accept = '.xlsx,.xls,.csv';
                    el.onchange = (e: any) => {
                      if (e.target.files && e.target.files[0]) {
                        handleExcelUpload(e.target.files[0]);
                      }
                    };
                    el.click();
                  }}
                  style={{ flexGrow: 1, justifyContent: 'center', minHeight: '220px' }}
                >
                  <FileSpreadsheet className="dropzone-icon" style={{ color: 'var(--secondary)' }} />
                  <p style={{ fontWeight: '600' }}>ჩააგდეთ Excel ან CSV ფაილი აქ</p>
                  <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>მხარდაჭერილია .xlsx, .xls, .csv ფაილები</p>
                </div>
              </div>

              {/* Text Paste panel */}
              <div className="glass-panel" style={{ display: 'flex', flexDirection: 'column', minHeight: '350px' }}>
                <h3 className="panel-title" style={{ marginBottom: '1rem' }}>
                  <Copy size={18} style={{ color: 'var(--primary)' }} /> ვარიანტი 2: ცხრილის ტექსტის ჩასმა
                </h3>
                
                <textarea
                  className="form-control"
                  style={{ flexGrow: 1, fontFamily: 'monospace', fontSize: '0.8rem', minHeight: '160px', marginBottom: '1rem', resize: 'vertical' }}
                  placeholder="პოზიცია&#9;რაოდ&#9;აღწერა&#9;სიგრძე&#10;1.1&#9;2&#9;50 x 50 x 3.2&#9;1090&#10;1.2&#9;1&#9;50 x 30 x 2.6&#9;2000"
                  value={rawPasteText}
                  onChange={(e) => setRawPasteText(e.target.value)}
                />

                <div style={{ display: 'flex', gap: '0.5rem', marginTop: 'auto' }}>
                  <button className="btn btn-primary" style={{ flexGrow: 1 }} onClick={handlePasteImport}>
                    გაანალიზება და ჩატვირთვა
                  </button>
                  <button
                    className="btn btn-secondary"
                    onClick={() =>
                      setRawPasteText(
                        `პოზიცია\tრაოდ\tაღწერა\tსიგრძე\n1\t1\tმოაჯირი - ორივე მხრით\n1.1\t2\t50 x 50 x 3.2\t1090\n1.2\t1\t50 x 30 x 2.6\t2000\n1.3\t15\t20 x 20 x 2.0\t960`
                      )
                    }
                  >
                    ნიმუშის ჩატვირთვა
                  </button>
                </div>
              </div>
            </div>
        </div>

        {/* Tab content 3: AI Screenshot Upload */}
        <div className={`ai-tab tab-content-panel ${activeTab === 'ai' ? 'active' : ''}`}>
            <div className="glass-panel">
              <h2 style={{ marginBottom: '0.5rem' }}>BOM-ის სკრინშოტის AI ანალიზი</h2>
              <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem', marginBottom: '1rem' }}>
                ატვირთეთ ან ჩააგდეთ BOM-ის ან გადანაჭრების ცხრილის სკრინშოტი. ხელოვნური ინტელექტი ავტომატურად ამოიკითხავს კონსტრუქციებს, პროფილებსა და ზომებს.
              </p>
              
              {!apiKey && (
                <div className="alert-banner warning">
                  <AlertTriangle size={18} />
                  <span>
                    ამ ფუნქციის გამოსაყენებლად საჭიროა მარცხენა მენიუში შეიყვანოთ <strong>Gemini API გასაღები</strong>.
                  </span>
                </div>
              )}
            </div>

            <div className="grid-2">
              {/* Left card: Image upload dropzone */}
              <div className="glass-panel" style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                <div className="panel-header">
                  <h3 className="panel-title">სკრინშოტის ატვირთვა</h3>
                  {imagePreview && (
                    <button className="btn btn-secondary btn-sm" onClick={clearImage}>
                      სურათის გასუფთავება
                    </button>
                  )}
                </div>

                {!imagePreview ? (
                  <div
                    className={`dropzone ${dragActive ? 'active' : ''}`}
                    onDragEnter={handleDrag}
                    onDragOver={handleDrag}
                    onDragLeave={handleDrag}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                    style={{ flexGrow: 1, justifyContent: 'center' }}
                  >
                    <input
                      type="file"
                      ref={fileInputRef}
                      style={{ display: 'none' }}
                      accept="image/*"
                      onChange={handleFileChange}
                    />
                    <Upload className="dropzone-icon" />
                    <p style={{ fontWeight: '600' }}>ჩააგდეთ სკრინშოტი აქ</p>
                    <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>მხარდაჭერილია PNG, JPG, JPEG</p>
                  </div>
                ) : (
                  <div style={{ flexGrow: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(0,0,0,0.4)', borderRadius: '10px', overflow: 'hidden', padding: '1rem', minHeight: '300px' }}>
                    <img
                      src={imagePreview}
                      alt="BOM სკრინშოტის წინასწარი ნახვა"
                      style={{ maxWidth: '100%', maxHeight: '400px', objectFit: 'contain', borderRadius: '6px' }}
                    />
                  </div>
                )}
              </div>

              {/* Right card: Logs and Trigger */}
              <div className="glass-panel" style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                <h3 className="panel-title" style={{ marginBottom: '1.25rem' }}>AI მოქმედებების პანელი</h3>
                
                <div style={{ flexGrow: 1, marginBottom: '1.5rem' }}>
                  {aiLoading ? (
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', gap: '1rem' }}>
                      <RefreshCw className="animate-spin" size={36} style={{ color: 'var(--primary)', animation: 'spin 1.5s linear infinite' }} />
                      <p style={{ fontWeight: '500' }}>AI აანალიზებს ცხრილს...</p>
                      
                      <div style={{ width: '100%', maxWidth: '300px', background: '#0b0f19', border: '1px solid var(--border-color)', padding: '0.75rem', borderRadius: '6px', fontSize: '0.75rem', fontFamily: 'monospace' }}>
                        {aiLogs.map((log, idx) => (
                          <div key={idx} style={{ color: idx === aiLogs.length - 1 ? 'white' : 'var(--text-muted)' }}>
                            &gt; {log}
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : aiError ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', color: 'var(--danger)', height: '100%', justifyContent: 'center', alignItems: 'center', textAlign: 'center' }}>
                      <AlertTriangle size={32} />
                      <p style={{ fontWeight: '600' }}>ამოკითხვა ვერ მოხერხდა</p>
                      <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>{aiError}</p>
                    </div>
                  ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', justifyContent: 'center', alignItems: 'center', textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>
                      <Info size={32} style={{ color: 'var(--primary)', marginBottom: '1rem' }} />
                      <p style={{ fontWeight: '500', color: 'white', marginBottom: '0.5rem' }}>მზადაა ანალიზისთვის</p>
                      <p style={{ fontSize: '0.85rem' }}>
                        სკრინშოტის არჩევის შემდეგ, დააჭირეთ "BOM ცხრილის ამოკითხვა"-ს AI ანალიზის დასაწყებად.
                      </p>
                    </div>
                  )}
                </div>

                <button
                  className="btn btn-primary"
                  style={{ width: '100%', padding: '1rem' }}
                  disabled={!imagePreview || aiLoading || !apiKey}
                  onClick={runAIExtraction}
                >
                  {aiLoading ? 'ამოიკითხება...' : 'BOM ცხრილის ამოკითხვა AI-ით'}
                </button>
              </div>
            </div>
        </div>

        {/* Tab content 4: Solve & Visualise Results */}
        <div className={`results-tab tab-content-panel ${activeTab === 'results' ? 'active' : ''}`}>
            
            {/* Statistics Row */}
            <div className="glass-panel">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.25rem', flexWrap: 'wrap', gap: '0.75rem' }}>
                <h2 style={{ margin: 0 }}>შეკვეთის გლობალური სტატისტიკა</h2>
                <div style={{ display: 'flex', gap: '0.5rem' }} className="no-print">
                  <button className="btn btn-secondary" onClick={exportToExcel}>
                    <FileSpreadsheet size={16} style={{ color: 'var(--secondary)' }} /> Excel ექსპორტი
                  </button>
                  <button className="btn btn-primary" onClick={() => setIsPrinting(true)}>
                    <Printer size={16} /> ბეჭდვა
                  </button>
                </div>
              </div>
              <div className="stats-grid">
                <div className="stat-card">
                  <div className="stat-label">საჭირო მასალის რაოდენობა (ღერო)</div>
                  <div className="stat-value" style={{ color: 'var(--primary)' }}>
                    {totalsSummary.totalBars} <span style={{ fontSize: '0.9rem', fontWeight: '500', color: 'var(--text-secondary)' }}>({settings.stockLength / 1000}მ თითო)</span>
                  </div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">ჯამური გადანაჭრები</div>
                  <div className="stat-value">{totalsSummary.totalCuts}</div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">ჯამური ნარჩენი / გადანაჭრები</div>
                  <div className="stat-value" style={{ color: 'var(--warning)' }}>
                    {(totalsSummary.totalScrapLength / 1000).toFixed(2)} მ
                  </div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">მასალის სასარგებლო გამოყენება</div>
                  <div className="stat-value" style={{ color: 'var(--secondary)' }}>
                    {totalsSummary.averageYield.toFixed(1)}%
                  </div>
                </div>
              </div>
            </div>

            {/* Profile Nesting Layouts */}
            <div className="glass-panel">
              <h2 style={{ marginBottom: '1.5rem' }}>პროფილების ნესტინგის დიაგრამები (1D)</h2>
              
              {compiledData.profilesList.length === 0 ? (
                <div style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-secondary)' }}>
                  <Info size={32} style={{ margin: '0 auto 1rem auto' }} />
                  <p>კონსტრუქციებში პროფილის გადანაჭრები ვერ მოიძებნა.</p>
                  <p style={{ fontSize: '0.8rem' }}>გადადით "კონსტრუქციების რედაქტირებაში" მილების/პროფილების დასამატებლად.</p>
                </div>
              ) : (
                compiledData.profilesList.map((profile, pIdx) => {
                  const res = solveResults[profile];
                  if (!res) return null;

                  return (
                    <div key={profile} className="nesting-profile-section">
                      <div className="profile-title-bar">
                        <div>
                          <h3 style={{ fontSize: '1.15rem' }}>
                            პროფილი: {profile}
                          </h3>
                          <p style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>
                            საწესდებო სიგრძე: {res.stockLength}მმ | ხერხი: {res.kerf}მმ | კიდის ჩამონაჭერი: {res.trim}მმ
                          </p>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem' }}>
                          <div style={{ fontSize: '0.875rem', textAlign: 'right' }}>
                            <span style={{ fontWeight: 'bold', color: 'var(--primary)' }}>{res.totalBars} ღერო</span>
                            <span style={{ color: 'var(--text-muted)', margin: '0 0.5rem' }}>|</span>
                            <span>სასარგებლო: <strong>{res.yieldPercent.toFixed(1)}%</strong></span>
                          </div>
                          <button
                            className="btn btn-secondary btn-sm"
                            onClick={() => copyCutInstructions(profile, res)}
                          >
                            {copiedStates[profile] ? <Check size={14} style={{ color: 'var(--secondary)' }} /> : <Copy size={14} />}
                            {copiedStates[profile] ? 'კოპირებულია!' : 'გადანაჭრების კოპირება'}
                          </button>
                        </div>
                      </div>

                      {/* Nesting Bars representation */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                        {getGroupedBars(res.bars).map((gBar, gIdx) => {
                          const bar = gBar.sampleBar;
                          const trimPct = (bar.trim / bar.stockLength) * 100;
                          const groupedCuts = getGroupedCutsInfo(gBar);
                          
                          // Format bar title
                          const barTitle = gBar.count > 1
                            ? `ღეროები: #${gBar.barIndices[0]} - #${gBar.barIndices[gBar.barIndices.length - 1]} (${gBar.count} ცალი - იდენტური გადანაჭრებით)`
                            : `ღერო #${gBar.barIndices[0]}`;

                          return (
                            <div key={gBar.cutsKey + '-' + gIdx} className="bar-row">
                              <div className="bar-info">
                                <span style={{ fontWeight: '600', color: gBar.count > 1 ? 'var(--secondary)' : 'var(--text-secondary)' }}>
                                  {barTitle}
                                </span>
                                <span>
                                  გამოყენებული: {(bar.usedLength).toFixed(0)}მმ / {res.stockLength}მმ 
                                  <span style={{ color: 'var(--text-muted)' }}> (ნარჩენი: {bar.waste}მმ)</span>
                                </span>
                              </div>

                              <div className="visual-bar-track">
                                {/* Left trim segment */}
                                {bar.trim > 0 && (
                                  <div 
                                    className="trim-segment" 
                                    style={{ width: `${trimPct}%` }}
                                    title={`ჩამონაჭერი: ${bar.trim}მმ`}
                                  />
                                )}

                                {/* Cut segments */}
                                {bar.cuts.map((cut) => {
                                  const cutPct = (cut.length / bar.stockLength) * 100;
                                  return (
                                    <div
                                      key={cut.id}
                                      className="nested-cut-segment"
                                      style={{
                                        width: `${cutPct}%`,
                                        backgroundColor: getProfileColor(profile, pIdx)
                                      }}
                                    >
                                      {cut.length}
                                      <div className="nested-cut-tooltip">
                                        <strong>პოზ {cut.position}</strong> ({cut.length}მმ)<br />
                                        კონსტრუქცია: {cut.constructionName}<br />
                                        კოორდინატი: {cut.startOffset}მმ &rarr; {cut.endOffset}მმ
                                      </div>
                                    </div>
                                  );
                                })}

                                {/* Remaining Waste segment */}
                                {bar.waste > 0 && (
                                  <div 
                                    className="waste-segment" 
                                    style={{ width: `${(bar.waste / bar.stockLength) * 100}%` }}
                                  >
                                    ნარჩენი ({bar.waste}მმ)
                                  </div>
                                )}

                                {/* Right trim segment */}
                                {bar.trim > 0 && (
                                  <div 
                                    className="trim-segment" 
                                    style={{ width: `${trimPct}%` }}
                                    title={`ჩამონაჭერი: ${bar.trim}მმ`}
                                  />
                                )}
                              </div>

                              {/* Listing Cuts under each grouped bar */}
                              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem', marginTop: '0.4rem', fontSize: '0.75rem', color: 'var(--text-secondary)' }}>
                                {groupedCuts.map((item, idx) => {
                                  const positionsText = item.positionsList.map(p => `პოზ ${p.pos} (${p.qty}ც)`).join(', ');
                                  return (
                                    <span key={idx} style={{ background: 'rgba(255,255,255,0.03)', padding: '0.15rem 0.4rem', borderRadius: '4px', border: '1px solid var(--border-color)' }}>
                                      გადანაჭერი: <strong>{item.length}მმ</strong> &times; {item.totalQty}ცალი ({positionsText})
                                    </span>
                                  );
                                })}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            {/* Sheet Metals Output Panel */}
            <div className="glass-panel">
              <h2 style={{ marginBottom: '0.5rem' }}>ფურცლოვანი ლითონის ჯამური ფართობი (მ²)</h2>
              <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem', marginBottom: '1.5rem' }}>
                საჭირო ფურცლების შეჯამება. ის აჯამებს ფართობებს და ითვალისწინებს დანაკარგის კოეფიციენტს (+{( (sheetMetalWasteFactor - 1) * 100 ).toFixed(0)}% დანაკარგი).
              </p>

              {Object.keys(compiledData.sheetMetals).length === 0 ? (
                <div style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>
                  <Info size={24} style={{ margin: '0 auto 0.5rem auto' }} />
                  <p>კონსტრუქციებში ფურცლოვანი ლითონი ვერ მოიძებნა.</p>
                </div>
              ) : (
                <div className="table-container">
                  <table className="custom-table">
                    <thead>
                      <tr>
                        <th>ფურცელი / სისქე</th>
                        <th>დეტალების შეჯამება</th>
                        <th>ჯამური სუფთა ფართობი (მ²)</th>
                        <th>ჯამური ბრუტო ფართობი (მ²) <span style={{ fontWeight: 'normal', color: 'var(--text-muted)' }}>({(sheetMetalWasteFactor * 100).toFixed(0)}% კოეფიციენტი)</span></th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.keys(compiledData.sheetMetals).map((key) => {
                        const stats = totalSheetMetalArea[key];
                        return (
                          <tr key={key}>
                            <td style={{ fontWeight: '600' }}>{key}</td>
                            <td>
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.2rem' }}>
                                {compiledData.sheetMetals[key].map((item) => (
                                  <span key={item.id} style={{ fontSize: '0.75rem' }}>
                                    - {item.length}x{item.width} მმ &times; {item.qty} ცალი ({item.constructionName})
                                  </span>
                                ))}
                              </div>
                            </td>
                            <td>{stats.totalArea.toFixed(3)} მ²</td>
                            <td style={{ color: 'var(--secondary)', fontWeight: 'bold' }}>
                              {stats.wasteArea.toFixed(3)} მ²
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

        </div>
      </main>
    </div>
  );
}
