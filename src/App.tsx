import React, { useState, useRef, useEffect, useCallback } from "react";
import {
  Upload,
  Printer,
  Settings2,
  Image as ImageIcon,
  FileText,
  FileImage,
  Type,
  Palette,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Move,
  RotateCcw,
  Eye,
  EyeOff,
  Layers,
  Sparkles,
  Download,
  FolderOpen,
} from "lucide-react";
import { toPng, toJpeg } from "html-to-image";
import { motion } from "motion/react";
import { jsPDF } from "jspdf";
import { Document, Packer, Paragraph, ImageRun } from "docx";
import { saveAs } from "file-saver";

// Base64 to Uint8Array converter for docx image insertion
function base64ToUint8Array(base64: string) {
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

type TextStyle = {
  fontSize: number;
  isBold: boolean;
  isItalic: boolean;
  fontFamily: string;
  textAlign?: "left" | "center" | "right";
  marginTop?: number;
  marginBottom?: number;
  paddingTop?: number;
  paddingBottom?: number;
  textTransform?: "none" | "uppercase" | "lowercase" | "capitalize";
};

const defaultStyle = (
  fontSize: number,
  isBold: boolean = false,
  textAlign?: "center" | "left" | "right",
  fontFamily: string = "inherit",
): TextStyle => ({
  fontSize,
  isBold,
  isItalic: false,
  fontFamily,
  textAlign: textAlign || "center",
  marginTop: 0,
  marginBottom: 0,
  paddingTop: 0,
  paddingBottom: 0,
  textTransform: "capitalize",
});

const DEFAULT_DATA = {
  collegeName: "Padmashree College",
  address: "Tinkune, Kathmandu",
  affiliation: "Affiliated to Tribhuvan University",
  program: "Bachelor Of Computer Application (BCA)",
  subject: "Icarus and the sun",
  assignmentType: "Lab Report",
  topic: "Falling means i once soared",
  submittedByTitle: "Submitted By",
  submitterNameLabel: "Name:",
  submitterName: "Sumit pokhrel",
  submitterRollLabel: "Symbol No.:",
  submitterRoll: "20420/80",
  submitterSemLabel: "Semester:",
  submitterSem: "5th Semester",
  submitterDateLabel: "Date:",
  submitterDate: "",
  submittedToTitle: "Submitted To",
  submittedToName: "Prof. Dr. Sarah Jenkins",
  logoUrl: "logo.png",
};

const DEFAULT_TEXT_STYLES = {
  collegeName: defaultStyle(34, true),
  address: defaultStyle(18, false),
  affiliation: defaultStyle(16, false),
  program: defaultStyle(24, true),
  subject: defaultStyle(32, true),
  assignmentType: defaultStyle(28, true),
  topic: defaultStyle(22, true, "center", "'Courier New', monospace"),
  submittedByTitle: defaultStyle(24, true),
  submitterNameLabel: defaultStyle(18, true),
  submitterName: defaultStyle(18, true, "center", "'Courier New', monospace"),
  submitterRollLabel: defaultStyle(18, true),
  submitterRoll: defaultStyle(18, true, "center", "'Courier New', monospace"),
  submitterSemLabel: defaultStyle(18, true, "left"),
  submitterSem: defaultStyle(18, true, "center", "'Courier New', monospace"),
  submitterDateLabel: defaultStyle(18, true, "left"),
  submitterDate: defaultStyle(18, true, "center", "'Courier New', monospace"),
  submittedToTitle: defaultStyle(24, true, "center"),
  submittedToName: defaultStyle(20, true),
};

const DEFAULT_GLOBAL_STYLES = {
  fontFamily: "'Arial', 'Helvetica', sans-serif",
  textColor: "#0f172a",
  borderColor: "#1e3a8a",
  borderWidth: 6,
  borderStyle: "solid",
  backgroundColor: "#ffffff",
  padding: 60,
};

const DEFAULT_VISIBLE_SECTIONS = {
  address: true,
  affiliation: true,
  logo: true,
  courseDetails: true,
  assignmentTopic: true,
  submittedBy: true,
  submittedTo: true,
};

const THEME_PRESETS = [
  {
    name: "Classic Navy",
    primary: "#1e3a8a",
    text: "#0f172a",
    bg: "#ffffff",
    font: "'Times New Roman', Times, serif",
    borderStyle: "solid",
    borderWidth: 6,
  },
  {
    name: "Crimson Academic",
    primary: "#991b1b",
    text: "#1e1b4b",
    bg: "#fffbeb",
    font: "'Georgia', serif",
    borderStyle: "double",
    borderWidth: 8,
  },
  {
    name: "Minimalist Slate",
    primary: "#334155",
    text: "#1e293b",
    bg: "#ffffff",
    font: "'Arial', 'Helvetica', sans-serif",
    borderStyle: "none",
    borderWidth: 0,
  },
  {
    name: "Forest Scholar",
    primary: "#065f46",
    text: "#022c22",
    bg: "#fafdfb",
    font: "'Georgia', serif",
    borderStyle: "solid",
    borderWidth: 4,
  },
];

const InputField = ({
  label,
  name,
  value,
  onChange,
  textStyle,
  onStyleChange,
  inputType = "text",
}: {
  label: string;
  name: string;
  value: string;
  onChange: (n: string, v: string) => void;
  textStyle?: TextStyle;
  onStyleChange?: (n: string, s: TextStyle) => void;
  inputType?: "text" | "date";
}) => {
  const [showStyle, setShowStyle] = useState(false);

  return (
    <div className="mb-4">
      <div className="flex justify-between items-center mb-1">
        <label className="block text-xs font-semibold text-gray-700 uppercase tracking-wider">
          {label}
        </label>
        {textStyle && onStyleChange && (
          <button
            type="button"
            onClick={() => setShowStyle(!showStyle)}
            className={`p-1 rounded transition-colors ${showStyle ? "bg-blue-100 text-blue-600" : "text-gray-400 hover:text-gray-600 hover:bg-gray-100"}`}
            title="Text Styling Options"
          >
            <Type size={14} />
          </button>
        )}
      </div>
      {inputType === "text" ? (
        <textarea
          value={value}
          onChange={(e) => onChange(name, e.target.value)}
          className="w-full px-3 py-1.5 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm min-h-[38px] resize-y"
          rows={value.split("\n").length || 1}
        />
      ) : (
        <input
          type={inputType}
          value={value}
          onChange={(e) => onChange(name, e.target.value)}
          className="w-full px-3 py-1.5 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm min-h-[38px]"
        />
      )}
      {showStyle && textStyle && onStyleChange && (
        <div className="mt-2 p-3 bg-gray-50 border border-gray-200 rounded-md grid grid-cols-2 gap-3 shadow-inner">
          <div>
            <label className="block text-[10px] text-gray-500 mb-1">
              Size (px)
            </label>
            <input
              type="number"
              value={textStyle.fontSize}
              onChange={(e) =>
                onStyleChange(name, {
                  ...textStyle,
                  fontSize: Number(e.target.value),
                })
              }
              className="w-full px-2 py-1 border border-gray-300 rounded text-sm"
            />
          </div>
          <div>
            <label className="block text-[10px] text-gray-500 mb-1">Font</label>
            <select
              value={textStyle.fontFamily}
              onChange={(e) =>
                onStyleChange(name, {
                  ...textStyle,
                  fontFamily: e.target.value,
                })
              }
              className="w-full px-2 py-1 border border-gray-300 rounded text-sm"
            >
              <option value="inherit">Default</option>
              <option value="'Arial', sans-serif">Arial</option>
              <option value="'Times New Roman', serif">Times New Roman</option>
              <option value="'Courier New', monospace">Courier</option>
              <option value="'Georgia', serif">Georgia</option>
              <option value="'Verdana', sans-serif">Verdana</option>
            </select>
          </div>
          <div className="col-span-2 flex gap-2">
            <button
              type="button"
              onClick={() =>
                onStyleChange(name, { ...textStyle, isBold: !textStyle.isBold })
              }
              className={`flex-1 py-1 px-2 border rounded text-sm font-bold transition-colors ${textStyle.isBold ? "bg-blue-100 border-blue-300 text-blue-700" : "bg-white border-gray-300"}`}
            >
              B
            </button>
            <button
              type="button"
              onClick={() =>
                onStyleChange(name, {
                  ...textStyle,
                  isItalic: !textStyle.isItalic,
                })
              }
              className={`flex-1 py-1 px-2 border rounded text-sm italic transition-colors ${textStyle.isItalic ? "bg-blue-100 border-blue-300 text-blue-700" : "bg-white border-gray-300"}`}
            >
              I
            </button>
          </div>
          <div className="col-span-2 grid grid-cols-2 gap-2">
            <div>
              <label className="block text-[10px] text-gray-500 mb-1">
                Alignment
              </label>
              <div className="flex border rounded border-gray-300 overflow-hidden h-7">
                <button
                  type="button"
                  onClick={() =>
                    onStyleChange(name, { ...textStyle, textAlign: "left" })
                  }
                  className={`flex-1 flex items-center justify-center transition-colors ${textStyle.textAlign === "left" ? "bg-blue-100 text-blue-700" : "bg-white hover:bg-gray-50 text-gray-600"}`}
                >
                  <AlignLeft size={14} />
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onStyleChange(name, { ...textStyle, textAlign: "center" })
                  }
                  className={`flex-1 border-l border-r border-gray-300 flex items-center justify-center transition-colors ${textStyle.textAlign === "center" ? "bg-blue-100 text-blue-700" : "bg-white hover:bg-gray-50 text-gray-600"}`}
                >
                  <AlignCenter size={14} />
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onStyleChange(name, { ...textStyle, textAlign: "right" })
                  }
                  className={`flex-1 flex items-center justify-center transition-colors ${textStyle.textAlign === "right" ? "bg-blue-100 text-blue-700" : "bg-white hover:bg-gray-50 text-gray-600"}`}
                >
                  <AlignRight size={14} />
                </button>
              </div>
            </div>
            <div>
              <label className="block text-[10px] text-gray-500 mb-1">
                Casing
              </label>
              <select
                value={textStyle.textTransform || "none"}
                onChange={(e) =>
                  onStyleChange(name, {
                    ...textStyle,
                    textTransform: e.target.value as any,
                  })
                }
                className="w-full px-2 h-7 border border-gray-300 rounded text-xs"
              >
                <option value="none">Normal</option>
                <option value="uppercase">UPPERCASE</option>
                <option value="lowercase">lowercase</option>
                <option value="capitalize">Capitalize</option>
              </select>
            </div>
          </div>
          <div className="col-span-2 grid grid-cols-4 gap-2 mt-2">
            <div>
              <label className="block text-[9px] text-gray-500 mb-1 leading-tight">
                Margin Top
              </label>
              <input
                type="number"
                value={textStyle.marginTop ?? ""}
                onChange={(e) =>
                  onStyleChange(name, {
                    ...textStyle,
                    marginTop: e.target.value
                      ? Number(e.target.value)
                      : undefined,
                  })
                }
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs text-center"
              />
            </div>
            <div>
              <label className="block text-[9px] text-gray-500 mb-1 leading-tight">
                Margin Btm
              </label>
              <input
                type="number"
                value={textStyle.marginBottom ?? ""}
                onChange={(e) =>
                  onStyleChange(name, {
                    ...textStyle,
                    marginBottom: e.target.value
                      ? Number(e.target.value)
                      : undefined,
                  })
                }
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs text-center"
              />
            </div>
            <div>
              <label className="block text-[9px] text-gray-500 mb-1 leading-tight">
                Pad Top
              </label>
              <input
                type="number"
                value={textStyle.paddingTop ?? ""}
                onChange={(e) =>
                  onStyleChange(name, {
                    ...textStyle,
                    paddingTop: e.target.value
                      ? Number(e.target.value)
                      : undefined,
                  })
                }
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs text-center"
              />
            </div>
            <div>
              <label className="block text-[9px] text-gray-500 mb-1 leading-tight">
                Pad Btm
              </label>
              <input
                type="number"
                value={textStyle.paddingBottom ?? ""}
                onChange={(e) =>
                  onStyleChange(name, {
                    ...textStyle,
                    paddingBottom: e.target.value
                      ? Number(e.target.value)
                      : undefined,
                  })
                }
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs text-center"
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [activeTab, setActiveTab] = useState<"content" | "styling" | "layout">(
    "content",
  );

  // Load from local storage or fallback to defaults
  const [data, setData] = useState(() => {
    const saved = localStorage.getItem("cover_builder_data");
    return saved ? JSON.parse(saved) : DEFAULT_DATA;
  });

  const [textStyles, setTextStyles] = useState<Record<string, TextStyle>>(
    () => {
      const saved = localStorage.getItem("cover_builder_text_styles");
      return saved ? JSON.parse(saved) : DEFAULT_TEXT_STYLES;
    },
  );

  const [globalStyles, setGlobalStyles] = useState(() => {
    const saved = localStorage.getItem("cover_builder_global_styles");
    return saved ? JSON.parse(saved) : DEFAULT_GLOBAL_STYLES;
  });

  const [visibleSections, setVisibleSections] = useState(() => {
    const saved = localStorage.getItem("cover_builder_visible_sections");
    return saved ? JSON.parse(saved) : DEFAULT_VISIBLE_SECTIONS;
  });

  const [layoutKey, setLayoutKey] = useState(0);
  const [isExporting, setIsExporting] = useState(false);
  const [scale, setScale] = useState(1);

  const previewContainerRef = useRef<HTMLDivElement>(null);
  const pageRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Sync state to local storage automatically
  useEffect(() => {
    localStorage.setItem("cover_builder_data", JSON.stringify(data));
  }, [data]);

  useEffect(() => {
    localStorage.setItem(
      "cover_builder_text_styles",
      JSON.stringify(textStyles),
    );
  }, [textStyles]);

  useEffect(() => {
    localStorage.setItem(
      "cover_builder_global_styles",
      JSON.stringify(globalStyles),
    );
  }, [globalStyles]);

  useEffect(() => {
    localStorage.setItem(
      "cover_builder_visible_sections",
      JSON.stringify(visibleSections),
    );
  }, [visibleSections]);

  // Dynamic Scaling of A4 Preview Page
  useEffect(() => {
    const updateScale = () => {
      if (previewContainerRef.current) {
        const container = previewContainerRef.current;
        const padding = 64;
        const availableWidth = container.clientWidth - padding;
        const availableHeight = container.clientHeight - padding;

        const a4Width = 794;
        const a4Height = 1123;

        const scaleX = availableWidth / a4Width;
        const scaleY = availableHeight / a4Height;

        setScale(Math.min(scaleX, scaleY, 1.4));
      }
    };

    updateScale();
    window.addEventListener("resize", updateScale);
    return () => window.removeEventListener("resize", updateScale);
  }, []);

  const handleDataChange = useCallback((name: string, value: string) => {
    setData((prev: typeof DEFAULT_DATA) => ({ ...prev, [name]: value }));
  }, []);

  const handleTextStyleChange = useCallback(
    (name: string, style: TextStyle) => {
      setTextStyles((prev) => ({ ...prev, [name]: style }));
    },
    [],
  );

  const handleGlobalStyleChange = (
    e: React.ChangeEvent<HTMLSelectElement | HTMLInputElement>,
  ) => {
    const { name, value } = e.target;
    setGlobalStyles((prev) => ({
      ...prev,
      [name]:
        name === "borderWidth" || name === "padding" ? Number(value) : value,
    }));
  };

  const toggleSection = (section: keyof typeof DEFAULT_VISIBLE_SECTIONS) => {
    setVisibleSections((prev) => ({ ...prev, [section]: !prev[section] }));
  };

  const applyPresetTheme = (preset: (typeof THEME_PRESETS)[0]) => {
    setGlobalStyles((prev) => ({
      ...prev,
      fontFamily: preset.font,
      textColor: preset.text,
      borderColor: preset.primary,
      backgroundColor: preset.bg,
      borderStyle: preset.borderStyle,
      borderWidth: preset.borderWidth,
    }));
  };

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setData((prev: typeof DEFAULT_DATA) => ({
          ...prev,
          logoUrl: reader.result as string,
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const removeLogo = () => {
    setData((prev: typeof DEFAULT_DATA) => ({ ...prev, logoUrl: "" }));
  };

  const handleReset = () => {
    if (
      confirm("Reset layout data, colors, styles, and visibility to defaults?")
    ) {
      setData(DEFAULT_DATA);
      setTextStyles(DEFAULT_TEXT_STYLES);
      setGlobalStyles(DEFAULT_GLOBAL_STYLES);
      setVisibleSections(DEFAULT_VISIBLE_SECTIONS);
      setLayoutKey((prev) => prev + 1);
    }
  };

  // Export current configuration state as JSON file
  const handleExportConfig = () => {
    const configData = {
      version: 1,
      data,
      textStyles,
      globalStyles,
      visibleSections,
    };
    const blob = new Blob([JSON.stringify(configData, null, 2)], {
      type: "application/json",
    });
    saveAs(blob, "cover-page-config.json");
  };

  // Import configuration state from JSON file
  const handleImportConfig = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target?.result as string);
        if (parsed.data && parsed.textStyles && parsed.globalStyles) {
          setData(parsed.data);
          setTextStyles(parsed.textStyles);
          setGlobalStyles(parsed.globalStyles);
          if (parsed.visibleSections) {
            setVisibleSections(parsed.visibleSections);
          }
          alert("Configuration loaded successfully.");
        } else {
          alert("Invalid configuration structure.");
        }
      } catch (err) {
        alert("Failed to parse configuration file.");
      }
    };
    reader.readAsText(file);
  };

  const handlePrint = () => {
    window.print();
  };

  const getElement = async () => {
    const element = document.getElementById("cover-page");
    if (!element) return null;
    await new Promise((resolve) => setTimeout(resolve, 100));
    return element;
  };

  const handleExportImage = async () => {
    setIsExporting(true);
    try {
      const element = await getElement();
      if (element) {
        const dataUrl = await toPng(element, { pixelRatio: 2 });
        const link = document.createElement("a");
        link.download = "cover-page.png";
        link.href = dataUrl;
        link.click();
      }
    } catch (err) {
      console.error(err);
      alert("Could not export image.");
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportPDF = async () => {
    setIsExporting(true);
    try {
      const element = await getElement();
      if (element) {
        const dataUrl = await toJpeg(element, { quality: 0.95, pixelRatio: 2 });
        const pdf = new jsPDF({
          orientation: "portrait",
          unit: "mm",
          format: "a4",
        });
        pdf.addImage(dataUrl, "JPEG", 0, 0, 210, 297);
        pdf.save("cover-page.pdf");
      }
    } catch (err) {
      console.error(err);
      alert("Could not export PDF.");
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportWord = async () => {
    setIsExporting(true);
    try {
      const element = await getElement();
      if (element) {
        const dataUrl = await toJpeg(element, { quality: 0.95, pixelRatio: 2 });
        const base64Data = dataUrl.split(",")[1];

        const doc = new Document({
          sections: [
            {
              properties: {
                page: {
                  margin: { top: 0, right: 0, bottom: 0, left: 0 },
                  size: { width: 11906, height: 16838 },
                },
              },
              children: [
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: base64ToUint8Array(base64Data),
                      transformation: {
                        width: 794,
                        height: 1123,
                      },
                      type: "jpg",
                    }),
                  ],
                }),
              ],
            },
          ],
        });

        Packer.toBlob(doc).then((blob) => {
          saveAs(blob, "cover-page.docx");
        });
      }
    } catch (err) {
      console.error(err);
    } finally {
      setIsExporting(false);
    }
  };

  const getStyleObj = (style: TextStyle) => ({
    fontSize: `${style.fontSize}px`,
    fontWeight: style.isBold ? "bold" : "normal",
    fontStyle: style.isItalic ? "italic" : "normal",
    fontFamily: style.fontFamily !== "inherit" ? style.fontFamily : "inherit",
    textAlign: style.textAlign || ("center" as any),
    marginTop: style.marginTop ? `${style.marginTop}px` : undefined,
    marginBottom: style.marginBottom ? `${style.marginBottom}px` : undefined,
    paddingTop: style.paddingTop ? `${style.paddingTop}px` : undefined,
    paddingBottom: style.paddingBottom ? `${style.paddingBottom}px` : undefined,
    textTransform:
      style.textTransform && style.textTransform !== "none"
        ? style.textTransform
        : undefined,
    whiteSpace: "pre-wrap" as any,
  });

  return (
    <div className="flex h-screen bg-slate-200 overflow-hidden font-sans antialiased text-slate-800">
      {/* Sidebar - Control Panel */}
      <div className="w-[440px] bg-white border-r border-slate-200 flex flex-col h-full z-10 no-print flex-shrink-0 shadow-lg">
        {/* App Title & Quick Tools */}
        <div className="p-4 bg-slate-50 border-b border-slate-200">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <Settings2 className="text-blue-600" size={20} />
              <h1 className="text-lg font-bold text-slate-900 tracking-tight">
                Cover Page Builder
              </h1>
            </div>
            <button
              onClick={handleReset}
              className="text-xs font-semibold text-red-600 hover:text-red-800 flex items-center gap-1 bg-red-50 hover:bg-red-100 px-2 py-1 rounded transition-colors"
              title="Reset configuration to defaults"
            >
              <RotateCcw size={12} /> Reset All
            </button>
          </div>

          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={handleExportConfig}
              className="flex items-center justify-center gap-1.5 text-xs bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 py-1.5 px-2 rounded font-medium transition-colors"
              title="Export layout configuration"
            >
              <Download size={13} /> Save Config
            </button>
            <label
              className="flex items-center justify-center gap-1.5 text-xs bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 py-1.5 px-2 rounded font-medium cursor-pointer transition-colors"
              title="Import layout configuration"
            >
              <FolderOpen size={13} /> Load Config
              <input
                type="file"
                accept=".json"
                onChange={handleImportConfig}
                className="hidden"
              />
            </label>
          </div>
        </div>

        {/* Dynamic Navigation Tabs */}
        <div className="flex border-b border-slate-200 bg-slate-50/50">
          <button
            onClick={() => setActiveTab("content")}
            className={`flex-1 py-3 text-xs font-bold uppercase tracking-wider flex items-center justify-center gap-2 border-b-2 transition-all ${activeTab === "content" ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-slate-500 hover:text-slate-800"}`}
          >
            <Layers size={14} /> Content
          </button>
          <button
            onClick={() => setActiveTab("styling")}
            className={`flex-1 py-3 text-xs font-bold uppercase tracking-wider flex items-center justify-center gap-2 border-b-2 transition-all ${activeTab === "styling" ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-slate-500 hover:text-slate-800"}`}
          >
            <Sparkles size={14} /> Design & Themes
          </button>
          <button
            onClick={() => setActiveTab("layout")}
            className={`flex-1 py-3 text-xs font-bold uppercase tracking-wider flex items-center justify-center gap-2 border-b-2 transition-all ${activeTab === "layout" ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-slate-500 hover:text-slate-800"}`}
          >
            <Eye size={14} /> Visibility
          </button>
        </div>

        {/* Tab Contents */}
        <div className="flex-1 overflow-y-auto p-5 space-y-6">
          {/* Tab 1: Content Setup */}
          {activeTab === "content" && (
            <div className="space-y-6 animate-fadeIn">
              <div className="bg-blue-50 text-blue-900 p-3 rounded-lg border border-blue-200 text-xs flex items-start gap-2">
                <Move
                  className="mt-0.5 flex-shrink-0 text-blue-600"
                  size={14}
                />
                <p>
                  You can click and drag elements directly inside the A4 sheet
                  preview to fine-tune spacing and vertical layout positioning.
                </p>
              </div>

              {/* Institution Details */}
              <div className="border border-slate-100 rounded-lg p-4 bg-slate-50/50 space-y-4">
                <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                  Institution
                </h2>
                <InputField
                  label="College / University Name"
                  name="collegeName"
                  value={data.collegeName}
                  onChange={handleDataChange}
                  textStyle={textStyles.collegeName}
                  onStyleChange={handleTextStyleChange}
                />
                {visibleSections.address && (
                  <InputField
                    label="Address"
                    name="address"
                    value={data.address}
                    onChange={handleDataChange}
                    textStyle={textStyles.address}
                    onStyleChange={handleTextStyleChange}
                  />
                )}
                {visibleSections.affiliation && (
                  <InputField
                    label="Affiliation"
                    name="affiliation"
                    value={data.affiliation}
                    onChange={handleDataChange}
                    textStyle={textStyles.affiliation}
                    onStyleChange={handleTextStyleChange}
                  />
                )}
              </div>

              {/* Course Info */}
              {visibleSections.courseDetails && (
                <div className="border border-slate-100 rounded-lg p-4 bg-slate-50/50 space-y-4">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                    Course / Program
                  </h2>
                  <InputField
                    label="Program & Semester"
                    name="program"
                    value={data.program}
                    onChange={handleDataChange}
                    textStyle={textStyles.program}
                    onStyleChange={handleTextStyleChange}
                  />
                  <InputField
                    label="Subject Name"
                    name="subject"
                    value={data.subject}
                    onChange={handleDataChange}
                    textStyle={textStyles.subject}
                    onStyleChange={handleTextStyleChange}
                  />
                </div>
              )}

              {/* Logo Upload Panel */}
              {visibleSections.logo && (
                <div className="border border-slate-100 rounded-lg p-4 bg-slate-50/50">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
                    Logo
                  </h2>
                  <div className="flex flex-col gap-2">
                    <label className="cursor-pointer bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 font-semibold py-2 px-4 rounded-md shadow-sm text-xs flex items-center justify-center gap-2 transition-colors">
                      <Upload size={14} />
                      <span>Upload New Logo</span>
                      <input
                        type="file"
                        className="hidden"
                        accept="image/*"
                        onChange={handleImageUpload}
                      />
                    </label>
                    {data.logoUrl && (
                      <button
                        onClick={removeLogo}
                        className="text-red-500 hover:text-red-700 text-xs font-semibold self-center mt-1"
                      >
                        Remove Current Logo
                      </button>
                    )}
                  </div>
                </div>
              )}

              {/* Assignment Topic */}
              {visibleSections.assignmentTopic && (
                <div className="border border-slate-100 rounded-lg p-4 bg-slate-50/50 space-y-4">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                    Topic
                  </h2>
                  <InputField
                    label="Assignment / Work Type"
                    name="assignmentType"
                    value={data.assignmentType}
                    onChange={handleDataChange}
                    textStyle={textStyles.assignmentType}
                    onStyleChange={handleTextStyleChange}
                  />
                  <InputField
                    label="Topic / Title"
                    name="topic"
                    value={data.topic}
                    onChange={handleDataChange}
                    textStyle={textStyles.topic}
                    onStyleChange={handleTextStyleChange}
                  />
                </div>
              )}

              {/* Submitter Metadata */}
              {visibleSections.submittedBy && (
                <div className="border border-slate-100 rounded-lg p-4 bg-slate-50/50 space-y-4">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                    Submitted By
                  </h2>
                  <InputField
                    label="Section Title"
                    name="submittedByTitle"
                    value={data.submittedByTitle}
                    onChange={handleDataChange}
                    textStyle={textStyles.submittedByTitle}
                    onStyleChange={handleTextStyleChange}
                  />
                  <div className="grid grid-cols-2 gap-2">
                    <InputField
                      label="Name Label"
                      name="submitterNameLabel"
                      value={data.submitterNameLabel}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterNameLabel}
                      onStyleChange={handleTextStyleChange}
                    />
                    <InputField
                      label="Name"
                      name="submitterName"
                      value={data.submitterName}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterName}
                      onStyleChange={handleTextStyleChange}
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <InputField
                      label="ID / Roll Label"
                      name="submitterRollLabel"
                      value={data.submitterRollLabel}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterRollLabel}
                      onStyleChange={handleTextStyleChange}
                    />
                    <InputField
                      label="ID / Roll Value"
                      name="submitterRoll"
                      value={data.submitterRoll}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterRoll}
                      onStyleChange={handleTextStyleChange}
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <InputField
                      label="Semester Label"
                      name="submitterSemLabel"
                      value={data.submitterSemLabel}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterSemLabel}
                      onStyleChange={handleTextStyleChange}
                    />
                    <InputField
                      label="Semester Value"
                      name="submitterSem"
                      value={data.submitterSem}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterSem}
                      onStyleChange={handleTextStyleChange}
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <InputField
                      label="Date Label"
                      name="submitterDateLabel"
                      value={data.submitterDateLabel}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterDateLabel}
                      onStyleChange={handleTextStyleChange}
                    />
                    <InputField
                      label="Date Value"
                      name="submitterDate"
                      value={data.submitterDate}
                      onChange={handleDataChange}
                      textStyle={textStyles.submitterDate}
                      onStyleChange={handleTextStyleChange}
                      inputType="date"
                    />
                  </div>
                </div>
              )}

              {/* Submitter Destination */}
              {visibleSections.submittedTo && (
                <div className="border border-slate-100 rounded-lg p-4 bg-slate-50/50 space-y-4">
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">
                    Submitted To
                  </h2>
                  <InputField
                    label="Section Title"
                    name="submittedToTitle"
                    value={data.submittedToTitle}
                    onChange={handleDataChange}
                    textStyle={textStyles.submittedToTitle}
                    onStyleChange={handleTextStyleChange}
                  />
                  <InputField
                    label="Instructor / Authority Name"
                    name="submittedToName"
                    value={data.submittedToName}
                    onChange={handleDataChange}
                    textStyle={textStyles.submittedToName}
                    onStyleChange={handleTextStyleChange}
                  />
                </div>
              )}
            </div>
          )}

          {/* Tab 2: Design and Theme Settings */}
          {activeTab === "styling" && (
            <div className="space-y-6 animate-fadeIn">
              {/* Presets Grid */}
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3 flex items-center gap-1">
                  <Palette size={14} className="text-blue-500" /> Style Presets
                </h3>
                <div className="grid grid-cols-2 gap-2">
                  {THEME_PRESETS.map((preset) => (
                    <button
                      key={preset.name}
                      onClick={() => applyPresetTheme(preset)}
                      className="border border-slate-200 hover:border-blue-500 bg-white p-3 rounded-lg text-left shadow-sm hover:shadow transition-all group flex flex-col justify-between h-20"
                    >
                      <span className="text-xs font-bold text-slate-800 leading-tight group-hover:text-blue-600 transition-colors">
                        {preset.name}
                      </span>
                      <div className="flex gap-1.5 items-center mt-1">
                        <span
                          className="w-4 h-4 rounded-full border border-slate-200 block"
                          style={{ backgroundColor: preset.primary }}
                        />
                        <span
                          className="w-4 h-4 rounded-full border border-slate-200 block"
                          style={{ backgroundColor: preset.text }}
                        />
                        <span
                          className="w-4 h-4 rounded-full border border-slate-200 block"
                          style={{ backgroundColor: preset.bg }}
                        />
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Custom Global Parameters */}
              <div className="border-t border-slate-100 pt-6 space-y-4">
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
                  Fine-tune Settings
                </h3>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">
                    Global Base Font
                  </label>
                  <select
                    name="fontFamily"
                    value={globalStyles.fontFamily}
                    onChange={handleGlobalStyleChange}
                    className="w-full px-3 py-2 border border-slate-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                  >
                    <option value="'Arial', 'Helvetica', sans-serif">
                      Arial
                    </option>
                    <option value="'Times New Roman', Times, serif">
                      Times New Roman
                    </option>
                    <option value="'Courier New', Courier, monospace">
                      Courier New
                    </option>
                    <option value="'Georgia', serif">Georgia</option>
                    <option value="'Verdana', Geneva, sans-serif">
                      Verdana
                    </option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Text Base Color
                    </label>
                    <div className="flex items-center gap-2 border border-slate-300 p-1.5 rounded-md">
                      <input
                        type="color"
                        name="textColor"
                        value={globalStyles.textColor}
                        onChange={handleGlobalStyleChange}
                        className="h-7 w-7 rounded cursor-pointer border-0"
                      />
                      <span className="text-xs text-slate-500 font-mono">
                        {globalStyles.textColor}
                      </span>
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Paper Color
                    </label>
                    <div className="flex items-center gap-2 border border-slate-300 p-1.5 rounded-md">
                      <input
                        type="color"
                        name="backgroundColor"
                        value={globalStyles.backgroundColor}
                        onChange={handleGlobalStyleChange}
                        className="h-7 w-7 rounded cursor-pointer border-0"
                      />
                      <span className="text-xs text-slate-500 font-mono">
                        {globalStyles.backgroundColor}
                      </span>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Border Color
                    </label>
                    <div className="flex items-center gap-2 border border-slate-300 p-1.5 rounded-md">
                      <input
                        type="color"
                        name="borderColor"
                        value={globalStyles.borderColor}
                        onChange={handleGlobalStyleChange}
                        className="h-7 w-7 rounded cursor-pointer border-0"
                      />
                      <span className="text-xs text-slate-500 font-mono">
                        {globalStyles.borderColor}
                      </span>
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Border Pattern
                    </label>
                    <select
                      name="borderStyle"
                      value={globalStyles.borderStyle}
                      onChange={handleGlobalStyleChange}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    >
                      <option value="solid">Solid</option>
                      <option value="dashed">Dashed</option>
                      <option value="dotted">Dotted</option>
                      <option value="double">Double</option>
                      <option value="groove">Groove</option>
                      <option value="none">None</option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Border Width (px)
                    </label>
                    <input
                      type="number"
                      name="borderWidth"
                      value={globalStyles.borderWidth}
                      onChange={handleGlobalStyleChange}
                      className="w-full px-3 py-1.5 border border-slate-300 rounded-md text-sm shadow-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Page Margins (px)
                    </label>
                    <input
                      type="number"
                      name="padding"
                      value={globalStyles.padding}
                      onChange={handleGlobalStyleChange}
                      className="w-full px-3 py-1.5 border border-slate-300 rounded-md text-sm shadow-sm"
                    />
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Tab 3: Layout Configuration & Field Visibility */}
          {activeTab === "layout" && (
            <div className="space-y-4 animate-fadeIn">
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
                Toggle Component Visibility
              </h3>
              <p className="text-xs text-slate-500 mb-4">
                Disable optional sections easily if your university format does
                not require them.
              </p>

              <div className="divide-y divide-slate-100 border border-slate-200 rounded-lg overflow-hidden bg-white">
                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Address Details
                    </span>
                    <span className="text-[10px] text-slate-500">
                      Displays address below college name
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("address")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.address ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.address ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>

                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Affiliation Statement
                    </span>
                    <span className="text-[10px] text-slate-500">
                      Displays university affiliation
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("affiliation")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.affiliation ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.affiliation ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>

                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Institution Logo
                    </span>
                    <span className="text-[10px] text-slate-500">
                      Shows uploaded image or a placeholder
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("logo")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.logo ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.logo ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>

                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Course & Program
                    </span>
                    <span className="text-[10px] text-slate-500">
                      Shows Degree and Subject names
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("courseDetails")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.courseDetails ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.courseDetails ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>

                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Assignment Topic Block
                    </span>
                    <span className="text-[10px] text-slate-500">
                      The core topic block and underline
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("assignmentTopic")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.assignmentTopic ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.assignmentTopic ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>

                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Submitted By Panel
                    </span>
                    <span className="text-[10px] text-slate-500">
                      Left metadata section (Name, Roll, Sem)
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("submittedBy")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.submittedBy ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.submittedBy ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>

                <div className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-800">
                      Submitted To Panel
                    </span>
                    <span className="text-[10px] text-slate-500">
                      Right metadata section (Instructor/Teacher)
                    </span>
                  </div>
                  <button
                    onClick={() => toggleSection("submittedTo")}
                    className={`p-1.5 rounded-md transition-all ${visibleSections.submittedTo ? "bg-blue-50 text-blue-600 hover:bg-blue-100" : "bg-slate-100 text-slate-400 hover:bg-slate-200"}`}
                  >
                    {visibleSections.submittedTo ? (
                      <Eye size={16} />
                    ) : (
                      <EyeOff size={16} />
                    )}
                  </button>
                </div>
              </div>

              <button
                onClick={() => setLayoutKey((prev) => prev + 1)}
                className="mt-4 w-full flex items-center justify-center py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold uppercase tracking-wider rounded transition-colors"
              >
                Reset Draggable Coordinates
              </button>
            </div>
          )}
        </div>

        {/* Actions Footer */}
        <div className="p-4 border-t border-slate-200 bg-slate-50 flex flex-col gap-3">
          <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-0.5">
            Document Outputs
          </div>
          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={handleExportPDF}
              disabled={isExporting}
              className="flex items-center justify-center gap-2 bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 py-2 px-3 rounded-md text-xs font-semibold transition-colors disabled:opacity-50"
            >
              <FileText size={14} /> PDF File
            </button>
            <button
              onClick={handleExportWord}
              disabled={isExporting}
              className="flex items-center justify-center gap-2 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 py-2 px-3 rounded-md text-xs font-semibold transition-colors disabled:opacity-50"
            >
              <FileText size={14} /> Word Docx
            </button>
            <button
              onClick={handleExportImage}
              disabled={isExporting}
              className="col-span-2 flex items-center justify-center gap-2 bg-green-50 hover:bg-green-100 text-green-700 border border-green-200 py-2.5 px-3 rounded-md text-xs font-semibold transition-colors disabled:opacity-50"
            >
              <FileImage size={14} /> Raster PNG Image
            </button>
          </div>
          <button
            onClick={handlePrint}
            disabled={isExporting}
            className="w-full flex items-center justify-center gap-2 bg-slate-900 hover:bg-slate-950 text-white font-bold text-sm py-3 px-4 rounded-md shadow transition-all disabled:opacity-50"
          >
            <Printer size={16} /> Print Direct
          </button>
        </div>
      </div>

      {/* Preview Workspace Area */}
      <div
        ref={previewContainerRef}
        className="flex-1 bg-slate-300 relative flex justify-center items-center overflow-hidden"
      >
        <div
          className="transition-transform duration-200 origin-center"
          style={{ transform: `scale(${scale})` }}
        >
          {/* A4 Page Layout Wrapper (Strictly 794x1123 @ 96 DPI) */}
          <div
            id="cover-page"
            className="shadow-2xl relative flex flex-col box-border overflow-hidden select-none"
            style={{
              width: "794px",
              height: "1123px",
              padding: `${globalStyles.padding}px`,
              fontFamily: globalStyles.fontFamily,
              color: globalStyles.textColor,
              backgroundColor: globalStyles.backgroundColor,
              border:
                globalStyles.borderWidth > 0
                  ? `${globalStyles.borderWidth}px ${globalStyles.borderStyle} ${globalStyles.borderColor}`
                  : "none",
            }}
          >
            <div
              key={layoutKey}
              className="flex-1 flex flex-col h-full w-full justify-between"
              ref={pageRef}
            >
              {/* College Header */}
              <motion.div
                drag
                dragConstraints={pageRef}
                dragMomentum={false}
                className={`flex flex-col p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-slate-100/60 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                style={{
                  textAlign: textStyles.collegeName?.textAlign || "center",
                }}
              >
                <h1
                  style={getStyleObj(textStyles.collegeName)}
                  className="mb-1 leading-tight"
                >
                  {data.collegeName || (
                    <span className="inline-block h-[0.8em] w-3/4 bg-slate-200/60 rounded no-print"></span>
                  )}
                </h1>
                {visibleSections.address && (
                  <p
                    style={getStyleObj(textStyles.address)}
                    className="mb-1 leading-tight"
                  >
                    {data.address || (
                      <span className="inline-block h-[0.8em] w-1/2 bg-slate-200/60 rounded no-print"></span>
                    )}
                  </p>
                )}
                {visibleSections.affiliation && (
                  <p
                    style={getStyleObj(textStyles.affiliation)}
                    className="leading-tight"
                  >
                    {data.affiliation || (
                      <span className="inline-block h-[0.8em] w-1/2 bg-slate-200/60 rounded no-print"></span>
                    )}
                  </p>
                )}
              </motion.div>

              {/* Course Info */}
              {visibleSections.courseDetails && (
                <motion.div
                  drag
                  dragConstraints={pageRef}
                  dragMomentum={false}
                  className={`flex flex-col p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-slate-100/60 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                  style={{
                    textAlign: textStyles.program?.textAlign || "center",
                  }}
                >
                  <h2
                    style={getStyleObj(textStyles.program)}
                    className="mb-2 leading-tight"
                  >
                    {data.program || (
                      <span className="inline-block h-[0.8em] w-2/3 bg-slate-200/60 rounded no-print"></span>
                    )}
                  </h2>
                  <h1
                    style={getStyleObj(textStyles.subject)}
                    className="leading-tight px-4"
                  >
                    {data.subject || (
                      <span className="inline-block h-[0.8em] w-3/4 bg-slate-200/60 rounded no-print"></span>
                    )}
                  </h1>
                </motion.div>
              )}

              {/* Logo Area */}
              {visibleSections.logo && (
                <motion.div
                  drag
                  dragConstraints={pageRef}
                  dragMomentum={false}
                  className={`flex items-center justify-center min-h-[140px] max-h-[280px] p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-slate-100/60 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                >
                  {data.logoUrl ? (
                    <img
                      src={data.logoUrl}
                      alt="Logo"
                      className="max-w-full max-h-[220px] object-contain pointer-events-none"
                    />
                  ) : (
                    <div className="w-[160px] h-[160px] border-[3px] border-dashed border-slate-300 text-slate-400 flex flex-col items-center justify-center rounded-xl no-print">
                      <ImageIcon size={48} className="mb-2 opacity-50" />
                      <span className="text-xs text-center px-4">
                        Upload in sidebar
                      </span>
                    </div>
                  )}
                </motion.div>
              )}

              {/* Topic Information Block */}
              {visibleSections.assignmentTopic && (
                <motion.div
                  drag
                  dragConstraints={pageRef}
                  dragMomentum={false}
                  className={`flex flex-col p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-slate-100/60 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                  style={{
                    textAlign: textStyles.assignmentType?.textAlign || "center",
                  }}
                >
                  <h2
                    style={getStyleObj(textStyles.assignmentType)}
                    className="mb-1.5 leading-tight"
                  >
                    {data.assignmentType || (
                      <span className="inline-block h-[0.8em] w-1/2 bg-slate-200/60 rounded no-print"></span>
                    )}
                  </h2>
                  <div
                    className={`inline-block border-b-[3px] border-dotted border-[currentColor] px-10 pb-2 ${textStyles.assignmentType?.textAlign === "left" ? "self-start" : textStyles.assignmentType?.textAlign === "right" ? "self-end" : "self-center"}`}
                  >
                    <span
                      style={getStyleObj(textStyles.topic)}
                      className="leading-tight"
                    >
                      {data.topic || (
                        <span className="inline-block h-[0.8em] w-32 bg-slate-200/60 rounded no-print"></span>
                      )}
                    </span>
                  </div>
                </motion.div>
              )}

              {/* Submitted By / Submitted To Container Block */}
              {(visibleSections.submittedBy || visibleSections.submittedTo) && (
                <motion.div
                  drag
                  dragConstraints={pageRef}
                  dragMomentum={false}
                  className={`flex flex-row justify-between w-full relative mt-auto p-2 rounded transition-colors z-10 ${!isExporting ? "cursor-move hover:bg-slate-100/60 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                >
                  {/* Vertical separator guide shown only if both panels are enabled */}
                  {visibleSections.submittedBy &&
                    visibleSections.submittedTo && (
                      <div className="absolute left-1/2 top-0 bottom-0 w-[2px] bg-[currentColor] opacity-30 -translate-x-1/2"></div>
                    )}

                  {/* Submitted By - Left Subpanel */}
                  {visibleSections.submittedBy ? (
                    <div
                      className={`${visibleSections.submittedTo ? "w-[46%]" : "w-full"} flex flex-col pr-4`}
                    >
                      <h3
                        style={getStyleObj(textStyles.submittedByTitle)}
                        className="border-b-[4px] border-[currentColor] inline-block pb-1.5 mb-5 self-start min-w-[140px]"
                      >
                        {data.submittedByTitle || (
                          <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                        )}
                      </h3>

                      <div className="flex mb-4 items-end">
                        <span
                          style={getStyleObj(textStyles.submitterNameLabel)}
                          className="mr-3 whitespace-nowrap min-w-[50px]"
                        >
                          {data.submitterNameLabel || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </span>
                        <div
                          style={getStyleObj(textStyles.submitterName)}
                          className="flex-grow border-b-[2px] border-dotted border-[currentColor] text-center pb-0.5 min-h-[1.8em]"
                        >
                          {data.submitterName || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </div>
                      </div>

                      <div className="flex mb-4 items-end">
                        <span
                          style={getStyleObj(textStyles.submitterRollLabel)}
                          className="mr-3 whitespace-nowrap min-w-[50px]"
                        >
                          {data.submitterRollLabel || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </span>
                        <div
                          style={getStyleObj(textStyles.submitterRoll)}
                          className="flex-grow border-b-[2px] border-dotted border-[currentColor] text-center pb-0.5 min-h-[1.8em]"
                        >
                          {data.submitterRoll || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </div>
                      </div>

                      <div className="flex mb-4 items-end">
                        <span
                          style={getStyleObj(textStyles.submitterSemLabel)}
                          className="mr-3 whitespace-nowrap min-w-[50px]"
                        >
                          {data.submitterSemLabel || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </span>
                        <div
                          style={getStyleObj(textStyles.submitterSem)}
                          className="flex-grow border-b-[2px] border-dotted border-[currentColor] pb-0.5 min-h-[1.8em] text-center"
                        >
                          {data.submitterSem || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </div>
                      </div>

                      <div className="flex mb-4 items-end">
                        <span
                          style={getStyleObj(textStyles.submitterDateLabel)}
                          className="mr-3 whitespace-nowrap min-w-[50px]"
                        >
                          {data.submitterDateLabel || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </span>
                        <div
                          style={getStyleObj(textStyles.submitterDate)}
                          className="flex-grow border-b-[2px] border-dotted border-[currentColor] pb-0.5 min-h-[1.8em] text-center"
                        >
                          {data.submitterDate || (
                            <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                          )}
                        </div>
                      </div>
                    </div>
                  ) : null}

                  {/* Submitted To - Right Subpanel */}
                  {visibleSections.submittedTo ? (
                    <div
                      className={`${visibleSections.submittedBy ? "w-[46%]" : "w-full"} flex flex-col pl-4`}
                    >
                      <h3
                        style={getStyleObj(textStyles.submittedToTitle)}
                        className="border-b-[4px] border-[currentColor] inline-block pb-1.5 mb-5 self-end min-w-[140px] text-right"
                      >
                        {data.submittedToTitle || (
                          <span className="inline-block h-[0.8em] w-full bg-slate-200/60 rounded no-print"></span>
                        )}
                      </h3>

                      <div className="w-full flex flex-col mt-auto mb-4">
                        <div className="w-full border-b-[2px] border-dotted border-[currentColor] mb-3"></div>
                        <div
                          style={getStyleObj(textStyles.submittedToName)}
                          className="w-full text-center min-h-[1.8em]"
                        >
                          {data.submittedToName || (
                            <span className="inline-block h-[0.8em] w-3/4 bg-slate-200/60 rounded no-print"></span>
                          )}
                        </div>
                      </div>
                    </div>
                  ) : null}
                </motion.div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
