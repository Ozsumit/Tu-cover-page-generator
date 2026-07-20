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
        <label className="block text-sm font-medium text-gray-700">
          {label}
        </label>
        {textStyle && onStyleChange && (
          <button
            type="button"
            onClick={() => setShowStyle(!showStyle)}
            className={`p-1 rounded ${showStyle ? "bg-blue-100 text-blue-600" : "text-gray-400 hover:text-gray-600 hover:bg-gray-100"}`}
            title="Text Styling"
          >
            <Type size={14} />
          </button>
        )}
      </div>
      {inputType === "text" ? (
        <textarea
          value={value}
          onChange={(e) => onChange(name, e.target.value)}
          className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 min-h-[42px] resize-y"
          rows={value.split("\n").length || 1}
        />
      ) : (
        <input
          type={inputType}
          value={value}
          onChange={(e) => onChange(name, e.target.value)}
          className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 min-h-[42px]"
        />
      )}
      {showStyle && textStyle && onStyleChange && (
        <div className="mt-2 p-3 bg-gray-50 border border-gray-200 rounded-md grid grid-cols-2 gap-3 shadow-inner">
          <div>
            <label className="block text-xs text-gray-500 mb-1">
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
            <label className="block text-xs text-gray-500 mb-1">Font</label>
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
              <label className="block text-[10px] text-gray-500 mb-1">
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
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs"
              />
            </div>
            <div>
              <label className="block text-[10px] text-gray-500 mb-1">
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
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs"
              />
            </div>
            <div>
              <label className="block text-[10px] text-gray-500 mb-1">
                Padding Top
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
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs"
              />
            </div>
            <div>
              <label className="block text-[10px] text-gray-500 mb-1">
                Padding Btm
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
                className="w-full px-1 py-1 border border-gray-300 rounded text-xs"
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [data, setData] = useState({
    collegeName: "Yeti International College",
    address: "New-Baneshwor, Kathmandu",
    affiliation: "Affiliated to Tribhuvan University",
    program: "",
    subject: "",
    assignmentType: "",
    topic: "",
    submittedByTitle: "Submitted By",
    submitterNameLabel: "Name:",
    submitterName: "",
    submitterRollLabel: "Symbol No.:",
    submitterRoll: "",
    submitterSemLabel: "Semester:",
    submitterSem: "",
    submitterDateLabel: "Date:",
    submitterDate: "",
    submittedToTitle: "Submitted To",
    submittedToName: "",
    logoUrl: "",
  });

  const [textStyles, setTextStyles] = useState<Record<string, TextStyle>>({
    collegeName: defaultStyle(34, true),
    address: defaultStyle(22, false),
    affiliation: defaultStyle(22, false),
    program: defaultStyle(26, true),
    subject: defaultStyle(38, true),
    assignmentType: defaultStyle(30, true),
    topic: defaultStyle(26, true, "center", "'Courier New', monospace"),
    submittedByTitle: defaultStyle(28, true),
    submitterNameLabel: defaultStyle(20, true),
    submitterName: defaultStyle(20, true, "center", "'Courier New', monospace"),
    submitterRollLabel: defaultStyle(20, true),
    submitterRoll: defaultStyle(20, true),
    submitterSemLabel: defaultStyle(20, true, "left"),
    submitterSem: defaultStyle(20, true, "center"),
    submitterDateLabel: defaultStyle(20, true, "left"),
    submitterDate: defaultStyle(20, true, "center", "'Courier New', monospace"),
    submittedToTitle: defaultStyle(28, true, "center"),
    submittedToName: defaultStyle(24, true),
  });

  const [globalStyles, setGlobalStyles] = useState({
    fontFamily: "'Arial', 'Helvetica', sans-serif",
    textColor: "#000000",
    borderColor: "#000000",
    borderWidth: 6,
    borderStyle: "solid",
    backgroundColor: "#ffffff",
    padding: 60,
  });

  const [layoutKey, setLayoutKey] = useState(0);

  const [isExporting, setIsExporting] = useState(false);
  const previewContainerRef = useRef<HTMLDivElement>(null);
  const pageRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);

  // Dynamic Scaling
  useEffect(() => {
    const updateScale = () => {
      if (previewContainerRef.current) {
        const container = previewContainerRef.current;
        const padding = 64; // padding around the A4 page
        const availableWidth = container.clientWidth - padding;
        const availableHeight = container.clientHeight - padding;

        // A4 size in pixels (at 96 DPI)
        const a4Width = 794;
        const a4Height = 1123;

        const scaleX = availableWidth / a4Width;
        const scaleY = availableHeight / a4Height;

        // Cap the max scale at 1.5x so it doesn't get ridiculously large on huge screens
        setScale(Math.min(scaleX, scaleY, 1.5));
      }
    };

    updateScale();
    window.addEventListener("resize", updateScale);
    return () => window.removeEventListener("resize", updateScale);
  }, []);

  const handleDataChange = useCallback((name: string, value: string) => {
    setData((prev) => ({ ...prev, [name]: value }));
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

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setData((prev) => ({ ...prev, logoUrl: reader.result as string }));
      };
      reader.readAsDataURL(file);
    }
  };

  const removeLogo = () => {
    setData((prev) => ({ ...prev, logoUrl: "" }));
  };

  const handlePrint = () => {
    window.print();
  };

  const getElement = async () => {
    const element = document.getElementById("cover-page");
    if (!element) return null;

    // Wait for any state-based UI changes (like removing borders) to apply to DOM
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
      console.error("Failed to export image", err);
      alert("Failed to export image. Please try again.");
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
      console.error("Failed to export PDF", err);
      alert("Failed to export PDF. Please try again.");
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
                  size: { width: 11906, height: 16838 }, // Twips equivalent to 210mm x 297mm
                },
              },
              children: [
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: base64ToUint8Array(base64Data),
                      transformation: {
                        width: 794, // 96 DPI width
                        height: 1123, // 96 DPI height
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
    <div className="flex h-screen bg-gray-100 overflow-hidden font-sans">
      {/* Sidebar Form - hidden when printing */}
      <div className="w-[420px] bg-white shadow-xl flex flex-col h-full z-10 no-print flex-shrink-0">
        <div className="p-5 border-b border-gray-200 bg-gray-50 flex items-center gap-3">
          <Settings2 className="text-blue-600" />
          <h1 className="text-xl font-bold text-gray-800">
            Cover Page Builder
          </h1>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-6">
          <div className="bg-blue-50 text-blue-800 p-3 rounded-md border border-blue-200 text-sm flex items-start gap-2 shadow-sm">
            <Move className="mt-0.5 flex-shrink-0" size={16} />
            <p>
              <strong>Interactive Layout:</strong> You can click and drag any
              section on the page to reposition it freely!
            </p>
          </div>

          <section className="bg-blue-50/50 p-4 rounded-lg border border-blue-100">
            <h2 className="text-sm font-semibold text-blue-900 flex items-center gap-2 mb-4">
              <Palette size={16} /> Global Styling & Colors
            </h2>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1 flex items-center gap-1">
                  Global Font Family
                </label>
                <select
                  name="fontFamily"
                  value={globalStyles.fontFamily}
                  onChange={handleGlobalStyleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
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
                  <option value="'Verdana', Geneva, sans-serif">Verdana</option>
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Text Color
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      name="textColor"
                      value={globalStyles.textColor}
                      onChange={handleGlobalStyleChange}
                      className="h-8 w-8 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-500">
                      {globalStyles.textColor}
                    </span>
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Background
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      name="backgroundColor"
                      value={globalStyles.backgroundColor}
                      onChange={handleGlobalStyleChange}
                      className="h-8 w-8 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-500">
                      {globalStyles.backgroundColor}
                    </span>
                  </div>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Border Color
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      name="borderColor"
                      value={globalStyles.borderColor}
                      onChange={handleGlobalStyleChange}
                      className="h-8 w-8 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-500">
                      {globalStyles.borderColor}
                    </span>
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Border Style
                  </label>
                  <select
                    name="borderStyle"
                    value={globalStyles.borderStyle}
                    onChange={handleGlobalStyleChange}
                    className="w-full px-3 py-1.5 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
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
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Border Width (px)
                  </label>
                  <input
                    type="number"
                    name="borderWidth"
                    value={globalStyles.borderWidth}
                    onChange={handleGlobalStyleChange}
                    className="w-full px-3 py-1.5 border border-gray-300 rounded-md text-sm"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Page Padding (px)
                  </label>
                  <input
                    type="number"
                    name="padding"
                    value={globalStyles.padding}
                    onChange={handleGlobalStyleChange}
                    className="w-full px-3 py-1.5 border border-gray-300 rounded-md text-sm"
                  />
                </div>
              </div>
            </div>
          </section>

          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b pb-2 mb-4">
              Header Information
            </h2>
            <InputField
              label="College / University Name"
              name="collegeName"
              value={data.collegeName}
              onChange={handleDataChange}
              textStyle={textStyles.collegeName}
              onStyleChange={handleTextStyleChange}
            />
            <InputField
              label="Address"
              name="address"
              value={data.address}
              onChange={handleDataChange}
              textStyle={textStyles.address}
              onStyleChange={handleTextStyleChange}
            />
            <InputField
              label="Affiliation"
              name="affiliation"
              value={data.affiliation}
              onChange={handleDataChange}
              textStyle={textStyles.affiliation}
              onStyleChange={handleTextStyleChange}
            />
          </section>

          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b pb-2 mb-4">
              Course Details
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
          </section>

          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b pb-2 mb-4">
              Logo
            </h2>
            <div className="flex flex-col gap-3">
              <label className="cursor-pointer bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 font-medium py-2 px-4 rounded-md shadow-sm flex items-center justify-center gap-2 transition-colors">
                <Upload size={18} />
                <span>Upload Logo Image</span>
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
                  className="text-red-500 hover:text-red-700 text-sm font-medium self-center"
                >
                  Remove Logo
                </button>
              )}
            </div>
          </section>

          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b pb-2 mb-4">
              Assignment Topic
            </h2>
            <InputField
              label="Assignment Type"
              name="assignmentType"
              value={data.assignmentType}
              onChange={handleDataChange}
              textStyle={textStyles.assignmentType}
              onStyleChange={handleTextStyleChange}
            />
            <InputField
              label="Topic"
              name="topic"
              value={data.topic}
              onChange={handleDataChange}
              textStyle={textStyles.topic}
              onStyleChange={handleTextStyleChange}
            />
          </section>

          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b pb-2 mb-4">
              Submitted By Section
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
                label="Name Value"
                name="submitterName"
                value={data.submitterName}
                onChange={handleDataChange}
                textStyle={textStyles.submitterName}
                onStyleChange={handleTextStyleChange}
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <InputField
                label="Roll Label"
                name="submitterRollLabel"
                value={data.submitterRollLabel}
                onChange={handleDataChange}
                textStyle={textStyles.submitterRollLabel}
                onStyleChange={handleTextStyleChange}
              />
              <InputField
                label="Roll Value"
                name="submitterRoll"
                value={data.submitterRoll}
                onChange={handleDataChange}
                textStyle={textStyles.submitterRoll}
                onStyleChange={handleTextStyleChange}
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <InputField
                label="Sem Label"
                name="submitterSemLabel"
                value={data.submitterSemLabel}
                onChange={handleDataChange}
                textStyle={textStyles.submitterSemLabel}
                onStyleChange={handleTextStyleChange}
              />
              <InputField
                label="Sem Value"
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
            <button
              onClick={() => setLayoutKey((prev) => prev + 1)}
              className="mt-4 w-full flex items-center justify-center py-2 px-3 bg-gray-100 hover:bg-gray-200 text-gray-700 text-sm font-medium rounded transition-colors"
            >
              Reset Layout Positions
            </button>
          </section>

          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b pb-2 mb-4">
              Submitted To Section
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
              label="Teacher Name"
              name="submittedToName"
              value={data.submittedToName}
              onChange={handleDataChange}
              textStyle={textStyles.submittedToName}
              onStyleChange={handleTextStyleChange}
            />
          </section>
        </div>

        {/* Actions Footer */}
        <div className="p-4 border-t border-gray-200 bg-gray-50 flex flex-col gap-3">
          <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-1">
            Export Options
          </div>
          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={handleExportPDF}
              disabled={isExporting}
              className="flex items-center justify-center gap-2 bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 py-2 px-3 rounded-md text-sm font-medium transition-colors disabled:opacity-50"
            >
              <FileText size={16} /> PDF
            </button>
            <button
              onClick={handleExportWord}
              disabled={isExporting}
              className="flex items-center justify-center gap-2 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 py-2 px-3 rounded-md text-sm font-medium transition-colors disabled:opacity-50"
            >
              <FileText size={16} /> Word
            </button>
            <button
              onClick={handleExportImage}
              disabled={isExporting}
              className="col-span-2 flex items-center justify-center gap-2 bg-green-50 hover:bg-green-100 text-green-700 border border-green-200 py-2 px-3 rounded-md text-sm font-medium transition-colors disabled:opacity-50"
            >
              <FileImage size={16} /> Image (PNG)
            </button>
          </div>
          <button
            onClick={handlePrint}
            disabled={isExporting}
            className="w-full flex items-center justify-center gap-2 bg-gray-900 hover:bg-black text-white font-medium py-2.5 px-4 rounded-md shadow-sm transition-colors mt-2 disabled:opacity-50"
          >
            <Printer size={18} /> Print directly
          </button>
        </div>
      </div>

      {/* Preview Area */}
      <div
        ref={previewContainerRef}
        className="flex-1 bg-gray-200 relative flex justify-center items-center overflow-hidden"
      >
        <div
          className="transition-transform duration-200 origin-center"
          style={{ transform: `scale(${scale})` }}
        >
          {/* A4 Paper Container - strictly 794x1123 px (210x297mm at 96 DPI) */}
          <div
            id="cover-page"
            className="shadow-2xl relative flex flex-col box-border overflow-hidden"
            style={{
              width: "794px",
              height: "1123px",
              padding: `${globalStyles.padding}px`,
              fontFamily: globalStyles.fontFamily,
              color: globalStyles.textColor,
              backgroundColor: globalStyles.backgroundColor,
              border: `${globalStyles.borderWidth}px ${globalStyles.borderStyle} ${globalStyles.borderColor}`,
            }}
          >
            <div
              key={layoutKey}
              className="flex-1 flex flex-col h-full w-full justify-between"
              ref={pageRef}
            >
              {/* Header */}
              <motion.div
                drag
                dragConstraints={pageRef}
                dragMomentum={false}
                className={`flex flex-col p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-gray-100/50 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                style={{
                  textAlign: textStyles.collegeName?.textAlign || "center",
                }}
              >
                <h1
                  style={getStyleObj(textStyles.collegeName)}
                  className="mb-2 leading-tight"
                >
                  {data.collegeName || (
                    <span className="inline-block h-[0.8em] w-3/4 bg-gray-200/60 rounded no-print"></span>
                  )}
                </h1>
                <p
                  style={getStyleObj(textStyles.address)}
                  className="mb-1 leading-tight"
                >
                  {data.address || (
                    <span className="inline-block h-[0.8em] w-1/2 bg-gray-200/60 rounded no-print"></span>
                  )}
                </p>
                <p
                  style={getStyleObj(textStyles.affiliation)}
                  className="leading-tight"
                >
                  {data.affiliation || (
                    <span className="inline-block h-[0.8em] w-1/2 bg-gray-200/60 rounded no-print"></span>
                  )}
                </p>
              </motion.div>

              {/* Course Info */}
              <motion.div
                drag
                dragConstraints={pageRef}
                dragMomentum={false}
                className={`flex flex-col p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-gray-100/50 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                style={{ textAlign: textStyles.program?.textAlign || "center" }}
              >
                <h2
                  style={getStyleObj(textStyles.program)}
                  className="mb-3 leading-tight"
                >
                  {data.program || (
                    <span className="inline-block h-[0.8em] w-2/3 bg-gray-200/60 rounded no-print"></span>
                  )}
                </h2>
                <h1
                  style={getStyleObj(textStyles.subject)}
                  className="leading-tight px-4"
                >
                  {data.subject || (
                    <span className="inline-block h-[0.8em] w-3/4 bg-gray-200/60 rounded no-print"></span>
                  )}
                </h1>
              </motion.div>

              {/* Logo */}
              <motion.div
                drag
                dragConstraints={pageRef}
                dragMomentum={false}
                className={`flex items-center justify-center min-h-[150px] max-h-[300px] p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-gray-100/50 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
              >
                {data.logoUrl ? (
                  <img
                    src={data.logoUrl}
                    alt="Logo"
                    className="max-w-full max-h-[250px] object-contain pointer-events-none"
                  />
                ) : (
                  <div className="w-[180px] h-[180px] border-[3px] border-dashed border-gray-300 text-gray-400 flex flex-col items-center justify-center rounded-xl no-print">
                    <ImageIcon size={64} className="mb-3 opacity-50" />
                    <span className="text-base text-center px-4">
                      Logo Placeholder
                      <br />
                      (Upload in sidebar)
                    </span>
                  </div>
                )}
              </motion.div>

              {/* Assignment Topic */}
              <motion.div
                drag
                dragConstraints={pageRef}
                dragMomentum={false}
                className={`flex flex-col p-2 rounded transition-colors relative z-10 ${!isExporting ? "cursor-move hover:bg-gray-100/50 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
                style={{
                  textAlign: textStyles.assignmentType?.textAlign || "center",
                }}
              >
                <h2
                  style={getStyleObj(textStyles.assignmentType)}
                  className="mb-2 leading-tight"
                >
                  {data.assignmentType || (
                    <span className="inline-block h-[0.8em] w-1/2 bg-gray-200/60 rounded no-print"></span>
                  )}
                </h2>
                <div
                  className={`inline-block border-b-[4px] border-dotted border-[currentColor] px-12 pb-2 ${textStyles.assignmentType?.textAlign === "left" ? "self-start" : textStyles.assignmentType?.textAlign === "right" ? "self-end" : "self-center"}`}
                >
                  <span
                    style={getStyleObj(textStyles.topic)}
                    className="leading-tight"
                  >
                    {data.topic || (
                      <span className="inline-block h-[0.8em] w-32 bg-gray-200/60 rounded no-print"></span>
                    )}
                  </span>
                </div>
              </motion.div>

              {/* Submitted By / To Section */}
              <motion.div
                drag
                dragConstraints={pageRef}
                dragMomentum={false}
                className={`flex flex-row justify-between w-full relative mt-auto p-2 rounded transition-colors z-10 ${!isExporting ? "cursor-move hover:bg-gray-100/50 hover:outline hover:outline-2 hover:outline-dashed hover:outline-blue-400" : ""}`}
              >
                {/* Vertical Divider */}
                <div className="absolute left-1/2 top-0 bottom-0 w-[5px] bg-[currentColor] -translate-x-1/2"></div>

                {/* Left Column - Submitted By */}
                <div className="w-[45%] flex flex-col pr-8">
                  <h3
                    style={getStyleObj(textStyles.submittedByTitle)}
                    className="border-b-[5px] border-[currentColor] inline-block pb-2 mb-6 self-start min-w-[150px]"
                  >
                    {data.submittedByTitle || (
                      <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                    )}
                  </h3>

                  <div className="flex mb-5 items-end">
                    <span
                      style={getStyleObj(textStyles.submitterNameLabel)}
                      className="mr-4 whitespace-nowrap min-w-[60px]"
                    >
                      {data.submitterNameLabel || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </span>
                    <div
                      style={getStyleObj(textStyles.submitterName)}
                      className="flex-grow border-b-[3px] border-dotted border-[currentColor] text-center pb-1 min-h-[2em]"
                    >
                      {data.submitterName || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </div>
                  </div>

                  <div className="flex mb-5 items-end">
                    <span
                      style={getStyleObj(textStyles.submitterRollLabel)}
                      className="mr-4 whitespace-nowrap min-w-[60px]"
                    >
                      {data.submitterRollLabel || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </span>
                    <div
                      style={getStyleObj(textStyles.submitterRoll)}
                      className="flex-grow border-b-[3px] border-dotted border-[currentColor] text-center pb-1 min-h-[2em]"
                    >
                      {data.submitterRoll || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </div>
                  </div>

                  <div className="flex mb-5 items-end">
                    <span
                      style={getStyleObj(textStyles.submitterSemLabel)}
                      className="mr-4 whitespace-nowrap min-w-[60px]"
                    >
                      {data.submitterSemLabel || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </span>
                    <div
                      style={getStyleObj(textStyles.submitterSem)}
                      className="flex-grow border-b-[3px] border-dotted border-[currentColor] pb-1 min-h-[2em] text-center"
                    >
                      {data.submitterSem || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </div>
                  </div>

                  <div className="flex mb-5 items-end">
                    <span
                      style={getStyleObj(textStyles.submitterDateLabel)}
                      className="mr-4 whitespace-nowrap min-w-[60px]"
                    >
                      {data.submitterDateLabel || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </span>
                    <div
                      style={getStyleObj(textStyles.submitterDate)}
                      className="flex-grow border-b-[3px] border-dotted border-[currentColor] pb-1 min-h-[2em] text-center"
                    >
                      {data.submitterDate || (
                        <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                      )}
                    </div>
                  </div>
                </div>

                {/* Right Column - Submitted To */}
                <div className="w-[45%] flex flex-col pl-8">
                  <h3
                    style={getStyleObj(textStyles.submittedToTitle)}
                    className="border-b-[5px] border-[currentColor] inline-block pb-2 mb-6 self-end min-w-[150px] text-right"
                  >
                    {data.submittedToTitle || (
                      <span className="inline-block h-[0.8em] w-full bg-gray-200/60 rounded no-print"></span>
                    )}
                  </h3>

                  <div className="w-full flex flex-col mt-auto mb-4">
                    <div className="w-full border-b-[3px] border-dotted border-[currentColor] mb-3"></div>
                    <div
                      style={getStyleObj(textStyles.submittedToName)}
                      className="w-full text-center min-h-[2em]"
                    >
                      {data.submittedToName || (
                        <span className="inline-block h-[0.8em] w-3/4 bg-gray-200/60 rounded no-print"></span>
                      )}
                    </div>
                  </div>
                </div>
              </motion.div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
