"use client";

import React, { useState } from "react";
import { motion } from "motion/react";
import {
  Upload,
  FileSpreadsheet,
  Download,
  X,
  Plus,
  FileCheck,
  Search,
  Settings,
  CheckCircle2,
  Loader2,
  HelpCircle,
  ChevronDown,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import * as XLSX from "xlsx";
import { AuroraBackground } from "@/components/ui/aurora-background";

type FieldConfig = {
  name: string;
  type: "manual" | "cell" | "formula";
  cellAddress?: string;
  formula?: string;
};

const MASTER_STRUCTURE: FieldConfig[] = [
  { name: "COD. CLIENTE", type: "manual" },
  {
    name: "VALOR FOB. DI",
    type: "formula",
    formula: "=SOMA([@[VALOR USD CLIENTE]]*[@[TX  PREVIA CLIENTE]])",
  },
  { name: "STATUS DI", type: "manual" },
  { name: "VALOR USD CLIENTE", type: "cell", cellAddress: "H24" },
  { name: "TX  PREVIA CLIENTE", type: "cell", cellAddress: "K8" },
  { name: "FECHAMENTO BCO", type: "cell", cellAddress: "H24" },
  { name: "TX", type: "manual" },
  {
    name: "VALOR USD EFETIVO",
    type: "formula",
    formula: "=SOMA([@TX]*[@[FECHAMENTO BCO]])",
  },
  { name: "CLIENTE", type: "manual" },
  { name: "STATUS SWIFT", type: "manual" },
  { name: "STATUS FINANCEIRO", type: "manual" },
  {
    name: "SALDO",
    type: "formula",
    formula: "=SOMA([@[VALOR FOB. DI]]-[@[VALOR USD EFETIVO]])",
  },
  { name: "OBSERVAÇÃO", type: "manual" },
  { name: "DESP. OPE. ENVI. NEEMAN", type: "manual" },
  { name: "SISCOMEX", type: "cell", cellAddress: "H26" },
  { name: "MARINHA MERCANTE", type: "cell", cellAddress: "H27" },
  { name: "OUTRAS DESP. ADUAN.", type: "cell", cellAddress: "H28" },
  { name: "IMPOSTO IMPOR. (I.I)", type: "cell", cellAddress: "K32" },
  { name: "PROG. INTE. SOC. (PIS)", type: "cell", cellAddress: "K33" },
  { name: "CONT. FINAN. SOC. COFINS", type: "cell", cellAddress: "K34" },
  { name: "IMP. PROD. IMP. (IPI)", type: "cell", cellAddress: "K35" },
  { name: "DUMPING", type: "cell", cellAddress: "K36" },
  { name: "IMP. CIRC. MERC. (ICMS)", type: "cell", cellAddress: "K37" },
  { name: "ARMAZ. ZONA PRIM.", type: "cell", cellAddress: "H41" },
  { name: "DIF. FRETE INTER", type: "cell", cellAddress: "H42" },
  { name: "DESPACHANTE", type: "cell", cellAddress: "H43" },
  { name: "CONSULTA ADM. / LI", type: "cell", cellAddress: "H44" },
  { name: "TAXA BL", type: "cell", cellAddress: "H45" },
  { name: "TARIFA CAMBIAL", type: "cell", cellAddress: "H46" },
  { name: "ESCOLTA", type: "cell", cellAddress: "H47" },
  { name: "TAXA EXPEDIENTE", type: "cell", cellAddress: "H48" },
  { name: "FRETE AO CLIENTE", type: "cell", cellAddress: "H49" },
  { name: "RETIFICAÇÃO D.I", type: "manual" },
  { name: "JANELA ESPECIAL", type: "manual" },
  { name: "CREDITO PIS", type: "cell", cellAddress: "K57" },
  { name: "CREDITO COFINS", type: "cell", cellAddress: "K58" },
  { name: "CREDITO IPI", type: "cell", cellAddress: "K59" },
  { name: "CREDITO ICMS", type: "cell", cellAddress: "K60" },
  { name: "CUSTO DO ESTOQUE", type: "cell", cellAddress: "K61" },
  { name: "CUSTO TRADING", type: "cell", cellAddress: "K62" },
  { name: "DEBITO PIS", type: "cell", cellAddress: "K63" },
  { name: "DEBITO COFINS", type: "cell", cellAddress: "K64" },
  { name: "DEBITO ICMS", type: "cell", cellAddress: "K65" },
  { name: "VALOR SEM IPI", type: "cell", cellAddress: "K67" },
  { name: "IPI DESTACADO", type: "cell", cellAddress: "K68" },
  {
    name: "PGTO EFETIVO NEEMAN",
    type: "formula",
    formula: "=SOMA([@[SISCOMEX]:[@[JANELA ESPECIAL]]])",
  },
  { name: "STATUS NEEMAN", type: "manual" },
  { name: "DESP. FRETE", type: "cell", cellAddress: "H49" },
  { name: "DESP. Escolta", type: "cell", cellAddress: "H47" },
  { name: "IMPOSTOS  IR", type: "cell", cellAddress: "K81" },
  { name: "IMPOSTO CSLL", type: "cell", cellAddress: "K82" },
  {
    name: "RESUMO",
    type: "formula",
    formula: "=[@[VALOR USD EFETIVO]]+[@SALDO]+[@[PGTO EFETIVO NEEMAN]]",
  },
  { name: "NF EMITIDA", type: "manual" },
  { name: "APROVAÇÃO CLIENTE", type: "cell", cellAddress: "K76" },
  {
    name: "CUSTO CLIENTE",
    type: "formula",
    formula: "=SOMA([@[APROVAÇÃO CLIENTE]]-[@[VALOR USD EFETIVO]])",
  },
  {
    name: "RESULTADO",
    type: "formula",
    formula: "=SOMA([@[APROVAÇÃO CLIENTE]]-[@RESUMO])",
  },
  {
    name: "CC NEMAN",
    type: "formula",
    formula: "=[@[DESP. OPE. ENVI. NEEMAN]]-[@[PGTO EFETIVO NEEMAN]]",
  },
  { name: "STATUS GERAL", type: "manual" },
  { name: "DESTINO", type: "manual" },
  { name: "OBSERVAÇÃO2", type: "manual" },
  { name: "DATA FINAL", type: "manual" },
];

type DataRow = Record<string, any>;
type CellMapping = { [fieldName: string]: string };
type FileMapping = {
  file: File;
  workbook: XLSX.WorkBook;
  mapping: CellMapping;
};

export default function Page() {
  const [baseFile, setBaseFile] = useState<File | null>(null);
  const [baseData, setBaseData] = useState<DataRow[]>([]);
  const [extraFiles, setExtraFiles] = useState<FileMapping[]>([]);
  const [consolidatedData, setConsolidatedData] = useState<DataRow[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showMapping, setShowMapping] = useState(false);
  const [mappingSearchFilter, setMappingSearchFilter] = useState("");
  const [selectedCategory, setSelectedCategory] = useState<string>("all");
  const [showHelp, setShowHelp] = useState(false);

  const [globalMapping, setGlobalMapping] = useState<CellMapping>(() => {
    const initialMapping: CellMapping = {};
    MASTER_STRUCTURE.forEach((field) => {
      if (field.type === "cell" && field.cellAddress) {
        initialMapping[field.name] = field.cellAddress;
      } else {
        initialMapping[field.name] = "";
      }
    });
    return initialMapping;
  });

  const baseHeaders = MASTER_STRUCTURE.map((field) => field.name);

  const readBaseFile = async (file: File): Promise<DataRow[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: "binary" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet) as DataRow[];

          const normalizedData = json.map((row) => {
            const normalizedRow: DataRow = {};
            MASTER_STRUCTURE.forEach((field) => {
              const fieldKey = Object.keys(row).find(
                (key) =>
                  key.trim().toUpperCase() === field.name.trim().toUpperCase()
              );
              normalizedRow[field.name] = fieldKey ? row[fieldKey] : "";
            });
            return normalizedRow;
          });

          resolve(normalizedData);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    });
  };

  const readExtraFile = async (file: File): Promise<XLSX.WorkBook> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: "binary" });
          resolve(workbook);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    });
  };

  const handleBaseFileUpload = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      setBaseFile(file);
      setConsolidatedData([]);

      try {
        const data = await readBaseFile(file);
        setBaseData(data);
      } catch (error) {
        console.error("Erro ao ler arquivo base:", error);
        alert("Erro ao ler arquivo base");
      }
    }
  };

  const handleExtraFilesUpload = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files);

      for (const file of newFiles) {
        try {
          const workbook = await readExtraFile(file);
          const fileMapping: FileMapping = {
            file,
            workbook,
            mapping: {},
          };
          setExtraFiles((prev) => [...prev, fileMapping]);
        } catch (error) {
          console.error("Erro ao ler arquivo:", file.name, error);
          alert(`Erro ao ler arquivo ${file.name}`);
        }
      }
    }
  };

  const removeExtraFile = (index: number) => {
    setExtraFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const removeBaseFile = () => {
    setBaseFile(null);
    setBaseData([]);
    setConsolidatedData([]);
    setExtraFiles([]);
  };

  const getCellValue = (workbook: XLSX.WorkBook, cellAddress: string): any => {
    try {
      if (!cellAddress || cellAddress.trim() === "") return "";

      const sheetName =
        workbook.SheetNames.find(
          (name) =>
            name.toLowerCase().includes("entrada") &&
            name.toLowerCase().includes("saida")
        ) || workbook.SheetNames[0];

      const worksheet = workbook.Sheets[sheetName];
      const cell = worksheet[cellAddress];
      return cell ? cell.v : "";
    } catch (error) {
      return "";
    }
  };

  const convertFormulaToExcelFormat = (
    formula: string,
    rowIndex: number
  ): string => {
    let excelFormula = formula.trim();

    const rangePattern = /\[@\[([^\]]+)\]:\[@?\[?([^\]]+)\]?\]?\]/g;
    excelFormula = excelFormula.replace(
      rangePattern,
      (match, startField, endField) => {
        const startIndex = MASTER_STRUCTURE.findIndex(
          (f) => f.name === startField
        );
        const endIndex = MASTER_STRUCTURE.findIndex((f) => f.name === endField);

        if (startIndex !== -1 && endIndex !== -1) {
          const startCol = XLSX.utils.encode_col(startIndex);
          const endCol = XLSX.utils.encode_col(endIndex);
          const cellRange = `${startCol}${rowIndex + 1}:${endCol}${
            rowIndex + 1
          }`;
          return cellRange;
        }

        return match;
      }
    );

    MASTER_STRUCTURE.forEach((field, colIndex) => {
      const excelCol = XLSX.utils.encode_col(colIndex);
      const excelCell = `${excelCol}${rowIndex + 1}`;

      const pattern1 = new RegExp(
        `\\[@\\[${field.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\]\\]`,
        "g"
      );
      excelFormula = excelFormula.replace(pattern1, excelCell);

      const pattern2 = new RegExp(
        `\\[@${field.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\]`,
        "g"
      );
      excelFormula = excelFormula.replace(pattern2, excelCell);
    });

    excelFormula = excelFormula.replace(/Tabela1/g, "");
    excelFormula = excelFormula.replace(/SOMA\s*\(/gi, "");

    if (excelFormula.trim().endsWith(")")) {
      excelFormula = excelFormula.trim().slice(0, -1);
    }

    return excelFormula;
  };

  const consolidatePlanilhas = async () => {
    if (!baseFile || extraFiles.length === 0) return;

    setIsProcessing(true);
    setConsolidatedData([]);

    await new Promise((resolve) => setTimeout(resolve, 100));

    try {
      const allData = [...baseData];

      extraFiles.forEach((fileMapping) => {
        const newRow: DataRow = {};

        MASTER_STRUCTURE.forEach((field) => {
          if (field.type === "formula") {
            newRow[field.name] = "__FORMULA__";
          } else {
            const cellAddress = globalMapping[field.name];
            if (cellAddress && cellAddress.trim() !== "") {
              newRow[field.name] = getCellValue(
                fileMapping.workbook,
                cellAddress
              );
            } else {
              newRow[field.name] = "";
            }
          }
        });

        allData.push(newRow);
      });

      setConsolidatedData(allData);
    } catch (error) {
      console.error("Erro ao consolidar planilhas:", error);
      alert(
        "Erro ao processar as planilhas. Verifique se os endereços de células estão corretos."
      );
    } finally {
      setIsProcessing(false);
    }
  };

  const handleNewConsolidation = () => {
    setBaseFile(null);
    setBaseData([]);
    setExtraFiles([]);
    setConsolidatedData([]);
    setIsProcessing(false);
  };

  const downloadConsolidatedFile = () => {
    if (consolidatedData.length === 0) return;

    const orderedData = consolidatedData.map((row) => {
      const orderedRow: DataRow = {};
      MASTER_STRUCTURE.forEach((field) => {
        orderedRow[field.name] = row[field.name] ?? "";
      });
      return orderedRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(orderedData, {
      header: baseHeaders,
    });

    orderedData.forEach((row, rowIndex) => {
      MASTER_STRUCTURE.forEach((field, colIndex) => {
        if (field.type === "formula" && field.formula) {
          const excelCol = XLSX.utils.encode_col(colIndex);
          const cellAddress = `${excelCol}${rowIndex + 2}`;
          const excelFormula = convertFormulaToExcelFormat(
            field.formula,
            rowIndex + 1
          );

          worksheet[cellAddress] = {
            f: excelFormula.replace(/^=/, ""),
          };
        }
      });
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dados Consolidados");

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "planilha_consolidada.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const updateGlobalMapping = (fieldName: string, value: string) => {
    setGlobalMapping((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const mappedFieldsCount = Object.values(globalMapping).filter(
    (v) => v && v.trim() !== ""
  ).length;
  const formulaFieldsCount = MASTER_STRUCTURE.filter(
    (f) => f.type === "formula"
  ).length;
  const totalFields = MASTER_STRUCTURE.length;

  const categories = [
    { id: "all", label: "Todos", count: MASTER_STRUCTURE.length },
    { id: "mapped", label: "Mapeados", count: mappedFieldsCount },
    { id: "formula", label: "Fórmulas", count: formulaFieldsCount },
    {
      id: "manual",
      label: "Manuais",
      count: MASTER_STRUCTURE.filter((f) => f.type === "manual").length,
    },
  ];

  const filteredFields = MASTER_STRUCTURE.filter((field) => {
    const matchesSearch = field.name
      .toLowerCase()
      .includes(mappingSearchFilter.toLowerCase());

    if (selectedCategory === "all") return matchesSearch;
    if (selectedCategory === "mapped")
      return matchesSearch && globalMapping[field.name];
    if (selectedCategory === "formula")
      return matchesSearch && field.type === "formula";
    if (selectedCategory === "manual")
      return matchesSearch && field.type === "manual";

    return matchesSearch;
  });

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="mx-auto max-w-7xl px-4 py-8">
        {/* Header */}
        <div className="mb-12 text-center">
          <h1 className="text-5xl font-bold text-gray-900 mb-3">
            Consolidador de Planilhas
          </h1>
          <p className="text-lg text-gray-600 max-w-2xl mx-auto">
            Automatize a consolidação de dados com extração inteligente
          </p>
        </div>

        {/* Mapping Panel - Full Screen Overlay */}
        {showMapping && (
          <div className="fixed inset-0 bg-gray-50 z-50 overflow-y-auto">
            <div className="min-h-screen">
              <div className="mx-auto max-w-7xl px-4 py-8">
                <Card className="overflow-hidden border border-gray-200">
                  <div className="bg-gray-900 p-6 text-white">
                    <div className="flex items-center justify-between">
                      <div>
                        <h2 className="text-2xl font-bold mb-1">
                          Configuração de Mapeamento
                        </h2>
                        <p className="text-gray-300">
                          Defina os endereços de célula para extração automática
                        </p>
                      </div>
                      <Button
                        onClick={() => setShowMapping(false)}
                        variant="ghost"
                        size="sm"
                        className="text-white hover:bg-gray-800"
                      >
                        <X className="h-5 w-5" />
                      </Button>
                    </div>
                  </div>

                  <div className="p-6">
                    {/* Search and Filter */}
                    <div className="mb-6 space-y-4">
                      <div className="relative">
                        <Search className="absolute left-4 top-1/2 transform -translate-y-1/2 h-5 w-5 text-gray-400" />
                        <Input
                          type="text"
                          placeholder="Buscar campo..."
                          value={mappingSearchFilter}
                          onChange={(e) =>
                            setMappingSearchFilter(e.target.value)
                          }
                          className="pl-12 h-12 border-gray-300"
                        />
                      </div>

                      <div className="flex gap-2">
                        {categories.map((cat) => (
                          <button
                            key={cat.id}
                            onClick={() => setSelectedCategory(cat.id)}
                            className={`px-4 py-2 rounded-lg text-sm font-medium transition-all ${
                              selectedCategory === cat.id
                                ? "bg-gray-900 text-white"
                                : "bg-gray-100 text-gray-700 hover:bg-gray-200"
                            }`}
                          >
                            {cat.label}
                            <span className="ml-2 px-1.5 py-0.5 rounded bg-white/20 text-xs">
                              {cat.count}
                            </span>
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* Fields Grid */}
                    <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3 max-h-[500px] overflow-y-auto pr-2 custom-scrollbar">
                      {filteredFields.map((field) => (
                        <div
                          key={field.name}
                          className={`p-4 rounded-lg border transition-all ${
                            field.type === "formula"
                              ? "bg-gray-50 border-gray-300"
                              : globalMapping[field.name]
                              ? "bg-green-50 border-green-200"
                              : "bg-white border-gray-200 hover:border-gray-300"
                          }`}
                        >
                          <div className="flex items-center justify-between mb-2">
                            <span className="font-medium text-sm text-gray-900 truncate">
                              {field.name}
                            </span>
                            {field.type === "formula" && (
                              <span className="px-2 py-0.5 bg-gray-200 text-gray-700 rounded text-xs font-medium shrink-0">
                                Auto
                              </span>
                            )}
                            {globalMapping[field.name] &&
                              field.type !== "formula" && (
                                <CheckCircle2 className="h-4 w-4 text-green-600 shrink-0" />
                              )}
                          </div>
                          {field.type === "formula" ? (
                            <Input
                              type="text"
                              value={field.formula || ""}
                              disabled
                              className="bg-gray-100 border-gray-300 text-gray-600 cursor-not-allowed text-xs"
                            />
                          ) : (
                            <Input
                              type="text"
                              placeholder="Ex: H24"
                              value={globalMapping[field.name] || ""}
                              onChange={(e) =>
                                updateGlobalMapping(field.name, e.target.value)
                              }
                              className="border-gray-300"
                            />
                          )}
                        </div>
                      ))}
                    </div>

                    {filteredFields.length === 0 && (
                      <div className="text-center py-12 text-gray-400">
                        <Search className="h-12 w-12 mx-auto mb-3 opacity-50" />
                        <p>Nenhum campo encontrado</p>
                      </div>
                    )}
                  </div>
                </Card>
              </div>
            </div>
          </div>
        )}

        {/* Main Content - Hidden when mapping is open */}
        {!showMapping && (
          <>
            {/* Upload Section */}
            {!consolidatedData.length && !isProcessing && (
              <div className="grid gap-6 lg:grid-cols-2 mb-8">
                {/* Base File */}
                <Card className="overflow-hidden border border-gray-200">
                  <div className="bg-gray-900 p-4 text-white">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet className="h-6 w-6" />
                        <div>
                          <h2 className="text-lg font-bold">Planilha Master</h2>
                          <p className="text-sm text-gray-300">
                            Estrutura base de dados
                          </p>
                        </div>
                      </div>
                      <Button
                        onClick={() => setShowMapping(!showMapping)}
                        size="sm"
                        variant={showMapping ? "default" : "outline"}
                        className="gap-2 bg-white/10 hover:bg-white/20 border-white/20 text-white"
                      >
                        <Settings className="h-4 w-4" />
                        {showMapping ? "Fechar" : "Configurar"}
                        <span className="ml-1 px-2 py-0.5 bg-white/20 rounded-full text-xs font-semibold">
                          {mappedFieldsCount}/{totalFields - formulaFieldsCount}
                        </span>
                      </Button>
                    </div>
                  </div>

                  <div className="p-6">
                    {!baseFile ? (
                      <label className="flex flex-col items-center justify-center w-full h-48 border-2 border-dashed border-gray-300 rounded-lg cursor-pointer bg-gray-50 hover:bg-gray-100 transition-all">
                        <div className="flex flex-col items-center justify-center gap-3">
                          <Upload className="h-8 w-8 text-gray-400" />
                          <div className="text-center">
                            <p className="text-base font-semibold text-gray-700 mb-1">
                              Clique ou arraste o arquivo
                            </p>
                            <p className="text-sm text-gray-500">
                              Excel (.xlsx, .xls) ou CSV
                            </p>
                          </div>
                        </div>
                        <input
                          type="file"
                          className="hidden"
                          accept=".xlsx,.xls,.csv"
                          onChange={handleBaseFileUpload}
                        />
                      </label>
                    ) : (
                      <div className="space-y-4">
                        <div className="flex items-center justify-between p-4 bg-green-50 rounded-lg border border-green-200">
                          <div className="flex items-center gap-3">
                            <FileCheck className="h-6 w-6 text-green-600" />
                            <div>
                              <p className="font-semibold text-gray-900 truncate max-w-[200px]">
                                {baseFile.name}
                              </p>
                              <p className="text-sm text-gray-600">
                                {(baseFile.size / 1024).toFixed(1)} KB
                              </p>
                            </div>
                          </div>
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={removeBaseFile}
                            className="hover:bg-red-50 hover:text-red-600"
                          >
                            <X className="h-5 w-5" />
                          </Button>
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div className="p-3 bg-gray-50 rounded-lg border border-gray-200">
                            <div className="text-2xl font-bold text-gray-900">
                              {baseHeaders.length}
                            </div>
                            <div className="text-xs text-gray-600">Campos</div>
                          </div>
                          <div className="p-3 bg-gray-50 rounded-lg border border-gray-200">
                            <div className="text-2xl font-bold text-gray-900">
                              {baseData.length}
                            </div>
                            <div className="text-xs text-gray-600">
                              Registros
                            </div>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </Card>

                {/* Extra Files */}
                <Card className="overflow-hidden border border-gray-200">
                  <div className="bg-gray-900 p-4 text-white">
                    <div className="flex items-center gap-3">
                      <Plus className="h-6 w-6" />
                      <div>
                        <h2 className="text-lg font-bold">Planilhas Avulsas</h2>
                        <p className="text-sm text-gray-300">
                          Adicione múltiplos arquivos
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="p-6">
                    <label
                      className={`flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg transition-all mb-4 ${
                        !baseFile
                          ? "border-gray-200 bg-gray-50 cursor-not-allowed opacity-50"
                          : "border-gray-300 bg-gray-50 hover:bg-gray-100 cursor-pointer"
                      }`}
                    >
                      <div className="flex flex-col items-center justify-center gap-2">
                        <Plus
                          className={`h-6 w-6 ${
                            !baseFile ? "text-gray-300" : "text-gray-400"
                          }`}
                        />
                        <div className="text-center">
                          <p
                            className={`text-sm font-semibold ${
                              !baseFile ? "text-gray-400" : "text-gray-700"
                            }`}
                          >
                            {!baseFile
                              ? "Carregue a Master primeiro"
                              : "Adicionar planilhas"}
                          </p>
                        </div>
                      </div>
                      <input
                        type="file"
                        className="hidden"
                        accept=".xlsx,.xls,.csv"
                        multiple
                        onChange={handleExtraFilesUpload}
                        disabled={!baseFile}
                      />
                    </label>

                    {extraFiles.length > 0 && (
                      <div className="space-y-2 max-h-52 overflow-y-auto custom-scrollbar">
                        {extraFiles.map((fileMapping, index) => (
                          <div
                            key={index}
                            className="flex items-center justify-between p-3 bg-gray-50 rounded-lg border border-gray-200 hover:border-gray-300 transition-all"
                          >
                            <div className="flex items-center gap-3 min-w-0">
                              <FileSpreadsheet className="h-4 w-4 text-gray-600" />
                              <div className="min-w-0">
                                <p className="text-sm font-medium text-gray-900 truncate">
                                  {fileMapping.file.name}
                                </p>
                                <p className="text-xs text-gray-500">
                                  {(fileMapping.file.size / 1024).toFixed(1)} KB
                                </p>
                              </div>
                            </div>
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => removeExtraFile(index)}
                              className="hover:bg-red-50 hover:text-red-600 shrink-0"
                            >
                              <X className="h-4 w-4" />
                            </Button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </Card>
              </div>
            )}

            {/* Consolidate Button */}
            {baseFile &&
              extraFiles.length > 0 &&
              !consolidatedData.length &&
              !isProcessing && (
                <div className="flex justify-center mb-8">
                  <Button
                    onClick={consolidatePlanilhas}
                    size="lg"
                    className="gap-2 px-8 py-6 text-lg font-semibold bg-gray-900 hover:bg-gray-800"
                  >
                    Consolidar Dados
                  </Button>
                </div>
              )}

            {/* Processing */}
            {isProcessing && (
              <Card className="p-12 border border-gray-200">
                <div className="flex flex-col items-center justify-center gap-6">
                  <Loader2 className="h-16 w-16 animate-spin text-gray-900" />
                  <div className="text-center">
                    <h2 className="text-3xl font-bold text-gray-900 mb-2">
                      Consolidando dados...
                    </h2>
                    <p className="text-gray-600">
                      Processando {extraFiles.length} planilha(s)
                    </p>
                  </div>
                </div>
              </Card>
            )}

            {/* Success */}
            {consolidatedData.length > 0 && !isProcessing && (
              <Card className="p-12 border border-gray-200">
                <div className="flex flex-col items-center justify-center gap-6">
                  <div className="p-6 rounded-full bg-green-100">
                    <CheckCircle2 className="h-16 w-16 text-green-600" />
                  </div>
                  <div className="text-center">
                    <h2 className="text-4xl font-bold text-gray-900 mb-3">
                      Consolidação Concluída!
                    </h2>
                    <p className="text-lg text-gray-600 mb-2">
                      <span className="font-bold text-gray-900">
                        {consolidatedData.length}
                      </span>{" "}
                      registros processados com sucesso
                    </p>
                  </div>

                  <div className="flex gap-4">
                    <Button
                      onClick={downloadConsolidatedFile}
                      size="lg"
                      className="gap-2 px-8 py-6 text-lg font-semibold bg-gray-900 hover:bg-gray-800"
                    >
                      <Download className="h-6 w-6" />
                      Baixar Planilha
                    </Button>
                    <Button
                      onClick={handleNewConsolidation}
                      variant="outline"
                      size="lg"
                      className="gap-2 px-6 py-6 text-lg font-semibold"
                    >
                      Nova Consolidação
                    </Button>
                  </div>
                </div>
              </Card>
            )}
          </>
        )}
      </div>

      {/* Help Button */}
      <div className="fixed bottom-6 right-6">
        <div className="relative">
          {showHelp && (
            <div className="absolute bottom-16 right-0 w-80 bg-white rounded-lg shadow-2xl border border-gray-200 p-6 animate-in slide-in-from-bottom-2 fade-in">
              <div className="flex items-start justify-between mb-4">
                <h3 className="text-lg font-bold text-gray-900">Como usar</h3>
                <button
                  onClick={() => setShowHelp(false)}
                  className="text-gray-400 hover:text-gray-600"
                >
                  <X className="h-5 w-5" />
                </button>
              </div>
              <ol className="space-y-3 text-sm text-gray-600">
                <li className="flex gap-3">
                  <span className="flex-shrink-0 flex items-center justify-center w-6 h-6 rounded-full bg-gray-900 text-white text-xs font-bold">
                    1
                  </span>
                  <span>
                    Carregue a{" "}
                    <strong className="text-gray-900">planilha master</strong>{" "}
                    com a estrutura base
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="flex-shrink-0 flex items-center justify-center w-6 h-6 rounded-full bg-gray-900 text-white text-xs font-bold">
                    2
                  </span>
                  <span>
                    Adicione uma ou mais{" "}
                    <strong className="text-gray-900">planilhas avulsas</strong>{" "}
                    para extrair dados
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="flex-shrink-0 flex items-center justify-center w-6 h-6 rounded-full bg-gray-900 text-white text-xs font-bold">
                    3
                  </span>
                  <span>
                    Configure o{" "}
                    <strong className="text-gray-900">
                      mapeamento de células
                    </strong>{" "}
                    se necessário
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="flex-shrink-0 flex items-center justify-center w-6 h-6 rounded-full bg-gray-900 text-white text-xs font-bold">
                    4
                  </span>
                  <span>
                    Clique em{" "}
                    <strong className="text-gray-900">Consolidar</strong> e
                    baixe o resultado
                  </span>
                </li>
              </ol>
            </div>
          )}
          <button
            onClick={() => setShowHelp(!showHelp)}
            className="flex items-center justify-center w-14 h-14 rounded-full bg-gray-900 text-white shadow-lg hover:bg-gray-800 transition-all hover:scale-105"
          >
            <HelpCircle className="h-6 w-6" />
          </button>
        </div>
      </div>

      <style jsx>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 8px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: #f1f5f9;
          border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #cbd5e1;
          border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #94a3b8;
        }
      `}</style>
    </div>
  );
}
