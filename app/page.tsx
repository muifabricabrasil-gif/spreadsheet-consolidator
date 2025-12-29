"use client";

import React, { useState, useRef } from "react";
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
import * as XLSX from "xlsx-js-style";
import { AuroraBackground } from "@/components/ui/aurora-background";

type CellCondition = {
  priority: number;
  type: "cell" | "formula" | "manual";
  cellAddress?: string;
  formula?: string;
};

type FieldConfig = {
  name: string;
  type: "manual" | "cell" | "formula" | "flex";
  cellAddress?: string;
  formula?: string;
  fallbackToManual?: boolean; // Se true, quando não houver valor na célula, deixa vazio (manual)
  cellCondition?: CellCondition[]; // Array de condições processadas por prioridade
};

const MASTER_STRUCTURE: FieldConfig[] = [
  // SEÇÃO - VALOR ADUNANEIRO - DI (índices 0-4)
  {
    name: "COD. CLIENTE",
    type: "flex",
    cellCondition: [
      { priority: 1, type: "cell", cellAddress: "H7" },
      { priority: 2, type: "manual" },
    ],
  },
  {
    name: "VALOR FOB. DI",
    type: "formula",
    formula: "=SOMA([@[VALOR USD CLIENTE]]*[@[TX  PREVIA CLIENTE]])",
  },
  { name: "STATUS DI", type: "manual" },
  { name: "VALOR USD CLIENTE", type: "cell", cellAddress: "H24" },
  { name: "TX  PREVIA CLIENTE", type: "cell", cellAddress: "K8" },

  // SEÇÃO - FECHAMENTO CAMBIO D.I - MARCHA (índices 5-12)
  { name: "FECHAMENTO BCO", type: "cell", cellAddress: "H24" },
  { name: "TX", type: "cell", cellAddress: "K10" },
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

  // SEÇÃO - PGTO ANDIAN. NACIONALIZAÇÃO (índice 13)
  {
    name: "DESP. OPE. ENVI. NEEMAN",
    type: "flex",
    cellCondition: [
      { priority: 1, type: "cell", cellAddress: "J88" },
      { priority: 2, type: "manual" },
    ],
  },

  // SEÇÃO - DESPESAS ADUNEIRAS (índices 14-16)
  { name: "SISCOMEX", type: "cell", cellAddress: "H26" },
  { name: "MARINHA MERCANTE", type: "cell", cellAddress: "H27" },
  { name: "OUTRAS DESP. ADUAN.", type: "cell", cellAddress: "H28" },

  // SEÇÃO - IMPOSTO (índices 17-22)
  { name: "IMPOSTO IMPOR. (I.I)", type: "cell", cellAddress: "K32" },
  { name: "PROG. INTE. SOC. (PIS)", type: "cell", cellAddress: "K33" },
  { name: "CONT. FINAN. SOC. COFINS", type: "cell", cellAddress: "K34" },
  { name: "IMP. PROD. IMP. (IPI)", type: "cell", cellAddress: "K35" },
  { name: "DUMPING", type: "cell", cellAddress: "K36" },
  { name: "IMP. CIRC. MERC. (ICMS)", type: "cell", cellAddress: "K37" },

  // SEÇÃO - DESPESAS OPERACIONAIS (índices 23-33)
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

  // SEÇÃO - OPERAÇÃO TRADING (índices 34-44)
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

  // SEÇÃO - PGTO NACIONA.NEEMAN (índice 45)
  {
    name: "PGTO EFETIVO NEEMAN",
    type: "formula",
    formula: "=SOMA([@[SISCOMEX]:[@[JANELA ESPECIAL]]])",
  },

  // SEÇÃO - STATUS NEEMAN (índice 46)
  { name: "STATUS NEEMAN", type: "manual" },

  // SEÇÃO - PGTO MARCHA ARMAZ (índices 47-48)
  { name: "DESP. FRETE", type: "cell", cellAddress: "H49" },
  { name: "DESP. Escolta", type: "cell", cellAddress: "H47" },

  // SEÇÃO - IMPOSTOS IR CSLL (índices 49-50)
  { name: "IMPOSTOS  IR", type: "cell", cellAddress: "K81" },
  { name: "IMPOSTO CSLL", type: "cell", cellAddress: "K82" },

  // SEÇÃO - RESUMO OPERAÇÃO (índices 51-54)
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

  // SEÇÃO - RESULTADO (índice 55)
  {
    name: "RESULTADO",
    type: "formula",
    formula: "=SOMA([@[APROVAÇÃO CLIENTE]]-[@RESUMO])",
  },

  // SEÇÃO - CRED NAC NEEMAN (índice 56)
  {
    name: "CC NEMAN",
    type: "formula",
    formula: "=[@[DESP. OPE. ENVI. NEEMAN]]-[@[PGTO EFETIVO NEEMAN]]",
  },

  // SEÇÃO - ANDAMENTO PROCESSO (índices 57-60)
  { name: "STATUS GERAL", type: "manual" },
  { name: "DESTINO", type: "manual" },
  { name: "OBSERVAÇÃO2", type: "manual" },
  { name: "DATA FINAL", type: "manual" },

  // SEÇÃO - VALOR FINAL CUSTO CLIENTE (índices 61-64)
  { name: "VALOR CUSTO CLIENTE", type: "cell", cellAddress: "J87" },
  {
    name: "ESTIMATIVA NACIONALIZAÇÃO EXTRATO C6",
    type: "cell",
    cellAddress: "J88",
  },
  { name: "CUSTO NACIONALIZAÇÃO NEEMAN", type: "cell", cellAddress: "J89" },
  { name: "DIFERENÇA NACIONALIZAÇÃO", type: "cell", cellAddress: "J90" },
];

// Estrutura de seções para mesclagem no Excel
const SECTIONS = [
  { name: "VALOR ADUNANEIRO - DI", startIndex: 0, endIndex: 4 },
  { name: "FECHAMENTO CAMBIO D.I - MARCHA", startIndex: 5, endIndex: 12 },
  { name: "PGTO ANDIAN. NACIONALIZAÇÃO", startIndex: 13, endIndex: 13 },
  { name: "DESPESAS ADUNEIRAS", startIndex: 14, endIndex: 16 },
  { name: "IMPOSTO", startIndex: 17, endIndex: 22 },
  { name: "DESPESAS OPERACIONAIS", startIndex: 23, endIndex: 33 },
  { name: "OPERAÇÃO TRADING", startIndex: 34, endIndex: 44 },
  { name: "PGTO NACIONA.NEEMAN", startIndex: 45, endIndex: 45 },
  { name: "STATUS NEEMAN", startIndex: 46, endIndex: 46 },
  { name: "PGTO MARCHA ARMAZ", startIndex: 47, endIndex: 48 },
  { name: "IMPOSTOS IR CSLL", startIndex: 49, endIndex: 50 },
  { name: "RESUMO OPERAÇÃO", startIndex: 51, endIndex: 54 },
  { name: "RESULTADO", startIndex: 55, endIndex: 55 },
  { name: "CRED NAC NEEMAN", startIndex: 56, endIndex: 56 },
  { name: "ANDAMENTO PROCESSO", startIndex: 57, endIndex: 60 },
  { name: "VALOR FINAL CUSTO CLIENTE", startIndex: 61, endIndex: 64 },
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
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState({
    current: 0,
    total: 0,
    currentFile: "",
  });
  const uploadCancelledRef = useRef(false);

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

          // Verifica se a planilha tem estrutura com seções (linha 1), headers (linha 2) e dados (linha 3+)
          // Verifica se a primeira linha contém algum nome de seção conhecido
          const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:Z1");
          let hasSectionRow = false;

          if (range.e.r >= 0) {
            // Verifica se a célula A1 contém algum nome de seção
            const cellA1 = worksheet["A1"];
            if (cellA1 && cellA1.v) {
              const cellValue = String(cellA1.v).toUpperCase();
              hasSectionRow = SECTIONS.some((section) =>
                cellValue.includes(section.name.toUpperCase())
              );
            }
          }

          let json: DataRow[];
          if (hasSectionRow && range.e.r >= 2) {
            // Pula linha 1 (seções) e linha 2 (headers), lê a partir da linha 3
            // Verifica se há dados além dos headers (linha 3 em diante)
            if (range.e.r > 2) {
              const dataRange = XLSX.utils.encode_range({
                s: { c: 0, r: 2 }, // Começa na linha 3 (índice 2)
                e: { c: range.e.c, r: range.e.r },
              });
              const dataSheet = XLSX.utils.sheet_to_json(worksheet, {
                range: dataRange,
                header: baseHeaders, // Usa os headers da estrutura master
              }) as DataRow[];
              json = dataSheet;
            } else {
              // Não há dados, apenas estrutura (seções e headers)
              json = [];
            }
          } else {
            // Estrutura normal: primeira linha são headers
            // Verifica se há mais de uma linha (além do header)
            if (range.e.r > 0) {
              json = XLSX.utils.sheet_to_json(worksheet) as DataRow[];
            } else {
              // Não há dados, apenas headers
              json = [];
            }
          }

          // Lista de nomes de seções e headers para filtrar
          const sectionNames = SECTIONS.map((s) => s.name.toUpperCase());
          const fieldNames = MASTER_STRUCTURE.map((f) => f.name.toUpperCase());

          const normalizedData = json
            .map((row) => {
              const normalizedRow: DataRow = {};
              MASTER_STRUCTURE.forEach((field) => {
                const fieldKey = Object.keys(row).find(
                  (key) =>
                    key.trim().toUpperCase() === field.name.trim().toUpperCase()
                );
                normalizedRow[field.name] = fieldKey ? row[fieldKey] : "";
              });
              return normalizedRow;
            })
            .filter((row) => {
              // Filtra linhas completamente vazias
              // Verifica se há pelo menos um campo com valor significativo
              const hasValue = Object.entries(row).some(
                ([fieldName, value]) => {
                  if (value === null || value === undefined) return false;
                  // Converte para string e verifica se não está vazio após trim
                  const strValue = String(value).trim();
                  if (strValue === "") return false;

                  // Ignora valores que são apenas nomes de seções ou headers
                  const upperValue = strValue.toUpperCase();
                  if (sectionNames.includes(upperValue)) return false;
                  if (
                    fieldNames.includes(upperValue) &&
                    fieldName.toUpperCase() !== upperValue
                  ) {
                    // Se o valor é um nome de campo mas não corresponde ao campo atual, ignora
                    return false;
                  }

                  return true;
                }
              );
              return hasValue;
            })
            .filter((row) => {
              // Filtra adicional: se todos os campos são strings vazias ou valores vazios
              // Isso pega casos onde o XLSX criou objetos vazios
              const allEmpty = Object.values(row).every((value) => {
                if (value === null || value === undefined) return true;
                const strValue = String(value).trim();
                return strValue === "";
              });
              return !allEmpty;
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
      setIsUploading(true);
      uploadCancelledRef.current = false;
      setUploadProgress({
        current: 0,
        total: newFiles.length,
        currentFile: "",
      });

      for (let i = 0; i < newFiles.length; i++) {
        if (uploadCancelledRef.current) {
          break;
        }

        const file = newFiles[i];
        setUploadProgress({
          current: i + 1,
          total: newFiles.length,
          currentFile: file.name,
        });

        try {
          const workbook = await readExtraFile(file);
          if (uploadCancelledRef.current) {
            break;
          }
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

      setIsUploading(false);
      setUploadProgress({ current: 0, total: 0, currentFile: "" });
      uploadCancelledRef.current = false;

      // Limpa o input para permitir upload do mesmo arquivo novamente
      if (e.target) {
        e.target.value = "";
      }
    }
  };

  const cancelUpload = () => {
    uploadCancelledRef.current = true;
    setIsUploading(false);
    setUploadProgress({ current: 0, total: 0, currentFile: "" });
  };

  const removeExtraFile = (index: number) => {
    setExtraFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const removeAllExtraFiles = () => {
    setExtraFiles([]);
  };

  const removeBaseFile = () => {
    setBaseFile(null);
    setBaseData([]);
    setConsolidatedData([]);
    setExtraFiles([]);
  };

  const downloadMasterTemplate = () => {
    // Cria worksheet vazio apenas com estrutura (seções e cabeçalhos)
    const worksheet = XLSX.utils.aoa_to_sheet([]);

    const maxCol = MASTER_STRUCTURE.length - 1;

    // Paleta de cores para as seções
    const sectionColors = [
      "5DADE2", // Azul médio - VALOR ADUNANEIRO - DI
      "52BE80", // Verde médio - FECHAMENTO CAMBIO D.I - MARCHA
      "F4D03F", // Amarelo médio - PGTO ANDIAN. NACIONALIZAÇÃO
      "EC7063", // Rosa médio - DESPESAS ADUNEIRAS
      "BB8FCE", // Roxo médio - IMPOSTO
      "5DADE2", // Azul médio - DESPESAS OPERACIONAIS
      "F39C12", // Laranja médio - OPERAÇÃO TRADING
      "1ABC9C", // Turquesa médio - PGTO NACIONA.NEEMAN
      "F7DC6F", // Dourado médio - STATUS NEEMAN
      "A569BD", // Lavanda médio - PGTO MARCHA ARMAZ
      "3498DB", // Azul mais escuro - IMPOSTOS IR CSLL
      "E74C3C", // Vermelho salmão - RESUMO OPERAÇÃO
      "27AE60", // Verde mais escuro - RESULTADO
      "F1C40F", // Amarelo ouro - CRED NAC NEEMAN
      "C39BD3", // Roxo rosado - ANDAMENTO PROCESSO
    ];

    // Função auxiliar para encontrar a cor de uma seção baseada no índice da coluna
    const getSectionColor = (colIndex: number): string => {
      const section = SECTIONS.find(
        (s) => colIndex >= s.startIndex && colIndex <= s.endIndex
      );
      if (section) {
        const sectionIndex = SECTIONS.indexOf(section);
        return sectionColors[sectionIndex % sectionColors.length];
      }
      return "FFFFFF"; // Branco como padrão
    };

    // Estilo de borda grossa para todas as células
    const thickBorder = {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } },
    };

    // Adiciona linha de seções na linha 1
    SECTIONS.forEach((section, index) => {
      const startCol = XLSX.utils.encode_col(section.startIndex);
      const startCell = `${startCol}1`;

      // Seleciona cor da paleta
      const color = sectionColors[index % sectionColors.length];

      // Define o valor da célula mesclada
      worksheet[startCell] = { v: section.name, t: "s" };
      worksheet[startCell].s = {
        alignment: { horizontal: "center", vertical: "center" },
        fill: {
          fgColor: { rgb: color },
          patternType: "solid",
        },
        font: { bold: true, sz: 11, color: { rgb: "000000" } },
        border: thickBorder,
      };

      // Adiciona merge
      if (!worksheet["!merges"]) {
        worksheet["!merges"] = [];
      }
      worksheet["!merges"].push({
        s: { c: section.startIndex, r: 0 },
        e: { c: section.endIndex, r: 0 },
      });
    });

    // Adiciona cabeçalhos na linha 2
    MASTER_STRUCTURE.forEach((field, colIndex) => {
      const excelCol = XLSX.utils.encode_col(colIndex);
      const headerCell = `${excelCol}2`;

      worksheet[headerCell] = { v: field.name, t: "s" };
      const sectionColor = getSectionColor(colIndex);
      worksheet[headerCell].s = {
        alignment: { horizontal: "center", vertical: "center" },
        fill: {
          fgColor: { rgb: sectionColor },
          patternType: "solid",
        },
        font: { bold: true, sz: 11, color: { rgb: "000000" } },
        border: thickBorder,
      };
    });

    // Ajusta o range do worksheet
    worksheet["!ref"] = `A1:${XLSX.utils.encode_col(maxCol)}2`;

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
    link.download = "planilha_master_vazia.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
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
          // Gera soma de células individuais (ex: O3+P3+Q3+...+AH3)
          const cells: string[] = [];
          for (let colIndex = startIndex; colIndex <= endIndex; colIndex++) {
            const excelCol = XLSX.utils.encode_col(colIndex);
            cells.push(`${excelCol}${rowIndex + 1}`);
          }
          return cells.join("+");
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

    // Remove qualquer "@" restante que possa ter ficado na fórmula
    excelFormula = excelFormula.replace(/@/g, "");

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
        // Pega o código do cliente da célula H7 da planilha avulsa
        const clientCodeFromFile = getCellValue(fileMapping.workbook, "H7");
        const clientCodeNormalized = clientCodeFromFile
          ? String(clientCodeFromFile).trim().toUpperCase()
          : "";

        const newRow: DataRow = {};

        MASTER_STRUCTURE.forEach((field) => {
          // Campos do tipo "flex" - processa condições por prioridade
          if (field.type === "flex" && field.cellCondition) {
            // Ordena condições por prioridade (menor número = maior prioridade)
            const sortedConditions = [...field.cellCondition].sort(
              (a, b) => a.priority - b.priority
            );

            let valueSet = false;

            for (const condition of sortedConditions) {
              if (valueSet) break;

              if (condition.type === "cell" && condition.cellAddress) {
                // Tenta pegar valor da célula
                const cellValue = getCellValue(
                  fileMapping.workbook,
                  condition.cellAddress
                );
                // Verifica se há valor válido (não vazio, não null, não undefined)
                const hasValue =
                  cellValue !== "" &&
                  cellValue !== null &&
                  cellValue !== undefined;
                if (hasValue) {
                  newRow[field.name] = cellValue;
                  valueSet = true;
                }
              } else if (condition.type === "formula" && condition.formula) {
                // Usa a fórmula
                newRow[field.name] = "__FORMULA__";
                valueSet = true;
              } else if (condition.type === "manual") {
                // Deixa vazio (manual)
                newRow[field.name] = "";
                valueSet = true;
              }
            }

            // Se nenhuma condição foi satisfeita, deixa vazio por padrão
            if (!valueSet) {
              newRow[field.name] = "";
            }
          }
          // Casos normais
          else if (field.type === "formula") {
            newRow[field.name] = "__FORMULA__";
          } else if (field.type === "cell") {
            const cellAddress = globalMapping[field.name];
            if (cellAddress && cellAddress.trim() !== "") {
              newRow[field.name] = getCellValue(
                fileMapping.workbook,
                cellAddress
              );
            } else {
              newRow[field.name] = "";
            }
          } else {
            // Tipo manual ou outros
            newRow[field.name] = "";
          }
        });

        // Verifica se já existe um registro com o mesmo código de cliente
        if (clientCodeNormalized) {
          const existingIndex = allData.findIndex((row) => {
            const existingCode = row["COD. CLIENTE"]
              ? String(row["COD. CLIENTE"]).trim().toUpperCase()
              : "";
            return existingCode === clientCodeNormalized;
          });

          if (existingIndex !== -1) {
            // Mescla inteligente: preserva campos manuais preenchidos, atualiza campos de célula
            const existingRow = allData[existingIndex];
            const mergedRow: DataRow = { ...existingRow };

            MASTER_STRUCTURE.forEach((field) => {
              const existingValue = existingRow[field.name];
              const newValue = newRow[field.name];

              // Para campos MANUAIS: preserva se já estava preenchido na master
              if (field.type === "manual") {
                if (
                  existingValue !== "" &&
                  existingValue !== null &&
                  existingValue !== undefined
                ) {
                  // Mantém o valor existente da master
                  mergedRow[field.name] = existingValue;
                } else {
                  // Se estava vazio, pode usar o novo valor (se houver)
                  mergedRow[field.name] = newValue || "";
                }
              }
              // Para campos CELL: sempre atualiza com o valor da avulsa
              else if (field.type === "cell") {
                mergedRow[field.name] = newValue;
              }
              // Para campos FORMULA: mantém como fórmula
              else if (field.type === "formula") {
                mergedRow[field.name] = "__FORMULA__";
              }
              // Para campos FLEX: usa o valor processado da avulsa
              else if (field.type === "flex") {
                mergedRow[field.name] = newValue;
              }
              // Outros casos: mantém o existente ou usa o novo
              else {
                mergedRow[field.name] =
                  newValue !== undefined ? newValue : existingValue;
              }
            });

            allData[existingIndex] = mergedRow;
          } else {
            // Adiciona como nova linha
            allData.push(newRow);
          }
        } else {
          // Se não houver código de cliente, adiciona como nova linha
          allData.push(newRow);
        }
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

    // Cria worksheet normalmente (cabeçalho na linha 1, dados a partir da linha 2)
    const worksheet = XLSX.utils.json_to_sheet(orderedData, {
      header: baseHeaders,
    });

    const maxRow = orderedData.length + 1; // dados + cabeçalho
    const maxCol = MASTER_STRUCTURE.length - 1;

    // Paleta de cores para as seções (cores mais escuras para diferenciar dos dados)
    const sectionColors = [
      "5DADE2", // Azul médio - VALOR ADUNANEIRO - DI
      "52BE80", // Verde médio - FECHAMENTO CAMBIO D.I - MARCHA
      "F4D03F", // Amarelo médio - PGTO ANDIAN. NACIONALIZAÇÃO
      "EC7063", // Rosa médio - DESPESAS ADUNEIRAS
      "BB8FCE", // Roxo médio - IMPOSTO
      "5DADE2", // Azul médio - DESPESAS OPERACIONAIS
      "F39C12", // Laranja médio - OPERAÇÃO TRADING
      "1ABC9C", // Turquesa médio - PGTO NACIONA.NEEMAN
      "F7DC6F", // Dourado médio - STATUS NEEMAN
      "A569BD", // Lavanda médio - PGTO MARCHA ARMAZ
      "3498DB", // Azul mais escuro - IMPOSTOS IR CSLL
      "E74C3C", // Vermelho salmão - RESUMO OPERAÇÃO
      "27AE60", // Verde mais escuro - RESULTADO
      "F1C40F", // Amarelo ouro - CRED NAC NEEMAN
      "C39BD3", // Roxo rosado - ANDAMENTO PROCESSO
    ];

    // Função auxiliar para encontrar a cor de uma seção baseada no índice da coluna
    const getSectionColor = (colIndex: number): string => {
      const section = SECTIONS.find(
        (s) => colIndex >= s.startIndex && colIndex <= s.endIndex
      );
      if (section) {
        const sectionIndex = SECTIONS.indexOf(section);
        return sectionColors[sectionIndex % sectionColors.length];
      }
      return "FFFFFF"; // Branco como padrão
    };

    // Função auxiliar para obter o formato de data apropriado
    const getDateFormat = (fieldName: string): string | null => {
      const fieldNameUpper = fieldName.toUpperCase();
      if (fieldNameUpper === "DATA FINAL") {
        // Formato de data brasileiro: dd/mm/aaaa
        return "dd/mm/yyyy";
      }
      return null;
    };

    // Função auxiliar para obter o formato de moeda apropriado
    const getCurrencyFormat = (fieldName: string): string | null => {
      const fieldNameUpper = fieldName.toUpperCase();

      // Campos que devem receber formatação de moeda mesmo tendo palavras comuns
      const currencyFields = [
        "FRETE AO CLIENTE",
        "APROVAÇÃO CLIENTE",
        "CUSTO CLIENTE",
        "VALOR CUSTO CLIENTE",
        "VALOR USD CLIENTE",
        "NF EMITIDA",
      ];

      // Verifica se é um campo que deve ter formatação de moeda
      if (
        currencyFields.some((currency) =>
          fieldNameUpper.includes(currency.toUpperCase())
        )
      ) {
        // VALOR USD CLIENTE usa formato contábil de dólar ($) - símbolo à esquerda, valor à direita
        if (fieldNameUpper.includes("VALOR USD CLIENTE")) {
          return '_("$"* #,##0.00_);_("$"* -#,##0.00_);_("$"* "-"??_);_(@_)'; // Formato contábil dólar
        }
        // Demais campos de moeda usam formato contábil Real brasileiro (R$) - símbolo à esquerda, valor à direita
        // Formato: positivo; negativo; zero; texto
        return '_("R$"* #,##0.00_);_("R$"* -#,##0.00_);_("R$"* "-"??_);_(@_)'; // Formato contábil Real
      }

      // Lista de campos que NÃO devem ser formatados como moeda (verificação exata)
      const nonCurrencyFields = [
        "STATUS DI",
        "STATUS SWIFT",
        "STATUS FINANCEIRO",
        "STATUS NEEMAN",
        "STATUS GERAL",
        "DESTINO",
        "OBSERVAÇÃO",
        "OBSERVAÇÃO2",
        "DATA FINAL",
        "COD. CLIENTE",
      ];

      // Verificação específica para "CLIENTE" (apenas o campo exato, não campos que contêm "CLIENTE")
      if (fieldNameUpper === "CLIENTE") {
        return null;
      }

      // Se está na lista de exceções, não formata como moeda
      if (
        nonCurrencyFields.some(
          (nonCurrency) => fieldNameUpper === nonCurrency.toUpperCase()
        )
      ) {
        return null;
      }

      // Demais campos usam formato contábil Real brasileiro (R$) - símbolo à esquerda, valor à direita
      // Isso inclui: TX, TAXA, FRETE AO CLIENTE, todas DESPESAS OPERACIONAIS, APROVAÇÃO CLIENTE, CUSTO CLIENTE
      // Formato: positivo; negativo; zero; texto
      return '_("R$"* #,##0.00_);_("R$"* -#,##0.00_);_("R$"* "-"??_);_(@_)'; // Formato contábil Real
    };

    // Estilo de borda grossa para todas as células
    const thickBorder = {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } },
    };

    // Primeiro, move dados para começar na linha 3 (de trás para frente para evitar sobrescrita)
    for (let rowIndex = maxRow; rowIndex >= 2; rowIndex--) {
      MASTER_STRUCTURE.forEach((field, colIndex) => {
        const excelCol = XLSX.utils.encode_col(colIndex);
        const oldCell = `${excelCol}${rowIndex}`;
        const newCell = `${excelCol}${rowIndex + 1}`;

        if (
          worksheet[oldCell] &&
          (worksheet[oldCell].v !== undefined ||
            worksheet[oldCell].f !== undefined)
        ) {
          worksheet[newCell] = worksheet[oldCell];
          // Aplica bordas, cor de fundo alternada e formatação de moeda nas células de dados
          if (rowIndex >= 2) {
            if (!worksheet[newCell].s) {
              worksheet[newCell].s = {};
            }
            worksheet[newCell].s.border = thickBorder;
            // Calcula se a linha é ímpar ou par (linha 3 = primeira linha de dados = ímpar)
            const dataRowIndex = rowIndex + 1 - 3; // Ajusta para começar em 0 na linha 3
            const rowColor = dataRowIndex % 2 === 0 ? "CAEDFB" : "FFFFFF"; // Azul claro para ímpares, branco para pares
            worksheet[newCell].s.fill = {
              fgColor: { rgb: rowColor },
              patternType: "solid",
            };
            // Aplica formatação de moeda ou data se necessário
            const field = MASTER_STRUCTURE[colIndex];
            if (field) {
              const currencyFormat = getCurrencyFormat(field.name);
              if (currencyFormat) {
                worksheet[newCell].z = currencyFormat;
              } else {
                const dateFormat = getDateFormat(field.name);
                if (dateFormat) {
                  worksheet[newCell].z = dateFormat;
                }
              }
            }
          }
          delete worksheet[oldCell];
        }
      });
    }

    // Move cabeçalhos da linha 1 para linha 2 e aplica cor de fundo
    MASTER_STRUCTURE.forEach((field, colIndex) => {
      const excelCol = XLSX.utils.encode_col(colIndex);
      const oldHeaderCell = `${excelCol}1`;
      const newHeaderCell = `${excelCol}2`;

      if (worksheet[oldHeaderCell]) {
        worksheet[newHeaderCell] = worksheet[oldHeaderCell];
        const sectionColor = getSectionColor(colIndex);
        worksheet[newHeaderCell].s = {
          alignment: { horizontal: "center", vertical: "center" },
          fill: {
            fgColor: { rgb: sectionColor },
            patternType: "solid",
          },
          font: { bold: true, sz: 11, color: { rgb: "000000" } },
          border: thickBorder,
        };
        delete worksheet[oldHeaderCell];
      }
    });

    // Adiciona linha de seções na linha 1
    SECTIONS.forEach((section, index) => {
      const startCol = XLSX.utils.encode_col(section.startIndex);
      const startCell = `${startCol}1`;

      // Seleciona cor da paleta (cicla se houver mais seções que cores)
      const color = sectionColors[index % sectionColors.length];

      // Define o valor da célula mesclada
      worksheet[startCell] = { v: section.name, t: "s" };
      worksheet[startCell].s = {
        alignment: { horizontal: "center", vertical: "center" },
        fill: {
          fgColor: { rgb: color },
          patternType: "solid",
        },
        font: { bold: true, sz: 11, color: { rgb: "000000" } },
        border: thickBorder,
      };

      // Adiciona merge
      if (!worksheet["!merges"]) {
        worksheet["!merges"] = [];
      }
      worksheet["!merges"].push({
        s: { c: section.startIndex, r: 0 },
        e: { c: section.endIndex, r: 0 },
      });
    });

    // Ajusta fórmulas para a nova posição e aplica bordas nas células de dados
    orderedData.forEach((row, rowIndex) => {
      MASTER_STRUCTURE.forEach((field, colIndex) => {
        const excelCol = XLSX.utils.encode_col(colIndex);
        const cellAddress = `${excelCol}${rowIndex + 3}`; // linha 3 em diante

        // Verifica se precisa aplicar fórmula
        if (row[field.name] === "__FORMULA__") {
          let formulaToUse: string | undefined;

          // Se for campo flex, encontra a condição com fórmula
          if (field.type === "flex" && field.cellCondition) {
            const formulaCondition = field.cellCondition.find(
              (c) => c.type === "formula" && c.formula
            );
            formulaToUse = formulaCondition?.formula;
          }
          // Se for campo formula normal
          else if (field.type === "formula" && field.formula) {
            formulaToUse = field.formula;
          }

          if (formulaToUse) {
            const excelFormula = convertFormulaToExcelFormat(
              formulaToUse,
              rowIndex + 2 // ajuste para linha 3
            );
            if (!worksheet[cellAddress]) {
              worksheet[cellAddress] = {};
            }
            worksheet[cellAddress].f = excelFormula.replace(/^=/, "");
          }
        }

        // Aplica bordas, cor de fundo alternada e formatação de moeda apenas em células com conteúdo
        if (
          worksheet[cellAddress] &&
          (worksheet[cellAddress].v !== undefined ||
            worksheet[cellAddress].f !== undefined)
        ) {
          if (!worksheet[cellAddress].s) {
            worksheet[cellAddress].s = {};
          }
          worksheet[cellAddress].s.border = thickBorder;
          // Linhas ímpares (rowIndex 0, 2, 4...) = azul, linhas pares (rowIndex 1, 3, 5...) = branco
          const rowColor = rowIndex % 2 === 0 ? "CAEDFB" : "FFFFFF"; // Azul claro para ímpares, branco para pares
          worksheet[cellAddress].s.fill = {
            fgColor: { rgb: rowColor },
            patternType: "solid",
          };

          // Aplica formatação de moeda ou data se necessário
          const currencyFormat = getCurrencyFormat(field.name);
          if (currencyFormat) {
            worksheet[cellAddress].z = currencyFormat;
          } else {
            const dateFormat = getDateFormat(field.name);
            if (dateFormat) {
              worksheet[cellAddress].z = dateFormat;
            }
          }
        }
      });
    });

    // Ajusta o range do worksheet
    worksheet["!ref"] = `A1:${XLSX.utils.encode_col(maxCol)}${maxRow + 2}`;

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
    (f) => f.type === "formula" || f.type === "flex"
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
      return (
        matchesSearch && (field.type === "formula" || field.type === "flex")
      );
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
                        className="text-white hover:bg-gray-800 hover:text-white"
                      >
                        <X className="h-5 w-5" />
                      </Button>
                    </div>
                  </div>

                  <div className="p-5">
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
                            field.type === "formula" || field.type === "flex"
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
                            {(field.type === "formula" ||
                              field.type === "flex") && (
                              <span className="px-2 py-0.5 bg-gray-200 text-gray-700 rounded text-xs font-medium shrink-0">
                                {field.type === "flex" ? "Flex" : "Auto"}
                              </span>
                            )}
                            {globalMapping[field.name] &&
                              field.type !== "formula" &&
                              field.type !== "flex" && (
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
                          ) : field.type === "flex" ? (
                            <div className="space-y-1">
                              {field.cellCondition?.map((condition, idx) => (
                                <div
                                  key={idx}
                                  className="text-xs text-gray-600 bg-gray-100 p-2 rounded"
                                >
                                  <span className="font-medium">
                                    {condition.priority}. {condition.type}
                                  </span>
                                  {condition.cellAddress && (
                                    <span className="ml-2">
                                      {condition.cellAddress}
                                    </span>
                                  )}
                                  {condition.formula && (
                                    <div className="mt-1 text-gray-500 truncate">
                                      {condition.formula}
                                    </div>
                                  )}
                                </div>
                              ))}
                            </div>
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
                      <div className="flex items-center gap-2">
                        <Button
                          onClick={downloadMasterTemplate}
                          size="sm"
                          variant="outline"
                          className="gap-1 md:gap-2 bg-white/10 hover:bg-white/20 border-white/20 text-white hover:text-white"
                        >
                          <Download className="h-4 w-4" />
                          <span className="hidden md:inline">
                            Baixar Template
                          </span>
                        </Button>
                        <Button
                          onClick={() => setShowMapping(!showMapping)}
                          size="sm"
                          variant={showMapping ? "default" : "outline"}
                          className="gap-1 md:gap-2 bg-white/10 hover:bg-white/20 border-white/20 text-white hover:text-white"
                        >
                          <Settings className="h-4 w-4" />
                          <span className="hidden md:inline">
                            {showMapping ? "Fechar" : "Configurar"}
                          </span>
                          <span className="ml-1 px-2 py-0.5 bg-white/20 rounded-full text-xs font-semibold hidden md:inline-flex">
                            {mappedFieldsCount}/
                            {totalFields - formulaFieldsCount}
                          </span>
                        </Button>
                      </div>
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
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <Plus className="h-6 w-6" />
                        <div>
                          <h2 className="text-lg font-bold">
                            Planilhas Avulsas
                          </h2>
                          <p className="text-sm text-gray-300">
                            Adicione múltiplos arquivos
                          </p>
                        </div>
                      </div>
                      {extraFiles.length > 0 && !isUploading && (
                        <div className="flex items-center gap-2">
                          <label>
                            <input
                              id="extra-files-input"
                              type="file"
                              className="hidden"
                              accept=".xlsx,.xls,.csv"
                              multiple
                              onChange={handleExtraFilesUpload}
                              disabled={!baseFile || isUploading}
                            />
                            <Button
                              type="button"
                              variant="outline"
                              size="sm"
                              className="gap-1 md:gap-2 bg-white/10 hover:bg-white/20 border-white/20 text-white hover:text-white"
                              disabled={!baseFile || isUploading}
                              onClick={() => {
                                const input = document.getElementById(
                                  "extra-files-input"
                                ) as HTMLInputElement;
                                input?.click();
                              }}
                            >
                              <Plus className="h-4 w-4" />
                              <span className="hidden md:inline">
                                Adicionar mais
                              </span>
                            </Button>
                          </label>
                          <Button
                            type="button"
                            variant="outline"
                            size="sm"
                            className="gap-1 md:gap-2 bg-red-600/20 hover:bg-red-600/30 border-red-400/50 text-white hover:text-white"
                            disabled={isUploading}
                            onClick={removeAllExtraFiles}
                          >
                            <X className="h-4 w-4" />
                            <span className="hidden md:inline">
                              Excluir todas
                            </span>
                          </Button>
                        </div>
                      )}
                      {isUploading && (
                        <Button
                          type="button"
                          variant="outline"
                          size="sm"
                          className="gap-2 bg-red-600/20 hover:bg-red-600/30 border-red-400/50 text-white hover:text-white"
                          onClick={cancelUpload}
                        >
                          <X className="h-4 w-4" />
                          Cancelar Upload
                        </Button>
                      )}
                    </div>
                  </div>

                  <div className="p-6 h-full flex flex-col">
                    {isUploading && (
                      <div className="flex-1 flex flex-col items-center justify-center bg-gradient-to-br from-blue-50 to-indigo-50 rounded-lg border-2 border-blue-200 p-8 max-h-[200px]">
                        <div className="flex flex-col items-center gap-4 w-full max-w-md">
                          {/* <div className="relative">
                            <Loader2 className="h-12 w-12 animate-spin text-blue-600" />
                            <div className="absolute inset-0 flex items-center justify-center">
                              <div className="h-6 w-6 bg-blue-100 rounded-full"></div>
                            </div>
                          </div> */}

                          <div className="text-center w-full">
                            <h3 className="text-lg font-semibold text-gray-900 mb-1">
                              Fazendo upload...
                            </h3>
                            <p className="text-sm text-gray-600">
                              {uploadProgress.current} de {uploadProgress.total}{" "}
                              arquivo(s)
                            </p>
                          </div>

                          {uploadProgress.currentFile && (
                            <div className="w-full bg-white/60 rounded-lg p-3 border border-blue-200">
                              <p className="text-xs font-medium text-gray-700 mb-1">
                                Arquivo atual:
                              </p>
                              <p className="text-sm text-gray-900 truncate font-medium">
                                {uploadProgress.currentFile}
                              </p>
                            </div>
                          )}

                          <div className="w-full space-y-2">
                            <div className="w-full bg-gray-200 rounded-full h-3 overflow-hidden shadow-inner">
                              <div
                                className="bg-gradient-to-r from-blue-500 to-indigo-600 h-3 rounded-full transition-all duration-500 ease-out shadow-sm"
                                style={{
                                  width: `${
                                    (uploadProgress.current /
                                      uploadProgress.total) *
                                    100
                                  }%`,
                                }}
                              />
                            </div>
                            <p className="text-xs text-center text-gray-500">
                              {Math.round(
                                (uploadProgress.current /
                                  uploadProgress.total) *
                                  100
                              )}
                              % concluído
                            </p>
                          </div>
                        </div>
                      </div>
                    )}

                    {extraFiles.length === 0 && !isUploading && (
                      <label
                        className={`flex flex-col items-center justify-center w-full h-full border-2 border-dashed rounded-lg transition-all ${
                          !baseFile || isUploading
                            ? "border-gray-200 bg-gray-50 cursor-not-allowed opacity-50"
                            : "border-gray-300 bg-gray-50 hover:bg-gray-100 cursor-pointer"
                        }`}
                      >
                        <div className="flex flex-col items-center justify-center gap-2">
                          <Plus
                            className={`h-6 w-6 ${
                              !baseFile || isUploading
                                ? "text-gray-300"
                                : "text-gray-400"
                            }`}
                          />
                          <div className="text-center">
                            <p
                              className={`text-sm font-semibold ${
                                !baseFile || isUploading
                                  ? "text-gray-400"
                                  : "text-gray-700"
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
                          disabled={!baseFile || isUploading}
                        />
                      </label>
                    )}

                    {extraFiles.length > 0 && !isUploading && (
                      <div className="space-y-2 max-h-52 overflow-y-auto custom-scrollbar pr-1">
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
                              disabled={isUploading}
                              className="hover:bg-red-50 hover:text-red-600 shrink-0 disabled:opacity-50"
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
              !isProcessing &&
              !isUploading && (
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
                      className="gap-2 px-8 py-6 text-lg font-semibold bg-blue-600 hover:bg-blue-700 text-white"
                    >
                      <Download className="h-6 w-6" />
                      Baixar Planilha
                    </Button>
                    <Button
                      onClick={handleNewConsolidation}
                      size="lg"
                      className="gap-2 px-6 py-6 text-lg font-semibold bg-gray-900 hover:bg-gray-800 text-white"
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
