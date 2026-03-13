import { useEffect, useMemo, useRef, useState } from "react";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import type { PDFDocumentProxy } from "pdfjs-dist/types/src/display/api";
import * as XLSX from "xlsx";

GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

type ExtractMode = "guided" | "data-to-columns";
type DelimiterOption = "double-space" | "semicolon" | "comma" | "tab" | "custom";

interface PositionedText {
  text: string;
  x: number;
  y: number;
  width: number;
}

interface RowGroup {
  y: number;
  items: PositionedText[];
}

interface CellPosition {
  row: number;
  column: number;
}

const MIN_GUIDE_GAP = 48;
const ROW_TOLERANCE = 8;
const MAX_VIEWPORT_WIDTH = 980;
const MIN_ZOOM = 0.5;
const MAX_ZOOM = 2;
const ZOOM_STEP = 0.1;

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(value, max));
}

function getColumnName(index: number): string {
  let name = "";
  let cursor = index + 1;
  while (cursor > 0) {
    const remainder = (cursor - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    cursor = Math.floor((cursor - 1) / 26);
  }
  return name;
}

function normalizeGuides(rawGuides: number[], width: number): number[] {
  if (width <= 0 || rawGuides.length === 0) {
    return [];
  }

  const sorted = [...rawGuides].sort((a, b) => a - b);
  const normalized: number[] = [];

  for (let index = 0; index < sorted.length; index += 1) {
    const remaining = sorted.length - index - 1;
    const min =
      normalized.length > 0 ? normalized[normalized.length - 1] + MIN_GUIDE_GAP : MIN_GUIDE_GAP / 2;
    const max = width - MIN_GUIDE_GAP / 2 - remaining * MIN_GUIDE_GAP;
    normalized.push(clamp(sorted[index], min, max));
  }

  return normalized;
}

function scaleGuides(guides: number[], oldWidth: number, newWidth: number): number[] {
  if (newWidth <= 0) {
    return [];
  }
  if (guides.length === 0 || oldWidth <= 0) {
    return [newWidth / 3, (newWidth * 2) / 3];
  }
  return normalizeGuides(guides.map((guide) => (guide / oldWidth) * newWidth), newWidth);
}

function groupRows(items: PositionedText[]): RowGroup[] {
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const groups: RowGroup[] = [];

  for (const item of sorted) {
    const last = groups[groups.length - 1];
    if (last && Math.abs(last.y - item.y) <= ROW_TOLERANCE) {
      const oldCount = last.items.length;
      last.items.push(item);
      last.y = (last.y * oldCount + item.y) / (oldCount + 1);
      continue;
    }
    groups.push({ y: item.y, items: [item] });
  }

  return groups;
}

function composeLineWithSpacing(items: PositionedText[]): string {
  const sorted = [...items].sort((a, b) => a.x - b.x);
  if (sorted.length === 0) {
    return "";
  }

  let line = sorted[0].text;
  for (let index = 1; index < sorted.length; index += 1) {
    const previous = sorted[index - 1];
    const current = sorted[index];
    const gap = current.x - (previous.x + previous.width);
    line += gap > 12 ? "  " : " ";
    line += current.text;
  }

  return line.replace(/\s{3,}/g, "  ").trim();
}

function buildGuidedRows(groups: RowGroup[], guides: number[], width: number): string[][] {
  const boundaries = [0, ...normalizeGuides(guides, width), width];
  const totalColumns = Math.max(boundaries.length - 1, 1);

  const rows = groups.map((group) => {
    const buckets: string[][] = Array.from({ length: totalColumns }, () => []);
    const sortedItems = [...group.items].sort((a, b) => a.x - b.x);

    for (const item of sortedItems) {
      const center = item.x + item.width / 2;
      let selected = -1;
      for (let column = 0; column < totalColumns; column += 1) {
        const start = boundaries[column];
        const end = boundaries[column + 1];
        if (center >= start && center < end) {
          selected = column;
          break;
        }
      }

      if (selected === -1 && center >= width) {
        selected = totalColumns - 1;
      }
      if (selected >= 0) {
        buckets[selected].push(item.text);
      }
    }

    return buckets.map((parts) => parts.join(" ").replace(/\s+/g, " ").trim());
  });

  return rows.filter((row) => row.some((cell) => cell.length > 0));
}

function getDelimiterRegex(option: DelimiterOption, custom: string): RegExp {
  switch (option) {
    case "semicolon":
      return /;/g;
    case "comma":
      return /,/g;
    case "tab":
      return /\t+/g;
    case "custom":
      if (custom.trim().length === 0) {
        return /\s{2,}/g;
      }
      try {
        return new RegExp(custom, "g");
      } catch {
        return /\s{2,}/g;
      }
    case "double-space":
    default:
      return /\s{2,}/g;
  }
}

function buildDataToColumnsRows(groups: RowGroup[], option: DelimiterOption, customDelimiter: string): string[][] {
  const delimiterRegex = getDelimiterRegex(option, customDelimiter);
  const rows = groups.map((group) => {
    const line = composeLineWithSpacing(group.items);
    if (!line) {
      return [];
    }
    return line
      .split(delimiterRegex)
      .map((part) => part.trim())
      .filter((part) => part.length > 0);
  });
  return rows.filter((row) => row.length > 0);
}

function normalizeTable(rows: string[][]): string[][] {
  const maxColumns = rows.reduce((currentMax, row) => Math.max(currentMax, row.length), 0);
  if (maxColumns === 0) {
    return [];
  }
  return rows.map((row) => {
    const normalized = [...row];
    while (normalized.length < maxColumns) {
      normalized.push("");
    }
    return normalized;
  });
}

function getClientX(event: MouseEvent | TouchEvent): number | null {
  if ("touches" in event && event.touches.length > 0) {
    return event.touches[0].clientX;
  }
  if ("changedTouches" in event && event.changedTouches.length > 0) {
    return event.changedTouches[0].clientX;
  }
  if ("clientX" in event) {
    return event.clientX;
  }
  return null;
}

export default function App() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const stageRef = useRef<HTMLDivElement | null>(null);
  const sheetWrapperRef = useRef<HTMLDivElement | null>(null);
  const editorRef = useRef<HTMLInputElement | null>(null);
  const pdfDocumentRef = useRef<PDFDocumentProxy | null>(null);
  const renderTokenRef = useRef(0);
  const previousCanvasWidthRef = useRef(0);

  const [pdfDocument, setPdfDocument] = useState<PDFDocumentProxy | null>(null);
  const [fileName, setFileName] = useState("");
  const [pageNumber, setPageNumber] = useState(1);
  const [pageCount, setPageCount] = useState(0);
  const [canvasSize, setCanvasSize] = useState({ width: 0, height: 0 });
  const [textItems, setTextItems] = useState<PositionedText[]>([]);
  const [guides, setGuides] = useState<number[]>([]);
  const [draggingGuideIndex, setDraggingGuideIndex] = useState<number | null>(null);
  const [extractMode, setExtractMode] = useState<ExtractMode>("guided");
  const [delimiterOption, setDelimiterOption] = useState<DelimiterOption>("double-space");
  const [customDelimiter, setCustomDelimiter] = useState("\\|");
  const [tableData, setTableData] = useState<string[][]>([]);
  const [selectedCell, setSelectedCell] = useState<CellPosition | null>(null);
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);
  const [selectedColumnIndex, setSelectedColumnIndex] = useState<number | null>(null);
  const [editingCell, setEditingCell] = useState<CellPosition | null>(null);
  const [pdfZoom, setPdfZoom] = useState(1);
  const [sheetZoom, setSheetZoom] = useState(1);
  const [status, setStatus] = useState("Anexe um PDF para iniciar.");

  const headers = useMemo(() => {
    const count = tableData[0]?.length ?? 0;
    return Array.from({ length: count }, (_, index) => getColumnName(index));
  }, [tableData]);
  const rowCount = tableData.length;
  const columnCount = headers.length;

  useEffect(() => {
    pdfDocumentRef.current = pdfDocument;
  }, [pdfDocument]);

  useEffect(() => {
    return () => {
      const currentDocument = pdfDocumentRef.current;
      if (currentDocument) {
        void currentDocument.destroy();
        pdfDocumentRef.current = null;
      }
    };
  }, []);

  useEffect(() => {
    if (!pdfDocument || !canvasRef.current) {
      return;
    }

    let cancelled = false;
    let activeRenderTask: { cancel: () => void; promise: Promise<void> } | null = null;
    const currentToken = renderTokenRef.current + 1;
    renderTokenRef.current = currentToken;

    const renderPage = async () => {
      try {
        setStatus(`Renderizando pagina ${pageNumber}...`);
        const page = await pdfDocument.getPage(pageNumber);
        const unscaledViewport = page.getViewport({ scale: 1 });
        const scale = (MAX_VIEWPORT_WIDTH / unscaledViewport.width) * pdfZoom;
        const viewport = page.getViewport({ scale });
        const canvas = canvasRef.current;
        if (!canvas) {
          return;
        }

        const context = canvas.getContext("2d");
        if (!context) {
          throw new Error("Falha ao acessar contexto do canvas.");
        }

        canvas.width = viewport.width;
        canvas.height = viewport.height;
        activeRenderTask = page.render({ canvasContext: context, viewport });
        await activeRenderTask.promise;

        const textContent = await page.getTextContent();
        const positioned: PositionedText[] = [];
        for (const item of textContent.items) {
          if (!("str" in item) || typeof item.str !== "string") {
            continue;
          }
          const cleaned = item.str.trim();
          if (!cleaned || !("transform" in item)) {
            continue;
          }
          const transform = item.transform;
          if (!Array.isArray(transform) || transform.length < 6) {
            continue;
          }

          const rawX = Number(transform[4]);
          const rawY = Number(transform[5]);
          const [viewportX, viewportY] = viewport.convertToViewportPoint(rawX, rawY);
          const rawWidth = "width" in item && typeof item.width === "number" ? item.width : 0;
          const [rightX] = viewport.convertToViewportPoint(rawX + rawWidth, rawY);
          const width = Math.max(2, Math.abs(rightX - viewportX));

          positioned.push({
            text: cleaned,
            x: clamp(viewportX, 0, viewport.width),
            y: clamp(viewportY, 0, viewport.height),
            width
          });
        }

        if (cancelled || renderTokenRef.current !== currentToken) {
          return;
        }

        setTextItems(positioned);
        setCanvasSize({ width: viewport.width, height: viewport.height });
        setGuides((previousGuides) => scaleGuides(previousGuides, previousCanvasWidthRef.current, viewport.width));
        previousCanvasWidthRef.current = viewport.width;
        setStatus(`Pagina ${pageNumber}/${pdfDocument.numPages} pronta.`);
      } catch (error) {
        if (cancelled) {
          return;
        }

        const maybeError = error as { name?: string; message?: string };
        if (maybeError?.name === "RenderingCancelledException") {
          return;
        }

        if (!cancelled) {
          const message = error instanceof Error ? error.message : "Erro ao renderizar PDF.";
          setStatus(message);
        }
      }
    };

    void renderPage();

    return () => {
      cancelled = true;
      activeRenderTask?.cancel();
    };
  }, [pdfDocument, pageNumber, pdfZoom]);

  useEffect(() => {
    if (draggingGuideIndex === null || canvasSize.width <= 0 || !stageRef.current) {
      return;
    }

    const handleMove = (event: MouseEvent | TouchEvent) => {
      const clientX = getClientX(event);
      if (clientX === null || !stageRef.current) {
        return;
      }

      const rect = stageRef.current.getBoundingClientRect();
      const relativeX = clientX - rect.left;

      setGuides((currentGuides) => {
        const sorted = [...currentGuides].sort((a, b) => a - b);
        if (draggingGuideIndex < 0 || draggingGuideIndex >= sorted.length) {
          return sorted;
        }

        const previous = draggingGuideIndex > 0 ? sorted[draggingGuideIndex - 1] : 0;
        const next =
          draggingGuideIndex < sorted.length - 1 ? sorted[draggingGuideIndex + 1] : canvasSize.width;
        sorted[draggingGuideIndex] = clamp(relativeX, previous + MIN_GUIDE_GAP, next - MIN_GUIDE_GAP);
        return sorted;
      });

      event.preventDefault();
    };

    const handleEnd = () => {
      setDraggingGuideIndex(null);
    };

    window.addEventListener("mousemove", handleMove);
    window.addEventListener("mouseup", handleEnd);
    window.addEventListener("touchmove", handleMove, { passive: false });
    window.addEventListener("touchend", handleEnd);

    return () => {
      window.removeEventListener("mousemove", handleMove);
      window.removeEventListener("mouseup", handleEnd);
      window.removeEventListener("touchmove", handleMove);
      window.removeEventListener("touchend", handleEnd);
    };
  }, [draggingGuideIndex, canvasSize.width]);

  useEffect(() => {
    if (textItems.length === 0 || canvasSize.width <= 0) {
      setTableData([]);
      return;
    }

    const groupedRows = groupRows(textItems);
    const extractedRows =
      extractMode === "guided"
        ? buildGuidedRows(groupedRows, guides, canvasSize.width)
        : buildDataToColumnsRows(groupedRows, delimiterOption, customDelimiter);

    setTableData(normalizeTable(extractedRows));
  }, [textItems, guides, extractMode, delimiterOption, customDelimiter, canvasSize.width]);

  useEffect(() => {
    if (!editingCell || !editorRef.current) {
      return;
    }
    editorRef.current.focus();
    editorRef.current.select();
  }, [editingCell]);

  useEffect(() => {
    if (rowCount === 0 || columnCount === 0) {
      setSelectedCell(null);
      setSelectedRowIndex(null);
      setSelectedColumnIndex(null);
      setEditingCell(null);
      return;
    }

    setSelectedCell((current) =>
      current ? { row: clamp(current.row, 0, rowCount - 1), column: clamp(current.column, 0, columnCount - 1) } : current
    );
    setSelectedRowIndex((current) => (current === null ? current : clamp(current, 0, rowCount - 1)));
    setSelectedColumnIndex((current) => (current === null ? current : clamp(current, 0, columnCount - 1)));
    setEditingCell((current) =>
      current ? { row: clamp(current.row, 0, rowCount - 1), column: clamp(current.column, 0, columnCount - 1) } : current
    );
  }, [rowCount, columnCount]);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (!selectedFile) {
      return;
    }

    try {
      setStatus("Lendo arquivo PDF...");
      const buffer = await selectedFile.arrayBuffer();
      const task = getDocument({ data: buffer });
      const loadedDocument = await task.promise;

      const previousDocument = pdfDocumentRef.current;
      if (previousDocument && previousDocument !== loadedDocument) {
        try {
          await previousDocument.destroy();
        } catch {
          // Ignora falhas de cleanup para priorizar novo carregamento.
        }
      }

      setPdfDocument(loadedDocument);
      pdfDocumentRef.current = loadedDocument;
      setFileName(selectedFile.name);
      setPageCount(loadedDocument.numPages);
      setPageNumber(1);
      setPdfZoom(1);
      setTextItems([]);
      setTableData([]);
      setStatus(`PDF "${selectedFile.name}" carregado. Ajuste as colunas e exporte.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Nao foi possivel abrir este PDF.";
      setStatus(message);
    } finally {
      event.target.value = "";
    }
  };

  const addColumnGuide = () => {
    if (canvasSize.width <= 0) {
      return;
    }
    setGuides((currentGuides) => {
      const sorted = normalizeGuides(currentGuides, canvasSize.width);
      const boundaries = [0, ...sorted, canvasSize.width];
      let largestGap = 0;
      let insertionPoint = canvasSize.width / 2;

      for (let index = 0; index < boundaries.length - 1; index += 1) {
        const gap = boundaries[index + 1] - boundaries[index];
        if (gap > largestGap) {
          largestGap = gap;
          insertionPoint = boundaries[index] + gap / 2;
        }
      }
      return normalizeGuides([...sorted, insertionPoint], canvasSize.width);
    });
  };

  const removeColumnGuide = () => {
    setGuides((currentGuides) => (currentGuides.length === 0 ? currentGuides : currentGuides.slice(0, -1)));
  };

  const updateCell = (rowIndex: number, columnIndex: number, value: string) => {
    setTableData((currentRows) => {
      const nextRows = currentRows.map((row) => [...row]);
      if (!nextRows[rowIndex]) {
        return currentRows;
      }
      nextRows[rowIndex][columnIndex] = value;
      return nextRows;
    });
  };

  const changePdfZoom = (delta: number) => {
    setPdfZoom((current) => clamp(Number((current + delta).toFixed(2)), MIN_ZOOM, MAX_ZOOM));
  };

  const changeSheetZoom = (delta: number) => {
    setSheetZoom((current) => clamp(Number((current + delta).toFixed(2)), MIN_ZOOM, MAX_ZOOM));
  };

  const focusSheet = () => {
    sheetWrapperRef.current?.focus();
  };

  const runOnEnterOrSpace = (event: React.KeyboardEvent<HTMLElement>, action: () => void) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      action();
    }
  };

  const activateCellSelection = (rowIndex: number, columnIndex: number) => {
    setSelectedCell({ row: rowIndex, column: columnIndex });
    setSelectedRowIndex(null);
    setSelectedColumnIndex(null);
    focusSheet();
  };

  const startCellEditing = (rowIndex: number, columnIndex: number) => {
    setSelectedCell({ row: rowIndex, column: columnIndex });
    setSelectedRowIndex(null);
    setSelectedColumnIndex(null);
    setEditingCell({ row: rowIndex, column: columnIndex });
  };

  const resolveAnchorCell = (): CellPosition => {
    if (selectedCell) {
      return selectedCell;
    }
    if (selectedRowIndex !== null) {
      return { row: clamp(selectedRowIndex, 0, rowCount - 1), column: 0 };
    }
    if (selectedColumnIndex !== null) {
      return { row: 0, column: clamp(selectedColumnIndex, 0, columnCount - 1) };
    }
    return { row: 0, column: 0 };
  };

  const handleSheetKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (rowCount === 0 || columnCount === 0 || editingCell) {
      return;
    }

    const anchor = resolveAnchorCell();
    let nextRow = anchor.row;
    let nextColumn = anchor.column;

    if (event.key === "ArrowUp") {
      event.preventDefault();
      nextRow = clamp(anchor.row - 1, 0, rowCount - 1);
    } else if (event.key === "ArrowDown") {
      event.preventDefault();
      nextRow = clamp(anchor.row + 1, 0, rowCount - 1);
    } else if (event.key === "ArrowLeft") {
      event.preventDefault();
      nextColumn = clamp(anchor.column - 1, 0, columnCount - 1);
    } else if (event.key === "ArrowRight") {
      event.preventDefault();
      nextColumn = clamp(anchor.column + 1, 0, columnCount - 1);
    } else if (event.key === "Tab") {
      event.preventDefault();
      nextColumn = clamp(anchor.column + (event.shiftKey ? -1 : 1), 0, columnCount - 1);
    } else if (event.key === "Enter") {
      event.preventDefault();
      nextRow = clamp(anchor.row + 1, 0, rowCount - 1);
    } else if (event.key === "F2") {
      event.preventDefault();
      startCellEditing(anchor.row, anchor.column);
      return;
    } else {
      return;
    }

    setSelectedCell({ row: nextRow, column: nextColumn });
    setSelectedRowIndex(null);
    setSelectedColumnIndex(null);
  };

  const addSheetRow = () => {
    setEditingCell(null);
    setTableData((currentRows) => {
      const width = currentRows[0]?.length ?? 1;
      return [...currentRows, Array.from({ length: width }, () => "")];
    });
  };

  const removeSheetRow = () => {
    setEditingCell(null);
    setTableData((currentRows) => {
      if (currentRows.length === 0) {
        return currentRows;
      }

      const targetRow =
        selectedRowIndex !== null
          ? clamp(selectedRowIndex, 0, currentRows.length - 1)
          : selectedCell
            ? clamp(selectedCell.row, 0, currentRows.length - 1)
            : currentRows.length - 1;

      return currentRows.filter((_, rowIndex) => rowIndex !== targetRow);
    });
  };

  const addSheetColumn = () => {
    setEditingCell(null);
    setTableData((currentRows) => {
      if (currentRows.length === 0) {
        return [[""]];
      }
      return currentRows.map((row) => [...row, ""]);
    });
  };

  const removeSheetColumn = () => {
    setEditingCell(null);
    setTableData((currentRows) => {
      const width = currentRows[0]?.length ?? 0;
      if (width === 0) {
        return currentRows;
      }
      if (width === 1) {
        return [];
      }

      const targetColumn =
        selectedColumnIndex !== null
          ? clamp(selectedColumnIndex, 0, width - 1)
          : selectedCell
            ? clamp(selectedCell.column, 0, width - 1)
            : width - 1;

      return currentRows.map((row) => row.filter((_, columnIndex) => columnIndex !== targetColumn));
    });
  };

  const exportToXlsx = () => {
    if (tableData.length === 0) {
      setStatus("Sem dados para exportar.");
      return;
    }
    const worksheet = XLSX.utils.aoa_to_sheet(tableData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");
    const cleanName = fileName.replace(/\.pdf$/i, "").trim() || "resultado";
    XLSX.writeFile(workbook, `${cleanName}.xlsx`);
    setStatus("Arquivo XLSX exportado com sucesso.");
  };

  return (
    <main className="app-shell">
      <section className="panel controls-panel">
        <h1>PDF para XLSX</h1>
        <p>Faça upload do PDF e ajuste as colunas.</p>

        <div className="controls-grid">
          <label className="field">
            <span>Arquivo PDF</span>
            <input type="file" accept="application/pdf" onChange={handleFileUpload} />
          </label>

          <label className="field">
            <span>Modo de extracao</span>
            <select
              value={extractMode}
              onChange={(event) => setExtractMode(event.target.value as ExtractMode)}
              disabled={!pdfDocument}
            >
              <option value="guided">Guias visuais de coluna</option>
              <option value="data-to-columns">Dados para Colunas (delimitador)</option>
            </select>
          </label>

          {extractMode === "data-to-columns" ? (
            <>
              <label className="field">
                <span>Delimitador</span>
                <select
                  value={delimiterOption}
                  onChange={(event) => setDelimiterOption(event.target.value as DelimiterOption)}
                >
                  <option value="double-space">Espaco duplo</option>
                  <option value="semicolon">Ponto e virgula (;)</option>
                  <option value="comma">Virgula (,)</option>
                  <option value="tab">Tabulacao</option>
                  <option value="custom">Expressao customizada</option>
                </select>
              </label>

              <label className="field">
                <span>Regex customizada</span>
                <input
                  type="text"
                  value={customDelimiter}
                  onChange={(event) => setCustomDelimiter(event.target.value)}
                  placeholder="Ex.: \\|"
                  disabled={delimiterOption !== "custom"}
                />
              </label>
            </>
          ) : (
            <div className="field actions-field">
              <span>Colunas visuais</span>
              <div className="inline-actions">
                <button
                  type="button"
                  className="icon-button"
                  onClick={addColumnGuide}
                  disabled={!pdfDocument}
                  aria-label="Adicionar separador de coluna"
                  title="Adicionar separador de coluna"
                >
                  <svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M12 4v16M8 12h8" />
                    <path d="M5 4v16M19 4v16" />
                  </svg>
                </button>
                <button
                  type="button"
                  className="icon-button"
                  onClick={removeColumnGuide}
                  disabled={!pdfDocument || guides.length === 0}
                  aria-label="Remover separador de coluna"
                  title="Remover separador de coluna"
                >
                  <svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M8 12h8" />
                    <path d="M5 4v16M19 4v16" />
                  </svg>
                </button>
                <small>{guides.length + 1} colunas</small>
              </div>
            </div>
          )}
        </div>

        <div className="toolbar">
          <div className="page-nav">
            <button
              type="button"
              className="icon-button"
              disabled={!pdfDocument || pageNumber <= 1}
              onClick={() => setPageNumber((page) => page - 1)}
              aria-label="Pagina anterior"
              title="Pagina anterior"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M14 6l-6 6 6 6" />
              </svg>
            </button>
            <span>
              Pagina {pageNumber} / {pageCount || 0}
            </span>
            <button
              type="button"
              className="icon-button"
              disabled={!pdfDocument || pageNumber >= pageCount}
              onClick={() => setPageNumber((page) => page + 1)}
              aria-label="Proxima pagina"
              title="Proxima pagina"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M10 6l6 6-6 6" />
              </svg>
            </button>
          </div>

          <button
            type="button"
            className="icon-button export-button"
            onClick={exportToXlsx}
            disabled={tableData.length === 0}
            aria-label="Exportar XLSX"
            title="Exportar XLSX"
          >
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path d="M12 3v12" />
              <path d="M8 11l4 4 4-4" />
              <path d="M4 19h16" />
            </svg>
          </button>
        </div>

        <p className="status-line">{status}</p>
      </section>

      <section className="workspace">
        <article className="panel viewer-panel">
          <div className="panel-heading">
            <h2>Pre-visualizacao</h2>
            <div className="zoom-controls">
              <button
                type="button"
                className="icon-button"
                onClick={() => changePdfZoom(-ZOOM_STEP)}
                disabled={pdfZoom <= MIN_ZOOM}
                aria-label="Zoom out PDF"
                title="Zoom out PDF"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <circle cx="11" cy="11" r="6" />
                  <path d="M20 20l-4-4" />
                  <path d="M8 11h6" />
                </svg>
              </button>
              <span className="zoom-label">{Math.round(pdfZoom * 100)}%</span>
              <button
                type="button"
                className="icon-button"
                onClick={() => changePdfZoom(ZOOM_STEP)}
                disabled={pdfZoom >= MAX_ZOOM}
                aria-label="Zoom in PDF"
                title="Zoom in PDF"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <circle cx="11" cy="11" r="6" />
                  <path d="M20 20l-4-4" />
                  <path d="M8 11h6M11 8v6" />
                </svg>
              </button>
            </div>
          </div>

          <div className="canvas-scroller">
            <div
              className="canvas-stage"
              ref={stageRef}
              style={{ width: canvasSize.width || undefined, height: canvasSize.height || undefined }}
            >
              <canvas ref={canvasRef} />
              {extractMode === "guided" && canvasSize.width > 0 && (
                <div className="guides-layer">
                  {guides.map((guidePosition, index) => (
                    <button
                      key={`${guidePosition}-${index}`}
                      type="button"
                      className={`guide-handle ${draggingGuideIndex === index ? "dragging" : ""}`}
                      style={{ left: `${guidePosition}px` }}
                      onMouseDown={(event) => {
                        event.preventDefault();
                        setDraggingGuideIndex(index);
                      }}
                      onTouchStart={(event) => {
                        event.preventDefault();
                        setDraggingGuideIndex(index);
                      }}
                      aria-label={`Separador ${index + 1}`}
                      title={`Separador ${index + 1}`}
                    />
                  ))}
                </div>
              )}
            </div>
          </div>
        </article>

        <article className="panel sheet-panel">
          <div className="panel-heading">
            <h2>Grade de dados</h2>
            <div className="zoom-controls">
              <button
                type="button"
                className="icon-button"
                onClick={() => changeSheetZoom(-ZOOM_STEP)}
                disabled={sheetZoom <= MIN_ZOOM}
                aria-label="Zoom out grade"
                title="Zoom out grade"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <circle cx="11" cy="11" r="6" />
                  <path d="M20 20l-4-4" />
                  <path d="M8 11h6" />
                </svg>
              </button>
              <span className="zoom-label">{Math.round(sheetZoom * 100)}%</span>
              <button
                type="button"
                className="icon-button"
                onClick={() => changeSheetZoom(ZOOM_STEP)}
                disabled={sheetZoom >= MAX_ZOOM}
                aria-label="Zoom in grade"
                title="Zoom in grade"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <circle cx="11" cy="11" r="6" />
                  <path d="M20 20l-4-4" />
                  <path d="M8 11h6M11 8v6" />
                </svg>
              </button>
            </div>
          </div>

          <div className="sheet-actions">
            <button
              type="button"
              className="icon-button"
              onClick={addSheetRow}
              aria-label="Adicionar linha"
              title="Adicionar linha"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M4 8h10M4 12h10M4 16h10" />
                <path d="M18 9v6M15 12h6" />
              </svg>
            </button>
            <button
              type="button"
              className="icon-button"
              onClick={removeSheetRow}
              disabled={rowCount === 0}
              aria-label="Remover linha"
              title="Remover linha"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M4 8h10M4 12h10M4 16h10" />
                <path d="M15 12h6" />
              </svg>
            </button>
            <button
              type="button"
              className="icon-button"
              onClick={addSheetColumn}
              aria-label="Adicionar coluna"
              title="Adicionar coluna"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <rect x="3" y="5" width="12" height="14" rx="2" />
                <path d="M18 8v8M14 12h8" />
              </svg>
            </button>
            <button
              type="button"
              className="icon-button"
              onClick={removeSheetColumn}
              disabled={columnCount === 0}
              aria-label="Remover coluna"
              title="Remover coluna"
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <rect x="3" y="5" width="12" height="14" rx="2" />
                <path d="M14 12h8" />
              </svg>
            </button>
            <small>
              {rowCount} linhas x {columnCount} colunas
            </small>
          </div>

          {tableData.length === 0 ? (
            <p className="empty-state">Os dados extraidos aparecerao aqui apos carregar um PDF.</p>
          ) : (
            <div
              className="sheet-wrapper"
              ref={sheetWrapperRef}
              tabIndex={0}
              onKeyDown={handleSheetKeyDown}
              onMouseDown={() => {
                if (!editingCell) {
                  focusSheet();
                }
              }}
            >
              <div
                className="sheet-zoom-stage"
                style={{
                  transform: `scale(${sheetZoom})`,
                  transformOrigin: "top left",
                  width: `${100 / sheetZoom}%`
                }}
              >
                <table className="sheet">
                  <thead>
                    <tr>
                      <th>#</th>
                      {headers.map((header, columnIndex) => (
                        <th key={header} className={selectedColumnIndex === columnIndex ? "header-selected" : ""}>
                          <div
                            className="sheet-header-trigger"
                            role="button"
                            tabIndex={0}
                            onClick={() => {
                              setSelectedColumnIndex(columnIndex);
                              setSelectedRowIndex(null);
                              setSelectedCell(null);
                              setEditingCell(null);
                              focusSheet();
                            }}
                            onKeyDown={(event) =>
                              runOnEnterOrSpace(event, () => {
                                setSelectedColumnIndex(columnIndex);
                                setSelectedRowIndex(null);
                                setSelectedCell(null);
                                setEditingCell(null);
                                focusSheet();
                              })
                            }
                          >
                            {header}
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {tableData.map((row, rowIndex) => (
                      <tr key={`row-${rowIndex}`}>
                        <th className={selectedRowIndex === rowIndex ? "header-selected" : ""}>
                          <div
                            className="sheet-header-trigger"
                            role="button"
                            tabIndex={0}
                            onClick={() => {
                              setSelectedRowIndex(rowIndex);
                              setSelectedColumnIndex(null);
                              setSelectedCell(null);
                              setEditingCell(null);
                              focusSheet();
                            }}
                            onKeyDown={(event) =>
                              runOnEnterOrSpace(event, () => {
                                setSelectedRowIndex(rowIndex);
                                setSelectedColumnIndex(null);
                                setSelectedCell(null);
                                setEditingCell(null);
                                focusSheet();
                              })
                            }
                          >
                            {rowIndex + 1}
                          </div>
                        </th>

                        {row.map((cell, columnIndex) => (
                          <td
                            key={`cell-${rowIndex}-${columnIndex}`}
                            className={[
                              selectedRowIndex === rowIndex ? "row-selected" : "",
                              selectedColumnIndex === columnIndex ? "column-selected" : "",
                              selectedCell?.row === rowIndex && selectedCell.column === columnIndex ? "cell-selected" : "",
                              editingCell?.row === rowIndex && editingCell.column === columnIndex ? "cell-editing" : ""
                            ]
                              .join(" ")
                              .trim()}
                            onClick={() => activateCellSelection(rowIndex, columnIndex)}
                            onDoubleClick={() => startCellEditing(rowIndex, columnIndex)}
                          >
                            {editingCell?.row === rowIndex && editingCell.column === columnIndex ? (
                              <input
                                ref={(node) => {
                                  editorRef.current = node;
                                }}
                                type="text"
                                value={cell}
                                onChange={(event) => updateCell(rowIndex, columnIndex, event.target.value)}
                                onBlur={() => {
                                  setEditingCell(null);
                                  focusSheet();
                                }}
                                onKeyDown={(event) => {
                                  if (event.key === "Enter" || event.key === "Escape") {
                                    event.preventDefault();
                                    setEditingCell(null);
                                    focusSheet();
                                  }
                                }}
                              />
                            ) : (
                              <span>{cell || "\u00A0"}</span>
                            )}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </article>
      </section>
    </main>
  );
}
