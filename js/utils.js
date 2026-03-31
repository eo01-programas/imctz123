(function () {
  function normalizeKey(value) {
    return String(value ?? "")
      .trim()
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  }

  function evaluateFormula(value) {
    const rawValue = String(value ?? "").trim();
    if (!rawValue) return 0;
    
    if (rawValue.startsWith("+") || rawValue.startsWith("=")) {
      try {
        const sanitized = rawValue.substring(1).replace(/[^0-9+\-*/().]/g, '');
        if (!sanitized) return 0;
        const result = new Function(`return ${sanitized}`)();
        return Number.isFinite(result) ? result : 0;
      } catch (e) {
        return safeNumber(rawValue);
      }
    }
    
    return safeNumber(rawValue);
  }

  function safeNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }

    const rawValue = String(value ?? "")
      .trim()
      .replace(/\s+/g, "");

    if (!rawValue) {
      return 0;
    }

    let sign = "";
    let normalized = rawValue;

    if (/^[+-]/.test(normalized)) {
      sign = normalized.charAt(0);
      normalized = normalized.slice(1);
    }

    const commaCount = (normalized.match(/,/g) || []).length;
    const dotCount = (normalized.match(/\./g) || []).length;

    if (commaCount > 0 && dotCount > 0) {
      const decimalSeparator = normalized.lastIndexOf(",") > normalized.lastIndexOf(".") ? "," : ".";
      const thousandsSeparator = decimalSeparator === "," ? "." : ",";

      normalized = normalized.split(thousandsSeparator).join("");
      normalized = normalized.replace(decimalSeparator, ".");
    } else if (commaCount > 1) {
      const parts = normalized.split(",");
      const looksLikeThousands = parts.slice(1).every((group) => /^\d{3}$/.test(group));

      if (looksLikeThousands) {
        normalized = parts.join("");
      } else {
        const decimalPart = parts.pop();
        const integerPart = parts.join("");
        normalized = `${integerPart}.${decimalPart}`;
      }
    } else if (dotCount > 1) {
      const parts = normalized.split(".");
      const looksLikeThousands = parts.slice(1).every((group) => /^\d{3}$/.test(group));

      if (looksLikeThousands) {
        normalized = parts.join("");
      } else {
        const decimalPart = parts.pop();
        const integerPart = parts.join("");
        normalized = `${integerPart}.${decimalPart}`;
      }
    } else if (commaCount === 1) {
      normalized = normalized.replace(",", ".");
    }

    normalized = normalized.replace(/[^0-9.]/g, "");
    if (!normalized) {
      return 0;
    }

    const parsed = Number(`${sign === "-" ? "-" : ""}${normalized}`);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function optionalNumber(value) {
    const rawValue = String(value ?? "").trim();
    return rawValue ? safeNumber(rawValue) : "";
  }

  function formatNumber(value, decimals = 3) {
    return safeNumber(value).toLocaleString("es-PE", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    });
  }

  function formatInteger(value) {
    const digits = String(value ?? "").replace(/\D/g, "").replace(/^0+(?=\d)/, "");

    if (!digits) {
      return "";
    }

    return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  function escapeCsvField(value) {
    const text = String(value ?? "");
    if (!/[",\n]/.test(text)) {
      return text;
    }

    return `"${text.replace(/"/g, '""')}"`;
  }

  function serializeCosturaRows(rows) {
    const matrix = [
      APP_CONFIG.costuraCsvHeaders,
      ...rows.map((row) => [
        row.codigo ?? "",
        row.bloque ?? "",
        row.operaciones ?? "",
        safeNumber(row.tiempoEstimado),
        row.tipoMaq ?? "",
        safeNumber(row.proteccion || APP_CONFIG.defaultProtection),
        row.tipoPta ?? "",
        safeNumber(row.tiempoMaq),
        safeNumber(row.tiempoManual),
        safeNumber(row.tiempoCotizacion),
      ]),
    ];

    return matrix
      .map((line) => line.map(escapeCsvField).join(","))
      .join("\n");
  }

  function serializeCorteRows(rows) {
    const matrix = [
      APP_CONFIG.corteCsvHeaders,
      ...rows.map((row) => [
        row.operaciones ?? "",
        row.tiempoEstimadoCorte ?? "",
        row.tiempoEstimadoHabilitado ?? "",
        safeNumber(row.proteccion || APP_CONFIG.defaultProtection),
        row.area ?? "",
        optionalNumber(row.tiempoCorte),
        optionalNumber(row.tiempoHab),
        optionalNumber(row.tiempoCotizacion),
      ]),
    ];

    return matrix
      .map((line) => line.map(escapeCsvField).join(","))
      .join("\n");
  }

  function serializeAcabadoRows(rows) {
    const matrix = [
      APP_CONFIG.acabadoCsvHeaders,
      ...rows.map((row) => [
        row.operaciones ?? "",
        optionalNumber(row.tiempoEstimado),
        safeNumber(row.proteccion || APP_CONFIG.defaultProtection),
        optionalNumber(row.tiempoCotizacion),
      ]),
    ];

    return matrix
      .map((line) => line.map(escapeCsvField).join(","))
      .join("\n");
  }

  function parseCsvText(csvText) {
    const text = String(csvText ?? "");

    if (!text.trim()) {
      return [];
    }

    const rows = [];
    let currentRow = [];
    let currentValue = "";
    let insideQuotes = false;

    for (let index = 0; index < text.length; index += 1) {
      const char = text[index];
      const nextChar = text[index + 1];

      if (insideQuotes) {
        if (char === '"' && nextChar === '"') {
          currentValue += '"';
          index += 1;
        } else if (char === '"') {
          insideQuotes = false;
        } else {
          currentValue += char;
        }
        continue;
      }

      if (char === '"') {
        insideQuotes = true;
        continue;
      }

      if (char === ",") {
        currentRow.push(currentValue);
        currentValue = "";
        continue;
      }

      if (char === "\n") {
        currentRow.push(currentValue);
        rows.push(currentRow);
        currentRow = [];
        currentValue = "";
        continue;
      }

      if (char !== "\r") {
        currentValue += char;
      }
    }

    currentRow.push(currentValue);
    rows.push(currentRow);

    return rows.filter((row) => row.some((cell) => String(cell).trim() !== ""));
  }

  function parseCosturaCsv(csvText) {
    const matrix = parseCsvText(csvText);

    if (matrix.length <= 1) {
      return [];
    }

    const headerMap = matrix[0].map((header) => normalizeKey(header));

    return matrix
      .slice(1)
      .filter((row) => row.some((cell) => String(cell).trim() !== ""))
      .map((row) => {
        const record = {};

        headerMap.forEach((header, index) => {
          const value = row[index] ?? "";

          switch (header) {
            case "CODIGO":
              record.codigo = value;
              break;
            case "BLOQUE":
              record.bloque = value;
              break;
            case "OPERACIONES":
              record.operaciones = value;
              break;
            case "TIEMPOS ESTIMADO":
            case "TIEMPO ESTIMADO":
              record.tiempoEstimado = safeNumber(value);
              break;
            case "TIPO MAQ":
              record.tipoMaq = value;
              break;
            case "% PROTECCION":
              record.proteccion = safeNumber(value);
              break;
            case "TIPO PTA":
              record.tipoPta = value;
              break;
            case "TIEMPOS MAQ C/PROTECC.":
            case "TIEMPOS MAQ C/PROTECC":
              record.tiempoMaq = safeNumber(value);
              break;
            case "TIEMPOS MANUAL C/PROTECC.":
            case "TIEMPOS MANUAL C/PROTECC":
              record.tiempoManual = safeNumber(value);
              break;
            case "TIEMPOS COTIZACION":
              record.tiempoCotizacion = safeNumber(value);
              break;
            default:
              break;
          }
        });

        return record;
      });
  }

  function parseCorteCsv(csvText) {
    const matrix = parseCsvText(csvText);

    if (matrix.length <= 1) {
      return [];
    }

    const headerMap = matrix[0].map((header) => normalizeKey(header));

    return matrix
      .slice(1)
      .filter((row) => row.some((cell) => String(cell).trim() !== ""))
      .map((row) => {
        const record = {};
        headerMap.forEach((header, index) => {
          const value = row[index] ?? "";
          switch (header) {
            case "OPERACIONES": record.operaciones = value; break;
            case "TIEMPOS ESTIMADO CORTE":
            case "TIEMPO ESTIMADO CORTE":
              record.tiempoEstimadoCorte = String(value).trim();
              break;
            case "TIEMPOS ESTIMADO HABILITADO":
            case "TIEMPO ESTIMADO HABILITADO":
              record.tiempoEstimadoHabilitado = String(value).trim();
              break;
            case "% PROTECCION": record.proteccion = safeNumber(value); break;
            case "AREA": record.area = value; break;
            case "TIEMPOS CORTE C/PROTECC.":
            case "TIEMPOS CORTE C/PROTECC":
            case "TIEMPO CORTE":
              record.tiempoCorte = optionalNumber(value);
              break;
            case "TIEMPOS HAB C/PROTECC.":
            case "TIEMPOS HAB C/PROTECC":
            case "TIEMPO HAB.":
              record.tiempoHab = optionalNumber(value);
              break;
            case "TIEMPOS COTIZACION":
              record.tiempoCotizacion = optionalNumber(value);
              break;
          }
        });
        return record;
      });
  }

  function parseAcabadoCsv(csvText) {
    const matrix = parseCsvText(csvText);

    if (matrix.length <= 1) {
      return [];
    }

    const headerMap = matrix[0].map((header) => normalizeKey(header));

    return matrix
      .slice(1)
      .filter((row) => row.some((cell) => String(cell).trim() !== ""))
      .map((row) => {
        const record = {};
        headerMap.forEach((header, index) => {
          const value = row[index] ?? "";
          switch (header) {
            case "OPERACIONES": record.operaciones = value; break;
            case "TIEMPOS ESTIMADO":
            case "TIEMPO ESTIMADO":
              record.tiempoEstimado = optionalNumber(value);
              break;
            case "% PROTECCION": record.proteccion = safeNumber(value); break;
            case "TIEMPOS COTIZACION":
              record.tiempoCotizacion = optionalNumber(value);
              break;
          }
        });
        return record;
      });
  }

  function parseVersionDate(record) {
    if (!record || !record.savedAt) {
      return null;
    }

    const date = new Date(record.savedAt);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function formatVersionDate(record) {
    const date = parseVersionDate(record);

    if (!date) {
      return "Sin fecha";
    }

    return date.toLocaleDateString("es-PE", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });
  }

  function formatVersionTime(record) {
    const date = parseVersionDate(record);

    if (!date) {
      return "Sin hora";
    }

    return date.toLocaleTimeString("es-PE", {
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function formatVersionMeta(record) {
    const dateLabel = formatVersionDate(record);
    const timeLabel = formatVersionTime(record);

    if (dateLabel === "Sin fecha") {
      return dateLabel;
    }

    if (timeLabel === "Sin hora") {
      return dateLabel;
    }

    return `${dateLabel} ${timeLabel}`;
  }

  window.AppUtils = {
    normalizeKey,
    evaluateFormula,
    safeNumber,
    formatNumber,
    formatInteger,
    serializeCosturaRows,
    serializeCorteRows,
    serializeAcabadoRows,
    parseCosturaCsv,
    parseCorteCsv,
    parseAcabadoCsv,
    formatVersionDate,
    formatVersionTime,
    formatVersionMeta,
  };
})();
