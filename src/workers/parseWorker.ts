import * as XLSX from "xlsx";

type ResultMessage = {
  type: "result";
  countryCustomerCounts: { [country: string]: number };
  dateCounts: { [date: string]: number };
};

type ProgressMessage = { type: "progress"; percent: number };

self.onmessage = async (ev: MessageEvent) => {
  try {
    const arrayBuffer = ev.data as ArrayBuffer;
    const data = new Uint8Array(arrayBuffer);
    const wb = XLSX.read(data, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];

    // 將工作表轉成二維陣列（含表頭）
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length === 0) {
      const res: ResultMessage = {
        type: "result",
        countryCustomerCounts: {},
        dateCounts: {},
      };
      (self as any).postMessage(res);
      return;
    }

    // 先發一個初始進度，避免前端長時間無回饋
    (self as any).postMessage({ type: "progress", percent: 1 });

    const header = rows[0].map((h: any) => String(h).trim());
    const idx = {
      StockCode: header.indexOf("StockCode"),
      Description: header.indexOf("Description"),
      Quantity: header.indexOf("Quantity"),
      InvoiceDate: header.indexOf("InvoiceDate"),
      UnitPrice: header.indexOf("UnitPrice"),
      CustomerID: header.indexOf("CustomerID"),
      Country: header.indexOf("Country"),
    };

    const countryCustomerSets = new Map<string, Set<string>>();
    const dateCounts = new Map<string, number>();

    const total = rows.length - 1;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      // defensive checks
      const country =
        idx.Country >= 0 ? String(r[idx.Country] ?? "").trim() : "";
      const customer =
        idx.CustomerID >= 0 ? String(r[idx.CustomerID] ?? "").trim() : "";
      const invoiceRaw = idx.InvoiceDate >= 0 ? r[idx.InvoiceDate] : "";

      if (country) {
        const set = countryCustomerSets.get(country) ?? new Set<string>();
        if (customer) set.add(customer);
        countryCustomerSets.set(country, set);
      }

      // 將日期正規化為 YYYY-MM-DD，支援 Excel 數字日期與字串日期
      let day = "";
      if (
        invoiceRaw !== undefined &&
        invoiceRaw !== null &&
        invoiceRaw !== ""
      ) {
        if (typeof invoiceRaw === "number") {
          try {
            const parse_date_code =
              (XLSX as any).SSF && (XLSX as any).SSF.parse_date_code;
            if (parse_date_code) {
              const pd: any = parse_date_code(invoiceRaw);
              if (pd && pd.y) {
                const yyyy = pd.y;
                const mm = String(pd.m).padStart(2, "0");
                const dd = String(pd.d).padStart(2, "0");
                day = `${yyyy}-${mm}-${dd}`;
              }
            }
            if (!day) {
              // fallback: Excel serial -> JS date，常見 offset 25569
              const jsDate = new Date(
                Math.round((invoiceRaw - 25569) * 86400 * 1000)
              );
              if (!isNaN(jsDate.getTime())) {
                const yyyy = jsDate.getFullYear();
                const mm = String(jsDate.getMonth() + 1).padStart(2, "0");
                const dd = String(jsDate.getDate()).padStart(2, "0");
                day = `${yyyy}-${mm}-${dd}`;
              }
            }
          } catch (e) {
            // ignore and try string path below
          }
        }

        if (!day) {
          try {
            const d = new Date(invoiceRaw as any);
            if (!isNaN(d.getTime())) {
              const yyyy = d.getFullYear();
              const mm = String(d.getMonth() + 1).padStart(2, "0");
              const dd = String(d.getDate()).padStart(2, "0");
              day = `${yyyy}-${mm}-${dd}`;
            } else {
              // fallback: 嘗試 yyyy/mm/dd 字串
              const s = String(invoiceRaw);
              const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
              if (m)
                day = `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(
                  m[3]
                ).padStart(2, "0")}`;
            }
          } catch (e) {
            // ignore
          }
        }
      }
      if (day) dateCounts.set(day, (dateCounts.get(day) ?? 0) + 1);

      // progress report every 5% or on last
      if (
        i % Math.max(1, Math.floor(total / 20)) === 0 ||
        i === rows.length - 1
      ) {
        const percent = Math.round((i / total) * 100);
        const pm: ProgressMessage = { type: "progress", percent };
        (self as any).postMessage(pm);
      }
    }

    const countryCustomerCounts: { [country: string]: number } = {};
    for (const [country, set] of countryCustomerSets.entries())
      countryCustomerCounts[country] = set.size;

    const dateCountsObj: { [date: string]: number } = {};
    for (const [d, c] of dateCounts.entries()) dateCountsObj[d] = c;

    const res: ResultMessage = {
      type: "result",
      countryCustomerCounts,
      dateCounts: dateCountsObj,
    };
    (self as any).postMessage(res);
  } catch (err) {
    (self as any).postMessage({ type: "error", message: String(err) });
  }
};

export {};
