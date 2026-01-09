import React, { useEffect, useRef, useState } from "react";
import { Pie, Bar } from "react-chartjs-2";
import Chart from "chart.js/auto";
import * as XLSX from "xlsx";

type ParseProgress = { type: "progress"; percent: number };
type ParseResult = {
  type: "result";
  countryCustomerCounts: Record<string, number>;
  dateCounts: Record<string, number>;
};

const UploadPage: React.FC = () => {
  const [progress, setProgress] = useState<number | null>(null);
  const [countryData, setCountryData] = useState<{
    labels: string[];
    data: number[];
  } | null>(null);
  const [dateData, setDateData] = useState<{
    labels: string[];
    data: number[];
  } | null>(null);
  const [monthlyData, setMonthlyData] = useState<{
    labels: string[];
    data: number[];
  } | null>(null);
  const [selectedMonthIndex, setSelectedMonthIndex] = useState<number | null>(
    null
  );
  const [fileName, setFileName] = useState<string | null>(null);
  const workerRef = useRef<Worker | null>(null);
  const [error, setError] = useState<string | null>(null);
  const workerAvailableRef = useRef<boolean>(false); // 用來記錄 worker 是否可用，失敗時退回主執行緒解析

  useEffect(() => {
    // instantiate worker (Vite: use new URL(..., import.meta.url) and set type)
    try {
      workerRef.current = new Worker(
        new URL("../workers/parseWorker.ts", import.meta.url),
        { type: "module" }
      );
      workerAvailableRef.current = true;
      workerRef.current.onmessage = (ev: MessageEvent) => {
        const msg = ev.data as ParseProgress | ParseResult | any;
        if (msg.type === "progress")
          setProgress((msg as ParseProgress).percent);
        else if (msg.type === "result") {
          const res = msg as ParseResult;
          console.log(res);
          const countryEntries = Object.entries(res.countryCustomerCounts).sort(
            (a, b) => b[1] - a[1]
          );
          setCountryData({
            labels: countryEntries.map((e) => e[0]),
            data: countryEntries.map((e) => e[1]),
          });

          const dateEntries = Object.entries(res.dateCounts).sort((a, b) =>
            a[0].localeCompare(b[0])
          );
          const dateLabels = dateEntries.map((e) => e[0]);
          const dateCountsArr = dateEntries.map((e) => e[1]);
          setDateData({ labels: dateLabels, data: dateCountsArr });

          // build monthly aggregates YYYY-MM
          const monthMap = new Map<string, number>();
          for (let i = 0; i < dateLabels.length; i++) {
            const m = dateLabels[i].slice(0, 7);
            monthMap.set(m, (monthMap.get(m) ?? 0) + dateCountsArr[i]);
          }
          const monthArr = Array.from(monthMap.entries()).sort((a, b) =>
            a[0].localeCompare(b[0])
          );
          setMonthlyData({
            labels: monthArr.map((e) => e[0]),
            data: monthArr.map((e) => e[1]),
          });
          setSelectedMonthIndex(null);
          setProgress(100);
        } else if (msg.type === "error") {
          console.error("Worker error", msg.message);
        }
      };
      workerRef.current.onerror = (err) => {
        console.error("Worker runtime error", err);
        setError("Worker runtime error. Will fallback to main-thread parsing.");
        workerAvailableRef.current = false;
      };
    } catch (err) {
      console.error("Failed to create worker", err);
      setError(String(err));
      workerAvailableRef.current = false;
    }
    return () => {
      workerRef.current?.terminate();
    };
  }, []);

  const handleFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setProgress(0);
    setCountryData(null);
    setDateData(null);

    const reader = new FileReader();
    reader.onload = (ev) => {
      const arrayBuffer = ev.target?.result as ArrayBuffer;
      if (!arrayBuffer) return;
      // try to send to worker (transferable), fallback to main-thread if it fails
      try {
        if (workerAvailableRef.current && workerRef.current) {
          workerRef.current.postMessage(arrayBuffer, [arrayBuffer]);
          return;
        }
      } catch (err) {
        console.warn(
          "Posting to worker failed, will parse on main thread",
          err
        );
      }

      // fallback: parse on main thread
      try {
        parseOnMainThread(arrayBuffer);
      } catch (err) {
        console.error("Main-thread parse failed", err);
        setError(String(err));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 主執行緒回退解析（與 worker 同步邏輯），避免 worker 建立失敗時卡住
  const parseOnMainThread = (arrayBuffer: ArrayBuffer) => {
    const data = new Uint8Array(arrayBuffer);
    const wb = XLSX.read(data, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (!rows || rows.length === 0) {
      setError("檔案沒有資料");
      return;
    }

    const header = rows[0].map((h: any) => String(h).trim());
    const idx = {
      InvoiceDate: header.indexOf("InvoiceDate"),
      CustomerID: header.indexOf("CustomerID"),
      Country: header.indexOf("Country"),
    };

    const countryCustomerSets = new Map<string, Set<string>>();
    const dateCounts = new Map<string, number>();

    const total = Math.max(1, rows.length - 1);
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
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
          const d = new Date(invoiceRaw as any);
          if (!isNaN(d.getTime())) {
            const yyyy = d.getFullYear();
            const mm = String(d.getMonth() + 1).padStart(2, "0");
            const dd = String(d.getDate()).padStart(2, "0");
            day = `${yyyy}-${mm}-${dd}`;
          } else {
            const s = String(invoiceRaw);
            const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
            if (m)
              day = `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(
                m[3]
              ).padStart(2, "0")}`;
          }
        }
      }
      if (day) dateCounts.set(day, (dateCounts.get(day) ?? 0) + 1);

      if (
        i % Math.max(1, Math.floor(total / 20)) === 0 ||
        i === rows.length - 1
      ) {
        setProgress(Math.round((i / total) * 100));
      }
    }

    const countryEntries = Array.from(countryCustomerSets.entries())
      .map(([k, v]) => [k, v.size] as [string, number])
      .sort((a, b) => b[1] - a[1]);
    setCountryData({
      labels: countryEntries.map((e) => e[0]),
      data: countryEntries.map((e) => e[1]),
    });

    const dateEntries = Array.from(dateCounts.entries()).sort((a, b) =>
      a[0].localeCompare(b[0])
    );
    const dateLabels = dateEntries.map((e) => e[0]);
    const dateCountsArr = dateEntries.map((e) => e[1]);
    setDateData({ labels: dateLabels, data: dateCountsArr });

    const monthMap = new Map<string, number>();
    for (let i = 0; i < dateLabels.length; i++) {
      const m = dateLabels[i].slice(0, 7);
      monthMap.set(m, (monthMap.get(m) ?? 0) + dateCountsArr[i]);
    }
    const monthArr = Array.from(monthMap.entries()).sort((a, b) =>
      a[0].localeCompare(b[0])
    );
    setMonthlyData({
      labels: monthArr.map((e) => e[0]),
      data: monthArr.map((e) => e[1]),
    });
    setSelectedMonthIndex(null);
    setProgress(100);
  };

  return (
    <div style={{ padding: 16 }}>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 12,
          alignItems: "center",
          justifyContent: "center",
          position: "absolute",
          right: 20,
          top: 20,
        }}
      >
        <label
          style={{
            padding: "8px 12px",
            background: "#111827",
            color: "white",
            borderRadius: 6,
            cursor: "pointer",
          }}
        >
          選取 Excel 檔案
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFile}
            style={{ display: "none" }}
          />
        </label>
        <div>{fileName ?? "尚未選取檔案"}</div>
        {progress !== null && (
          <div style={{ width: 150 }}>
            <div style={{ fontSize: 12, color: "#374151" }}>
              進度：{progress}% {progress < 100 ? "（解析中）" : "完成"}
            </div>
            <div
              style={{
                marginTop: 4,
                height: 8,
                background: "#e5e7eb",
                borderRadius: 999,
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  width: `${Math.max(1, Math.min(100, progress))}%`,
                  height: "100%",
                  background: "linear-gradient(90deg, #3b82f6, #2563eb)",
                  transition: "width 0.2s ease",
                }}
              />
            </div>
          </div>
        )}
        {error && <div style={{ color: "#dc2626" }}>錯誤：{error}</div>}
      </div>

      <div
        style={{
          display: "flex",
          gap: 24,
          marginTop: 24,
          justifyContent: "center",
        }}
      >
        <section
          style={{
            width: "100%",
          }}
        >
          <h3>客戶數量按國家（圓餅圖）</h3>
          {countryData ? (
            <Pie
              data={{
                labels: countryData.labels,
                datasets: [
                  {
                    data: countryData.data,
                    backgroundColor: countryData.labels.map(
                      (_, i) => `hsl(${(i * 47) % 360} 70% 50%)`
                    ),
                  },
                ],
              }}
            />
          ) : (
            <div style={{ color: "#6b7280" }}>等待處理或尚無資料</div>
          )}
        </section>

        <section
          style={{
            width: "100%",
          }}
        >
          <h3>每日發票數（長條圖）</h3>
          {dateData ? (
            (() => {
              const total = dateData.data.reduce((s, v) => s + v, 0) || 1;

              // 資料量過大時先進行彙總，避免長軸過長難閱讀
              let labels = dateData.labels.slice();
              let counts = dateData.data.slice();
              let aggNote: string | null = null;
              const labelMap = new Map<string, number>();
              for (let i = 0; i < labels.length; i++) {
                labelMap.set(labels[i], counts[i]);
              }
              labels = Array.from(labelMap.keys()).sort((a, b) =>
                a.localeCompare(b)
              );
              counts = labels.map((lbl) => labelMap.get(lbl) || 0);

              if (labels.length > 90) {
                // aggregate by month (YYYY-MM)
                const m = new Map<string, number>();
                for (let i = 0; i < labels.length; i++) {
                  const key = labels[i].slice(0, 7); // YYYY-MM
                  m.set(key, (m.get(key) ?? 0) + counts[i]);
                }
                const arr = Array.from(m.entries()).sort((a, b) =>
                  a[0].localeCompare(b[0])
                );
                labels = arr.map((e) => e[0]);
                counts = arr.map((e) => e[1]);
                aggNote = "（已按月彙總）";
              } else if (labels.length > 30) {
                // show recent 30 days
                const n = 30;
                const start = Math.max(0, labels.length - n);
                labels = labels.slice(start);
                counts = counts.slice(start);
                aggNote = `（僅顯示最近 ${labels.length} 筆）`;
              }

              const percentData = counts.map((v) =>
                Number(((v / total) * 100).toFixed(2))
              );

              // prepare months list (from monthlyData if available, else derive from labels)
              const months =
                monthlyData && monthlyData.labels.length > 0
                  ? monthlyData.labels
                  : Array.from(new Set(labels.map((l) => l.slice(0, 7)))).sort(
                      (a, b) => a.localeCompare(b)
                    );

              // monthly overview (no month selected)
              if (selectedMonthIndex === null) {
                const monthsCounts =
                  monthlyData && monthlyData.data.length === months.length
                    ? monthlyData.data
                    : months.map((m) => {
                        let s = 0;
                        for (let i = 0; i < labels.length; i++)
                          if (labels[i].startsWith(m)) s += counts[i];
                        return s;
                      });

                const monthsPercent = monthsCounts.map((v) =>
                  Number(((v / total) * 100).toFixed(2))
                );

                return (
                  <div>
                    {aggNote && (
                      <div style={{ color: "#6b7280", marginBottom: 8 }}>
                        {aggNote}
                      </div>
                    )}

                    <div
                      style={{
                        display: "flex",
                        gap: 8,
                        alignItems: "center",
                        marginBottom: 8,
                      }}
                    >
                      <label style={{ color: "#374151" }}>選擇月份：</label>
                      <select
                        onChange={(e) =>
                          setSelectedMonthIndex(Number(e.target.value))
                        }
                        defaultValue={-1}
                      >
                        <option value={-1} disabled>
                          —— 檢視指定月份 ——
                        </option>
                        {months.map((m, i) => (
                          <option key={m} value={i}>
                            {m}
                          </option>
                        ))}
                      </select>
                      <div style={{ marginLeft: 12, color: "#6b7280" }}>
                        月份總覽
                      </div>
                    </div>

                    <div style={{ maxHeight: 400, overflowY: "auto" }}>
                      <Bar
                        data={{
                          labels: months,
                          datasets: [
                            {
                              label: "Invoices (%)",
                              data: monthsPercent,
                              backgroundColor: "rgba(59,130,246,0.8)",
                            },
                          ],
                        }}
                        options={{
                          maintainAspectRatio: false,
                          scales: {
                            x: { ticks: { autoSkip: true, maxTicksLimit: 12 } },
                            y: {
                              beginAtZero: true,
                              ticks: { callback: (v: any) => `${v}%` },
                            },
                          },
                          plugins: {
                            tooltip: {
                              callbacks: {
                                label: (context: any) =>
                                  `${context.raw}% (${
                                    monthsCounts[context.dataIndex]
                                  } invoices)`,
                              },
                            },
                          },
                        }}
                        height={500}
                      />
                    </div>
                  </div>
                );
              }

              // month selected -> show daily breakdown for that month
              const selIdx = selectedMonthIndex as number;
              const selMonth = months[selIdx];
              const dailyLabels: string[] = [];
              const dailyCounts: number[] = [];
              for (let i = 0; i < dateData.labels.length; i++) {
                if (dateData.labels[i].startsWith(selMonth)) {
                  dailyLabels.push(dateData.labels[i]);
                  dailyCounts.push(dateData.data[i]);
                }
              }
              const dailyTotal = dailyCounts.reduce((s, v) => s + v, 0) || 1;
              const dailyPercent = dailyCounts.map((v) =>
                Number(((v / dailyTotal) * 100).toFixed(2))
              );

              return (
                <div>
                  <div
                    style={{
                      display: "flex",
                      gap: 8,
                      alignItems: "center",
                      marginBottom: 8,
                    }}
                  >
                    <button
                      onClick={() =>
                        setSelectedMonthIndex((i) => Math.max(0, (i ?? 0) - 1))
                      }
                      disabled={selIdx === 0}
                    >
                      上一個月
                    </button>
                    <button
                      onClick={() =>
                        setSelectedMonthIndex((i) =>
                          Math.min(months.length - 1, (i ?? 0) + 1)
                        )
                      }
                      disabled={selIdx === months.length - 1}
                    >
                      下一個月
                    </button>
                    <button
                      onClick={() => setSelectedMonthIndex(null)}
                      style={{ marginLeft: 12 }}
                    >
                      返回月份總覽
                    </button>
                    <div style={{ marginLeft: 12, color: "#374151" }}>
                      {selMonth} 的每日分佈
                    </div>
                  </div>

                  <div style={{ maxHeight: 400, overflowY: "auto" }}>
                    <Bar
                      data={{
                        labels: dailyLabels,
                        datasets: [
                          {
                            label: "Invoices (%)",
                            data: dailyPercent,
                            backgroundColor: "rgba(59,130,246,0.8)",
                          },
                        ],
                      }}
                      options={{
                        maintainAspectRatio: false,
                        scales: {
                          x: { ticks: { autoSkip: true, maxTicksLimit: 12 } },
                          y: {
                            beginAtZero: true,
                            ticks: { callback: (v: any) => `${v}%` },
                          },
                        },
                        plugins: {
                          tooltip: {
                            callbacks: {
                              label: (context: any) =>
                                `${context.raw}% (${
                                  dailyCounts[context.dataIndex]
                                } invoices)`,
                            },
                          },
                        },
                      }}
                      height={500}
                    />
                  </div>
                </div>
              );
            })()
          ) : (
            <div style={{ color: "#6b7280" }}>等待處理或尚無資料</div>
          )}
        </section>
      </div>
    </div>
  );
};

export default UploadPage;
