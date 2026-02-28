"use client";

import { useState, useRef, useCallback } from "react";

type Status = "idle" | "uploading" | "done" | "error";

interface ContractSummary {
  customer: string;
  merchant_id: string;
  receivable: number;
  income: number;
  bank_matched: number;
  invoice_matched: number;
  notes: string;
}

interface CalcResult {
  summary: ContractSummary[];
  total_receivable: number;
  total_income: number;
  contract_count: number;
  files: {
    lease: string;
    single: string;
    income: string;
  };
}

function formatMoney(n: number): string {
  if (n == null || isNaN(n)) return "—";
  return "¥" + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function downloadBase64(b64: string, filename: string) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// Get current year-month defaults
function getDefaultDates() {
  const now = new Date();
  const startYear = now.getFullYear();
  const startMonth = String(now.getMonth() + 1).padStart(2, "0");
  const endYear = now.getMonth() >= 11 ? now.getFullYear() + 1 : now.getFullYear();
  const endMonth = String(((now.getMonth() + 1) % 12) + 1).padStart(2, "0");
  // Actually let's do start = first day of current month, end = 11 months later
  const start = `${startYear}-${startMonth}`;
  const end = `${endYear}-${String(now.getMonth() + 2 > 12 ? 1 : now.getMonth() + 2).padStart(2, "0")}`;
  return { start, end };
}

export default function Home() {
  const defaults = getDefaultDates();
  const [file, setFile] = useState<File | null>(null);
  const [startMonth, setStartMonth] = useState(defaults.start);
  const [endMonth, setEndMonth] = useState(defaults.end);
  const [status, setStatus] = useState<Status>("idle");
  const [result, setResult] = useState<CalcResult | null>(null);
  const [errorMsg, setErrorMsg] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = useCallback((f: File) => {
    if (!f.name.endsWith(".xlsx")) {
      setErrorMsg("请上传 .xlsx 格式的文件");
      setStatus("error");
      return;
    }
    setFile(f);
    setStatus("idle");
    setErrorMsg("");
    setResult(null);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const f = e.dataTransfer.files[0];
      if (f) handleFileSelect(f);
    },
    [handleFileSelect]
  );

  const handleSubmit = async () => {
    if (!file) {
      setErrorMsg("请先选择 Excel 文件");
      setStatus("error");
      return;
    }
    if (!startMonth || !endMonth) {
      setErrorMsg("请填写起始月和结束月");
      setStatus("error");
      return;
    }

    setStatus("uploading");
    setErrorMsg("");
    setResult(null);

    const formData = new FormData();
    formData.append("file", file);
    formData.append("start", startMonth + "-01");
    formData.append("end", endMonth + "-01");

    try {
      const resp = await fetch("/api/calculate", {
        method: "POST",
        body: formData,
      });

      const data = await resp.json();

      if (!resp.ok || data.error) {
        setErrorMsg(data.error || `请求失败 (${resp.status})`);
        setStatus("error");
        return;
      }

      setResult(data);
      setStatus("done");
    } catch (e) {
      setErrorMsg("网络错误，请稍后重试。" + (e instanceof Error ? e.message : ""));
      setStatus("error");
    }
  };

  return (
    <main className="max-w-4xl mx-auto px-4 py-10">
      {/* Header */}
      <div className="mb-8 text-center">
        <h1 className="text-3xl font-bold text-gray-800 mb-2">租赁合同计算器</h1>
        <p className="text-gray-500 text-sm">
          上传租赁合同 Excel 数据，自动计算应收总额、收入总额，并生成月度明细报表
        </p>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-200 p-6 space-y-6">
        {/* Step 1: Download template */}
        <section>
          <h2 className="text-base font-semibold text-gray-700 mb-2">
            <span className="inline-block bg-blue-100 text-blue-700 rounded-full w-6 h-6 text-xs font-bold text-center leading-6 mr-2">1</span>
            下载数据填写模板
          </h2>
          <p className="text-sm text-gray-500 mb-3">
            模板包含 3 个工作表（合同数据、银行对账单、发票汇总）及填写说明，请按格式填写后上传。
          </p>
          <a
            href="/template.xlsx"
            download="租赁合同数据模板.xlsx"
            className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 transition-colors"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            下载数据填写模板
          </a>
        </section>

        <hr className="border-gray-100" />

        {/* Step 2: Upload file */}
        <section>
          <h2 className="text-base font-semibold text-gray-700 mb-2">
            <span className="inline-block bg-blue-100 text-blue-700 rounded-full w-6 h-6 text-xs font-bold text-center leading-6 mr-2">2</span>
            上传填好的数据文件
          </h2>
          <div
            className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors ${
              isDragging
                ? "border-blue-400 bg-blue-50"
                : file
                ? "border-green-400 bg-green-50"
                : "border-gray-300 hover:border-blue-300 hover:bg-gray-50"
            }`}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleFileSelect(f);
              }}
            />
            {file ? (
              <div>
                <svg className="w-10 h-10 mx-auto mb-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                <p className="text-sm font-medium text-gray-700">{file.name}</p>
                <p className="text-xs text-gray-400 mt-1">
                  {(file.size / 1024).toFixed(1)} KB · 点击更换文件
                </p>
              </div>
            ) : (
              <div>
                <svg className="w-10 h-10 mx-auto mb-2 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
                <p className="text-sm text-gray-600">拖放 .xlsx 文件到此处，或<span className="text-blue-600 font-medium">点击选择</span></p>
                <p className="text-xs text-gray-400 mt-1">仅支持 .xlsx 格式，最大 10MB</p>
              </div>
            )}
          </div>
        </section>

        <hr className="border-gray-100" />

        {/* Step 3: Date range */}
        <section>
          <h2 className="text-base font-semibold text-gray-700 mb-3">
            <span className="inline-block bg-blue-100 text-blue-700 rounded-full w-6 h-6 text-xs font-bold text-center leading-6 mr-2">3</span>
            选择计算区间
          </h2>
          <div className="flex flex-wrap gap-4 items-center">
            <div>
              <label className="block text-xs text-gray-500 mb-1">起始月</label>
              <input
                type="month"
                value={startMonth}
                onChange={(e) => setStartMonth(e.target.value)}
                className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"
              />
            </div>
            <div className="text-gray-400 mt-4">→</div>
            <div>
              <label className="block text-xs text-gray-500 mb-1">结束月（含）</label>
              <input
                type="month"
                value={endMonth}
                onChange={(e) => setEndMonth(e.target.value)}
                className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"
              />
            </div>
            <p className="text-xs text-gray-400 mt-4 self-end mb-2">
              计算结果包含两端月份
            </p>
          </div>
        </section>

        <hr className="border-gray-100" />

        {/* Submit */}
        <section>
          <button
            onClick={handleSubmit}
            disabled={status === "uploading"}
            className="w-full py-3 bg-blue-600 text-white font-semibold rounded-xl hover:bg-blue-700 disabled:opacity-60 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2"
          >
            {status === "uploading" ? (
              <>
                <svg className="w-5 h-5 animate-spin" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                </svg>
                计算中，请稍候…
              </>
            ) : (
              "开始计算"
            )}
          </button>
        </section>

        {/* Error */}
        {status === "error" && (
          <div className="bg-red-50 border border-red-200 rounded-xl p-4 text-sm text-red-700">
            <strong>错误：</strong>{errorMsg}
          </div>
        )}
      </div>

      {/* Results */}
      {status === "done" && result && (
        <div className="mt-8 space-y-6">
          {/* Summary cards */}
          <div className="bg-white rounded-2xl shadow-sm border border-gray-200 p-6">
            <h2 className="text-lg font-bold text-gray-800 mb-4">计算结果汇总</h2>
            <div className="grid grid-cols-3 gap-4 mb-6">
              <div className="bg-blue-50 rounded-xl p-4 text-center">
                <p className="text-xs text-blue-500 mb-1">合同数量</p>
                <p className="text-2xl font-bold text-blue-700">{result.contract_count}</p>
              </div>
              <div className="bg-green-50 rounded-xl p-4 text-center">
                <p className="text-xs text-green-500 mb-1">应收总额</p>
                <p className="text-xl font-bold text-green-700">{formatMoney(result.total_receivable)}</p>
              </div>
              <div className="bg-purple-50 rounded-xl p-4 text-center">
                <p className="text-xs text-purple-500 mb-1">收入总额</p>
                <p className="text-xl font-bold text-purple-700">{formatMoney(result.total_income)}</p>
              </div>
            </div>

            {/* Download buttons */}
            <div className="flex flex-wrap gap-3">
              <button
                onClick={() => downloadBase64(result.files.lease, "lease.xlsx")}
                className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 rounded-lg text-sm hover:bg-gray-50 transition-colors"
              >
                <svg className="w-4 h-4 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                下载 lease.xlsx（合同汇总）
              </button>
              <button
                onClick={() => downloadBase64(result.files.single, "single.xlsx")}
                className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 rounded-lg text-sm hover:bg-gray-50 transition-colors"
              >
                <svg className="w-4 h-4 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                下载 single.xlsx（月度应收明细）
              </button>
              <button
                onClick={() => downloadBase64(result.files.income, "income.xlsx")}
                className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 rounded-lg text-sm hover:bg-gray-50 transition-colors"
              >
                <svg className="w-4 h-4 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                下载 income.xlsx（月度收入明细）
              </button>
            </div>
          </div>

          {/* Detail table */}
          {result.summary && result.summary.length > 0 && (
            <div className="bg-white rounded-2xl shadow-sm border border-gray-200 overflow-hidden">
              <div className="px-6 py-4 border-b border-gray-100">
                <h3 className="font-semibold text-gray-700">逐合同明细</h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="text-left px-4 py-3 text-gray-500 font-medium">客户名称</th>
                      <th className="text-left px-4 py-3 text-gray-500 font-medium">商户编号</th>
                      <th className="text-right px-4 py-3 text-gray-500 font-medium">应收总额</th>
                      <th className="text-right px-4 py-3 text-gray-500 font-medium">收入总额</th>
                      <th className="text-right px-4 py-3 text-gray-500 font-medium">银行匹配</th>
                      <th className="text-right px-4 py-3 text-gray-500 font-medium">发票匹配</th>
                      <th className="text-left px-4 py-3 text-gray-500 font-medium">备注</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {result.summary.map((row, i) => (
                      <tr key={i} className="hover:bg-gray-50 transition-colors">
                        <td className="px-4 py-3 text-gray-800">{row.customer}</td>
                        <td className="px-4 py-3 text-gray-500">{row.merchant_id}</td>
                        <td className="px-4 py-3 text-right font-medium text-gray-800">{formatMoney(row.receivable)}</td>
                        <td className="px-4 py-3 text-right text-gray-600">{formatMoney(row.income)}</td>
                        <td className="px-4 py-3 text-right text-gray-600">{formatMoney(row.bank_matched)}</td>
                        <td className="px-4 py-3 text-right text-gray-600">{formatMoney(row.invoice_matched)}</td>
                        <td className="px-4 py-3 text-xs text-orange-600">{row.notes || ""}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      <p className="text-center text-xs text-gray-400 mt-10">
        数据仅在浏览器和服务器之间传输，计算完成后服务器端自动清除临时文件
      </p>
    </main>
  );
}
