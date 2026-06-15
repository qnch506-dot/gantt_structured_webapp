"use client";

import { useMemo, useState } from "react";
import { AlertCircle, CheckCircle2, Download, FileSpreadsheet, UploadCloud } from "lucide-react";

type PreviewTask = {
  rowNo: number;
  group: string;
  taskName: string;
  workDays: number | null;
  startDate: string | null;
  endDate: string | null;
  status: "정상" | "자동계산" | "오류" | string;
  message: string;
  warnings: string[];
};

type PreviewResponse = {
  ok: boolean;
  fileName: string;
  count: number;
  tasks: PreviewTask[];
  summary: {
    normal: number;
    auto: number;
    error: number;
  };
};

type PreviewMode = "daily" | "weekly" | "monthly";

type Period = {
  key: string;
  label: string;
  topLabel: string;
  start: Date;
  end: Date;
  widthClass: string;
};

const API_BASE = "https://ganttstructuredwebapp-production.up.railway.app";
const PREVIEW_ROW_LIMIT = 150;


const GROUP_COLORS = [
  "#dc2626",
  "#2563eb",
  "#16a34a",
  "#9333ea",
  "#ea580c",
  "#0891b2",
  "#be123c",
  "#4f46e5",
  "#65a30d",
  "#c026d3"
];

function getGroupColor(groupName: string): string {
  const text = groupName || "기본";
  const sum = Array.from(text).reduce((acc, char) => acc + char.charCodeAt(0), 0);
  return GROUP_COLORS[sum % GROUP_COLORS.length];
}


export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [projectStart, setProjectStart] = useState("2026-01-01");
  const [mode, setMode] = useState("smart");
  const [previewMode, setPreviewMode] = useState<PreviewMode>("daily");
  const [preview, setPreview] = useState<PreviewResponse | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  const canGenerate = useMemo(() => {
    return Boolean(file && preview && preview.summary.error === 0);
  }, [file, preview]);

  function onFileChange(selected: File | null) {
    setFile(selected);
    setPreview(null);
    setMessage("");
  }

  async function requestPreview() {
    if (!file) {
      setMessage("먼저 Excel 파일을 선택하세요.");
      return;
    }

    setLoading(true);
    setMessage("");

    const formData = new FormData();
    formData.append("file", file);
    formData.append("project_start", projectStart);
    formData.append("auto_schedule_mode", mode);

    try {
      const res = await fetch(`${API_BASE}/api/preview`, {
        method: "POST",
        body: formData
      });

      const data = await res.json();

      if (!res.ok || !data.ok) {
        setMessage(data.message ?? "미리보기 생성에 실패했습니다.");
        return;
      }

      setPreview(data);
    } catch (error) {
      setMessage("백엔드 서버에 연결할 수 없습니다. FastAPI가 켜져 있는지 확인하세요.");
    } finally {
      setLoading(false);
    }
  }

  async function generateExcel() {
    if (!file) return;

    setLoading(true);
    setMessage("");

    const formData = new FormData();
    formData.append("file", file);
    formData.append("project_start", projectStart);
    formData.append("auto_schedule_mode", mode);

    try {
      const res = await fetch(`${API_BASE}/api/generate`, {
        method: "POST",
        body: formData
      });

      if (!res.ok) {
        const errorData = await res.json().catch(() => null);
        setMessage(errorData?.message ?? "간트차트 생성에 실패했습니다.");
        return;
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "gantt_result.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      setMessage("백엔드 서버에 연결할 수 없습니다.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="min-h-screen bg-[#F8F8FB] px-5 py-8">
      <div className="mx-auto max-w-[1800px]">
        <section className="mb-8 grid gap-8 xl:grid-cols-[32%_68%]">
          <div className="rounded-[28px] border border-violet-100 bg-white p-7 shadow-sm">
            <p className="mb-2 text-sm font-semibold text-violet-600">Excel Gantt Generator</p>
            <h1 className="text-3xl font-bold tracking-tight text-gray-950">
              백데이터 검증 후 간트차트 생성
            </h1>
            <p className="mt-3 text-sm leading-6 text-gray-600">
              시작일/종료일이 있으면 그대로 반영하고, 없는 작업은 작업일을 기준으로 자동 배치합니다.
              이름/공종은 같은 값끼리 세로 병합하고, 작업명은 공종 내부 순번으로 정리합니다.
            </p>

            <label
              className="mt-7 flex min-h-72 cursor-pointer flex-col items-center justify-center rounded-[24px] border-2 border-dashed border-violet-300 bg-[#FAF8FF] p-8 text-center transition hover:border-violet-500 hover:bg-violet-50"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const dropped = e.dataTransfer.files?.[0];
                if (dropped) onFileChange(dropped);
              }}
            >
              <UploadCloud className="mb-4 h-14 w-14 text-violet-500" />
              <input
                type="file"
                accept=".xlsx,.xlsm"
                className="hidden"
                onChange={(e) => onFileChange(e.target.files?.[0] ?? null)}
              />
              <span className="text-lg font-semibold text-gray-900">
                엑셀 파일을 드래그하세요
              </span>
              <span className="mt-2 text-sm text-gray-500">
                권장 컬럼: 이름 / 작업명 / 작업일 / 시작일 / 종료일
              </span>
              {file && (
                <span className="mt-5 inline-flex max-w-full items-center gap-2 rounded-full bg-white px-4 py-2 text-sm font-medium text-gray-700 shadow-sm">
                  <FileSpreadsheet className="h-4 w-4 shrink-0" />
                  <span className="truncate">{file.name}</span>
                </span>
              )}
            </label>

            <div className="mt-6 rounded-[24px] border border-gray-200 bg-white p-5">
              <h2 className="mb-4 text-lg font-bold">생성 옵션</h2>

              <label className="mb-4 block">
                <span className="mb-2 block text-sm font-medium text-gray-700">공사 시작일</span>
                <input
                  type="date"
                  value={projectStart}
                  onChange={(e) => setProjectStart(e.target.value)}
                  className="w-full rounded-xl border border-gray-300 px-4 py-3 outline-none focus:border-violet-500"
                />
              </label>

              <label className="mb-5 block">
                <span className="mb-2 block text-sm font-medium text-gray-700">자동 배치 방식</span>
                <select
                  value={mode}
                  onChange={(e) => setMode(e.target.value)}
                  className="w-full rounded-xl border border-gray-300 px-4 py-3 outline-none focus:border-violet-500"
                >
                  <option value="smart">Smart: 공종별 일부 겹침</option>
                  <option value="sequential">Sequential: 완전 순차</option>
                </select>
              </label>

              <div className="flex flex-col gap-3">
                <button
                  onClick={requestPreview}
                  disabled={!file || loading}
                  className="rounded-xl bg-gray-900 px-5 py-3 font-semibold text-white disabled:cursor-not-allowed disabled:opacity-50"
                >
                  {loading ? "처리 중..." : "1단계: 백데이터 검증"}
                </button>

                <button
                  onClick={generateExcel}
                  disabled={!canGenerate || loading}
                  className="inline-flex items-center justify-center gap-2 rounded-xl bg-violet-600 px-5 py-3 font-semibold text-white disabled:cursor-not-allowed disabled:opacity-50"
                >
                  <Download className="h-4 w-4" />
                  2단계: 간트차트 다운로드
                </button>
              </div>

              {message && (
                <div className="mt-4 rounded-xl bg-red-50 p-4 text-sm text-red-700">
                  {message}
                </div>
              )}
            </div>
          </div>

          <div className="rounded-[28px] border border-violet-100 bg-white p-5 shadow-sm">
            <GanttPreview
              preview={preview}
              previewMode={previewMode}
              setPreviewMode={setPreviewMode}
            />
          </div>
        </section>

        {preview && (
          <section className="rounded-[28px] bg-white p-6 shadow-sm">
            <div className="mb-5 flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
              <div>
                <h2 className="text-xl font-bold">백데이터 검증 결과</h2>
                <p className="mt-1 text-sm text-gray-500">
                  총 {preview.count}개 작업 인식
                </p>
              </div>

              <div className="grid grid-cols-3 gap-3 text-center">
                <SummaryCard label="정상" value={preview.summary.normal} />
                <SummaryCard label="자동계산" value={preview.summary.auto} />
                <SummaryCard label="오류" value={preview.summary.error} danger={preview.summary.error > 0} />
              </div>
            </div>

            <div className="overflow-hidden rounded-2xl border border-gray-200">
              <div className="max-h-[560px] overflow-auto">
                <table className="w-full min-w-[960px] border-collapse text-sm">
                  <thead className="sticky top-0 bg-gray-900 text-white">
                    <tr>
                      <Th>상태</Th>
                      <Th>이름</Th>
                      <Th>작업명</Th>
                      <Th>작업일</Th>
                      <Th>시작일</Th>
                      <Th>종료일</Th>
                      <Th>메시지</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.tasks.map((task, index) => (
                      <tr key={`${task.rowNo}-${index}`} className="border-t border-gray-100">
                        <td className="px-4 py-3">
                          <StatusBadge status={task.status} />
                        </td>
                        <td className="px-4 py-3 text-gray-700">{task.group}</td>
                        <td className="px-4 py-3 font-medium text-gray-900">{task.taskName}</td>
                        <td className="px-4 py-3 text-center">{task.workDays ?? "-"}</td>
                        <td className="px-4 py-3 text-center">{task.startDate ?? "-"}</td>
                        <td className="px-4 py-3 text-center">{task.endDate ?? "-"}</td>
                        <td className="px-4 py-3 text-gray-500">
                          {task.message || task.warnings?.join(" / ") || "-"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {preview.summary.error > 0 && (
              <div className="mt-4 flex gap-2 rounded-xl bg-amber-50 p-4 text-sm text-amber-800">
                <AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />
                오류가 있는 작업은 간트차트 생성에서 제외되지 않고, 생성 자체가 중단됩니다.
                작업일 또는 시작일/종료일을 백데이터에 보완하세요.
              </div>
            )}
          </section>
        )}
      </div>
    </main>
  );
}


function shouldShowGroupCell(tasks: PreviewTask[], index: number): boolean {
  if (index === 0) return true;
  return tasks[index].group !== tasks[index - 1].group;
}

function getGroupRowSpan(tasks: PreviewTask[], index: number): number {
  const group = tasks[index].group;
  let span = 1;
  for (let i = index + 1; i < tasks.length; i += 1) {
    if (tasks[i].group !== group) break;
    span += 1;
  }
  return span;
}

function getTaskNumberInGroup(tasks: PreviewTask[], index: number): number {
  const group = tasks[index].group;
  let count = 0;
  for (let i = 0; i <= index; i += 1) {
    if (tasks[i].group === group) count += 1;
  }
  return count;
}

function formatPreviewTaskName(tasks: PreviewTask[], index: number): string {
  const taskName = tasks[index].taskName ?? "";
  if (/^\d+\s*[\.\)]\s*/.test(taskName.trim())) {
    return taskName;
  }
  return `${getTaskNumberInGroup(tasks, index)}. ${taskName}`;
}


function GanttPreview({
  preview,
  previewMode,
  setPreviewMode
}: {
  preview: PreviewResponse | null;
  previewMode: PreviewMode;
  setPreviewMode: (mode: PreviewMode) => void;
}) {
  const validTasks = useMemo(() => {
    return preview?.tasks.filter((task) => task.startDate && task.endDate && task.status !== "오류") ?? [];
  }, [preview]);

  const displayTasks = useMemo(() => validTasks.slice(0, PREVIEW_ROW_LIMIT), [validTasks]);
  const hiddenTaskCount = Math.max(validTasks.length - displayTasks.length, 0);
  const periods = useMemo(() => buildPeriods(validTasks, previewMode), [validTasks, previewMode]);

  if (!preview) {
    return (
      <div className="flex min-h-[760px] flex-col overflow-hidden rounded-[22px] border border-gray-200 bg-white">
        <ExcelWindowHeader title="간트차트 미리보기" />
        <div className="flex flex-1 flex-col items-center justify-center bg-[linear-gradient(#f1f5f9_1px,transparent_1px),linear-gradient(90deg,#f1f5f9_1px,transparent_1px)] bg-[size:28px_28px] p-8 text-center">
          <FileSpreadsheet className="mb-4 h-16 w-16 text-violet-400" />
          <h2 className="text-xl font-bold text-gray-900">Excel 미리보기 대기 중</h2>
          <p className="mt-2 max-w-md text-sm leading-6 text-gray-500">
            파일을 업로드하고 백데이터 검증을 누르면 일별/주간별/월간별 간트차트 미리보기가 표시됩니다.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex min-h-[760px] flex-col overflow-hidden rounded-[22px] border border-gray-200 bg-white">
      <ExcelWindowHeader title="간트차트 미리보기" />

      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-gray-200 bg-white px-4 py-3">
        <div>
          <div className="text-sm font-bold text-gray-900">{preview.fileName}</div>
          <div className="text-xs text-gray-500">
            정상 {preview.summary.normal} · 자동계산 {preview.summary.auto} · 오류 {preview.summary.error}
          </div>
        </div>

        <div className="flex rounded-xl bg-gray-100 p-1">
          <PreviewTab label="일별" active={previewMode === "daily"} onClick={() => setPreviewMode("daily")} />
          <PreviewTab label="주간별" active={previewMode === "weekly"} onClick={() => setPreviewMode("weekly")} />
          <PreviewTab label="월간별" active={previewMode === "monthly"} onClick={() => setPreviewMode("monthly")} />
        </div>
      </div>

      {hiddenTaskCount > 0 && (
        <div className="border-b border-amber-200 bg-amber-50 px-4 py-2 text-xs text-amber-800">
          웹 미리보기는 속도 때문에 상위 {PREVIEW_ROW_LIMIT}개 작업만 표시합니다. 다운로드 Excel에는 전체 {validTasks.length}개 작업이 모두 포함됩니다.
        </div>
      )}

      <div className="flex-1 overflow-auto bg-[#fbfbfd]">
        {validTasks.length === 0 ? (
          <div className="flex h-full items-center justify-center p-10 text-center text-sm text-gray-500">
            미리보기 가능한 정상 작업이 없습니다.
          </div>
        ) : (
          <table className="min-w-full border-collapse text-xs">
            <thead className="sticky top-0 z-20">
              <tr>
                <th className="sticky left-0 z-30 w-32 min-w-32 border border-gray-200 bg-gray-100 px-3 py-2 text-left font-semibold text-gray-700">
                  공종
                </th>
                <th className="sticky left-32 z-30 w-72 min-w-72 border border-gray-200 bg-gray-100 px-3 py-2 text-left font-semibold text-gray-700">
                  작업명
                </th>
                {periods.map((period) => (
                  <th
                    key={`top-${period.key}`}
                    className={`${period.widthClass} border border-gray-200 bg-gray-100 px-1 py-1 text-center font-semibold text-gray-500`}
                  >
                    {period.topLabel}
                  </th>
                ))}
              </tr>
              <tr>
                <th className="sticky left-0 z-30 border border-gray-200 bg-gray-900 px-3 py-2 text-left font-semibold text-white">
                  이름
                </th>
                <th className="sticky left-32 z-30 border border-gray-200 bg-gray-900 px-3 py-2 text-left font-semibold text-white">
                  작업명
                </th>
                {periods.map((period) => (
                  <th
                    key={period.key}
                    className={`${period.widthClass} whitespace-pre-line border border-gray-200 bg-gray-900 px-1 py-2 text-center font-semibold text-white`}
                  >
                    {period.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {displayTasks.map((task, rowIndex) => {
                const taskStart = parseDate(task.startDate);
                const taskEnd = parseDate(task.endDate);
                const groupColor = getGroupColor(task.group);

                return (
                  <tr key={`${task.rowNo}-${rowIndex}`} className="h-9">
                    {shouldShowGroupCell(displayTasks, rowIndex) && (
                      <td
                        rowSpan={getGroupRowSpan(displayTasks, rowIndex)}
                        className="sticky left-0 z-10 border border-gray-300 bg-gray-50 px-3 py-2 text-center align-middle text-[12px] font-bold text-gray-800"
                      >
                        {task.group}
                      </td>
                    )}
                    <td className="sticky left-32 z-10 border border-gray-200 bg-white px-3 py-2 text-gray-900">
                      <span className="line-clamp-1">{formatPreviewTaskName(displayTasks, rowIndex)}</span>
                    </td>
                    {periods.map((period) => {
                      const isActive = taskStart && taskEnd && overlaps(taskStart, taskEnd, period.start, period.end);
                      const isMilestone = task.workDays === 0 && isActive;
                      const startsHere = taskStart && contains(period.start, period.end, taskStart);
                      const endsHere = taskEnd && contains(period.start, period.end, taskEnd);

                      return (
                        <td
                          key={`${task.rowNo}-${period.key}`}
                          className={`${period.widthClass} border border-gray-200 p-0`}
                        >
                          <div className="relative flex h-8 items-center justify-center bg-white">
                            {isMilestone ? (
                              <span className="relative z-10 text-[14px] font-black" style={{ color: groupColor }}>◆</span>
                            ) : isActive ? (
                              <>
                                <div
                                  className="absolute left-[-2px] right-[-2px] top-1/2 h-[3px] -translate-y-1/2 rounded-full"
                                  style={{ backgroundColor: groupColor }}
                                />
                                {startsHere && (
                                  <span
                                    className="absolute left-[-4px] top-1/2 z-10 -translate-y-1/2 text-[14px] font-black leading-none"
                                    style={{ color: groupColor }}
                                  >
                                    ◀
                                  </span>
                                )}
                                {endsHere && (
                                  <span
                                    className="absolute right-[-4px] top-1/2 z-10 -translate-y-1/2 text-[14px] font-black leading-none"
                                    style={{ color: groupColor }}
                                  >
                                    ▶
                                  </span>
                                )}
                              </>
                            ) : (
                              ""
                            )}
                          </div>
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      <div className="flex items-center gap-2 border-t border-gray-200 bg-white px-4 py-3 text-xs text-gray-500">
        <span className="rounded-md bg-red-50 px-2 py-1 font-semibold text-red-700">화살표</span>
        <span>이름/공종 병합 · 같은 이름/공종 = 같은 색 · 굵은 연결선 · ◀ 시작 · ▶ 종료 · ◆ 마일스톤</span>
        <span className="mx-1 text-gray-300">/</span>
        <span className="rounded-md bg-violet-100 px-2 py-1 font-semibold text-violet-700">일별</span>
        <span>하루 = 한 칸</span>
        <span className="mx-1 text-gray-300">/</span>
        <span className="rounded-md bg-violet-100 px-2 py-1 font-semibold text-violet-700">주간별</span>
        <span>1주 = 한 칸</span>
        <span className="mx-1 text-gray-300">/</span>
        <span className="rounded-md bg-violet-100 px-2 py-1 font-semibold text-violet-700">월간별</span>
        <span>1개월 = 한 칸</span>
      </div>
    </div>
  );
}

function ExcelWindowHeader({ title }: { title: string }) {
  return (
    <div className="border-b border-gray-200 bg-gradient-to-b from-gray-50 to-white">
      <div className="flex items-center gap-2 px-4 py-3">
        <span className="h-3 w-3 rounded-full bg-red-400" />
        <span className="h-3 w-3 rounded-full bg-yellow-400" />
        <span className="h-3 w-3 rounded-full bg-green-400" />
        <span className="ml-3 text-sm font-semibold text-gray-700">{title}</span>
      </div>
      <div className="flex gap-1 border-t border-gray-100 px-4 py-2 text-xs text-gray-500">
        <span className="rounded-md bg-gray-100 px-2 py-1">파일</span>
        <span className="rounded-md bg-gray-100 px-2 py-1">홈</span>
        <span className="rounded-md bg-gray-100 px-2 py-1">삽입</span>
        <span className="rounded-md bg-gray-100 px-2 py-1">보기</span>
      </div>
    </div>
  );
}

function PreviewTab({
  label,
  active,
  onClick
}: {
  label: string;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`rounded-lg px-3 py-2 text-sm font-semibold transition ${
        active ? "bg-white text-violet-700 shadow-sm" : "text-gray-500 hover:text-gray-900"
      }`}
    >
      {label}
    </button>
  );
}

function buildPeriods(tasks: PreviewTask[], mode: PreviewMode): Period[] {
  const dates = tasks
    .flatMap((task) => [task.startDate, task.endDate])
    .filter(Boolean)
    .map((value) => parseDate(value))
    .filter((value): value is Date => Boolean(value));

  if (dates.length === 0) return [];

  const minDate = new Date(Math.min(...dates.map((d) => d.getTime())));
  const maxDate = new Date(Math.max(...dates.map((d) => d.getTime())));

  if (mode === "daily") {
    const periods: Period[] = [];
    let cur = startOfDay(minDate);
    while (cur <= maxDate) {
      periods.push({
        key: isoDate(cur),
        topLabel: cur.getDate() === 1 || periods.length === 0 ? `${cur.getFullYear()}.${String(cur.getMonth() + 1).padStart(2, "0")}` : "",
        label: String(cur.getDate()),
        start: new Date(cur),
        end: new Date(cur),
        widthClass: "min-w-8 w-8"
      });
      cur = addDays(cur, 1);
    }
    return periods;
  }

  if (mode === "weekly") {
    const periods: Period[] = [];
    let cur = startOfWeek(minDate);
    let index = 1;
    while (cur <= maxDate) {
      const end = addDays(cur, 6);
      const visibleStart = cur < minDate ? minDate : cur;
      const visibleEnd = end > maxDate ? maxDate : end;
      periods.push({
        key: `week-${isoDate(cur)}`,
        topLabel: `${visibleStart.getFullYear()}.${String(visibleStart.getMonth() + 1).padStart(2, "0")}`,
        label: `${index}주\n${formatMonthDay(visibleStart)}~${formatMonthDay(visibleEnd)}`,
        start: new Date(cur),
        end,
        widthClass: "min-w-20 w-20"
      });
      cur = addDays(cur, 7);
      index += 1;
    }
    return periods;
  }

  const periods: Period[] = [];
  let cur = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
  while (cur <= maxDate) {
    const end = new Date(cur.getFullYear(), cur.getMonth() + 1, 0);
    periods.push({
      key: `month-${cur.getFullYear()}-${cur.getMonth()}`,
      topLabel: String(cur.getFullYear()),
      label: `${cur.getMonth() + 1}월`,
      start: new Date(cur),
      end,
      widthClass: "min-w-24 w-24"
    });
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  return periods;
}

function parseDate(value: string | null): Date | null {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function startOfWeek(date: Date): Date {
  const result = startOfDay(date);
  const day = result.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  result.setDate(result.getDate() + mondayOffset);
  return result;
}

function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function isoDate(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatMonthDay(date: Date): string {
  return `${String(date.getMonth() + 1).padStart(2, "0")}/${String(date.getDate()).padStart(2, "0")}`;
}

function overlaps(start: Date, end: Date, periodStart: Date, periodEnd: Date): boolean {
  return start <= periodEnd && end >= periodStart;
}

function contains(periodStart: Date, periodEnd: Date, target: Date): boolean {
  return periodStart <= target && target <= periodEnd;
}

function SummaryCard({
  label,
  value,
  danger
}: {
  label: string;
  value: number;
  danger?: boolean;
}) {
  return (
    <div className={`rounded-2xl px-5 py-3 ${danger ? "bg-red-50 text-red-700" : "bg-gray-50 text-gray-800"}`}>
      <div className="text-xs font-medium">{label}</div>
      <div className="text-xl font-bold">{value}</div>
    </div>
  );
}

function Th({ children }: { children: React.ReactNode }) {
  return <th className="px-4 py-3 text-left text-xs font-semibold">{children}</th>;
}

function StatusBadge({ status }: { status: string }) {
  if (status === "정상") {
    return (
      <span className="inline-flex items-center gap-1 rounded-full bg-green-50 px-3 py-1 text-xs font-semibold text-green-700">
        <CheckCircle2 className="h-3 w-3" />
        정상
      </span>
    );
  }

  if (status === "자동계산") {
    return (
      <span className="inline-flex items-center gap-1 rounded-full bg-blue-50 px-3 py-1 text-xs font-semibold text-blue-700">
        자동계산
      </span>
    );
  }

  return (
    <span className="inline-flex items-center gap-1 rounded-full bg-red-50 px-3 py-1 text-xs font-semibold text-red-700">
      <AlertCircle className="h-3 w-3" />
      오류
    </span>
  );
}
