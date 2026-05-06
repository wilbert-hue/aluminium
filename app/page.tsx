import fs from 'fs/promises'
import path from 'path'
import Image from 'next/image'
import { ChevronDown } from 'lucide-react'
import {
  getCmiExcelFilename,
  parseCmiWorkbookFromBuffer,
  type CmiHeaderCell,
  type CmiSheetModel,
} from '@/lib/cmi-excel'

function dashboardSubtitleFromSheets(sheets: CmiSheetModel[]): string {
  const title = sheets[0]?.banner.title ?? ''
  if (/dead\s*burned/i.test(title)) return 'Dead burned Magnesia Buyers'
  if (/fused magnesia/i.test(title)) return 'Fused Magnesia Buyers'
  return 'Fused Magnesia Buyers'
}

function headerCellClass(cell: CmiHeaderCell): string {
  const base =
    'border border-black px-2 py-2 text-center align-middle text-gray-900 leading-snug'
  if (cell.variant === 'sno') return `${base} bg-[#f9e79f] font-semibold`
  if (cell.variant === 'leaf')
    return `${base} bg-[#e8f5e9] text-xs font-semibold`
  return `${base} bg-[#e8f5e9] text-xs font-semibold`
}

function CmiPropositionBlock({ sheet }: { sheet: CmiSheetModel }) {
  const hasTable =
    sheet.headerRows.length > 0 && sheet.columnCount > 0

  return (
    <details
      open
      className="group rounded-lg border border-gray-200 bg-white shadow-sm overflow-hidden"
    >
      <summary className="flex cursor-pointer list-none items-center justify-between gap-3 px-4 py-3 bg-white border-b border-gray-200 hover:bg-gray-50 [&::-webkit-details-marker]:hidden">
        <span className="text-base font-semibold text-gray-900">
          {sheet.displayTitle}
        </span>
        <ChevronDown className="h-5 w-5 shrink-0 text-gray-600 transition-transform group-open:rotate-180" />
      </summary>

      <div className="p-4 bg-gray-100">
        {!hasTable ? (
          <p className="text-sm text-gray-600">No table structure in this sheet.</p>
        ) : (
          <div className="overflow-x-auto rounded-md border border-gray-300 bg-white">
            {/*
              Flex column with w-max + min-w-full: width follows the wide table (p3)
              so the blue banner stretches to the same width when horizontal scroll is needed.
            */}
            <div className="flex w-max min-w-full flex-col">
              <div className="shrink-0 bg-[#2c3e50] px-4 py-3 text-center text-white">
                <div className="text-sm font-semibold leading-tight">
                  {sheet.banner.title}
                </div>
                {sheet.banner.subtitle ? (
                  <div className="mt-1 text-xs leading-snug text-white/90">
                    {sheet.banner.subtitle}
                  </div>
                ) : null}
              </div>

            <table className="min-w-max border-collapse border border-black text-sm text-gray-900">
              <thead>
                {sheet.headerRows.map((row, ri) => (
                  <tr key={ri}>
                    {row.cells.map((cell, ci) => (
                      <th
                        key={`${ri}-${ci}`}
                        scope="col"
                        rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                        colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                        className={headerCellClass(cell)}
                      >
                        {cell.text || '\u00a0'}
                      </th>
                    ))}
                  </tr>
                ))}
              </thead>
              <tbody>
                {sheet.bodyRows.map((row, ri) => (
                  <tr key={ri}>
                    {Array.from({ length: sheet.columnCount }, (_, ci) => (
                      <td
                        key={ci}
                        className="border border-black px-2 py-1.5 whitespace-nowrap bg-white"
                      >
                        {row[ci] === '' || row[ci] == null ? (
                          <span className="text-gray-500">—</span>
                        ) : (
                          String(row[ci])
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
      </div>
    </details>
  )
}

export default async function DashboardPage() {
  const filePath = path.join(process.cwd(), getCmiExcelFilename())
  let sheets: CmiSheetModel[] = []
  let loadError: string | null = null

  try {
    const buf = await fs.readFile(filePath)
    sheets = parseCmiWorkbookFromBuffer(buf)
  } catch (e) {
    loadError =
      e instanceof Error ? e.message : 'Could not read the CMI Excel file.'
  }

  const dashboardSubtitle =
    sheets.length > 0
      ? dashboardSubtitleFromSheets(sheets)
      : 'Fused Magnesia Buyers'

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Top bar — logo left, titles centered (Coherent-style) */}
      <header className="bg-white border-b border-gray-200">
        <div className="container mx-auto flex max-w-[1800px] items-center gap-4 px-4 py-5">
          <div className="flex w-[clamp(140px,26vw,200px)] shrink-0 justify-start">
            <Image
              src="/logo.png"
              alt="Coherent Market Insights"
              width={180}
              height={72}
              className="h-auto w-auto max-w-[180px]"
              priority
            />
          </div>
          <div className="min-w-0 flex-1 px-2 text-center">
            <h1 className="text-2xl font-bold text-gray-900 md:text-3xl">
              Coherent Dashboard
            </h1>
            <p className="mt-1 text-sm text-gray-500 md:text-base">
              {dashboardSubtitle}
            </p>
          </div>
          <div
            className="hidden w-[clamp(140px,26vw,200px)] shrink-0 sm:block"
            aria-hidden
          />
        </div>
      </header>

      <div className="container mx-auto max-w-[1800px] px-4 py-6">
        <h2 className="text-xl font-bold text-gray-900">
          Customer Intelligence Database
        </h2>
        <p className="mt-2 mb-5 max-w-4xl text-sm font-medium uppercase tracking-wide text-amber-900 bg-amber-50 border border-amber-200 rounded-md px-3 py-2.5">
          NOTE: All the data in the dashboard is demo data. No real world data
          is related to this.
        </p>

        {loadError ? (
          <div
            className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800"
            role="alert"
          >
            <p className="font-medium">Unable to load workbook</p>
            <p className="mt-1">{loadError}</p>
            <p className="mt-2 text-red-700">
              Place{' '}
              <code className="rounded bg-red-100 px-1 py-0.5 text-xs">
                {getCmiExcelFilename()}
              </code>{' '}
              in the project root.
            </p>
          </div>
        ) : (
          <div className="grid grid-cols-1 gap-6 lg:grid-cols-12">
            {/* Sidebar — chart view */}
            <aside className="lg:col-span-3">
              <div className="sticky top-6 rounded-lg border border-gray-200 bg-white p-4 shadow-sm">
                <h2 className="text-xs font-bold uppercase tracking-wider text-gray-500 mb-3">
                  Chart view
                </h2>
                <div className="rounded-lg border border-teal-200 border-l-4 border-l-teal-500 bg-cyan-50/90 p-3 shadow-sm">
                  <div className="flex items-start gap-2">
                    <span className="text-lg" aria-hidden>
                      👤
                    </span>
                    <div>
                      <div className="text-sm font-semibold text-teal-900">
                        Customer Intelligence
                      </div>
                      <p className="mt-1 text-xs text-teal-800/90 leading-snug">
                        Customer database with proposition tables sourced from
                        the CMI workbook.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </aside>

            <main className="lg:col-span-9 space-y-6">
              {sheets.length === 0 ? (
                <p className="text-gray-600">No proposition sheets found.</p>
              ) : (
                sheets.map((sheet) => (
                  <CmiPropositionBlock key={sheet.sheetName} sheet={sheet} />
                ))
              )}
            </main>
          </div>
        )}
      </div>
    </div>
  )
}
