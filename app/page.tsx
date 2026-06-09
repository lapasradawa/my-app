'use client'

import { useRef, useState } from 'react'
import { processExcel, type ResultRow, type ProcessResult } from '@/lib/excel-parser'
import { exportToExcel } from '@/lib/excel-exporter'

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [result, setResult] = useState<ProcessResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [filename, setFilename] = useState('')
  const [dragging, setDragging] = useState(false)

  async function handleFile(file: File) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setError('กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น')
      return
    }
    setLoading(true)
    setError(null)
    setResult(null)
    setFilename(file.name)
    try {
      const buffer = await file.arrayBuffer()
      const data = processExcel(buffer)
      setResult(data)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'เกิดข้อผิดพลาด')
    } finally {
      setLoading(false)
    }
  }

  function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) handleFile(file)
  }

  function onDrop(e: React.DragEvent) {
    e.preventDefault()
    setDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) handleFile(file)
  }

  function onDragOver(e: React.DragEvent) {
    e.preventDefault()
    setDragging(true)
  }

  function onDragLeave() {
    setDragging(false)
  }

  function handleExport() {
    if (!result) return
    exportToExcel(result.rows, result.containerNames, filename.replace('.xlsx', '') + '-PO-Matching.xlsx')
  }

  function reset() {
    setResult(null)
    setError(null)
    setFilename('')
    if (inputRef.current) inputRef.current.value = ''
  }

  const fmt = (n: number) => n === 0 ? '-' : n.toLocaleString()

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-screen-xl mx-auto">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">CI / Packing List — PO Matching</h1>
          <p className="text-sm text-gray-500 mt-1">อัปโหลดไฟล์ Excel เพื่อ match สินค้าในแต่ละตู้กับ PO</p>
        </div>

        {/* Upload area */}
        {!result && !loading && (
          <div
            className={`border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-colors ${
              dragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 bg-white hover:border-blue-400'
            }`}
            onClick={() => inputRef.current?.click()}
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
          >
            <div className="text-4xl mb-3">📂</div>
            <p className="text-gray-700 font-medium">ลากไฟล์มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์</p>
            <p className="text-gray-400 text-sm mt-1">รองรับไฟล์ .xlsx และ .xls</p>
            <p className="text-gray-400 text-xs mt-2">ไฟล์ต้องมีชีท CI และชีทตู้ (เช่น UETU2248127)</p>
            <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={onFileChange} />
          </div>
        )}

        {loading && (
          <div className="text-center py-12 text-gray-500">กำลังประมวลผล...</div>
        )}

        {error && (
          <div className="mt-4 bg-red-50 border border-red-200 text-red-700 rounded-lg px-4 py-3 text-sm">
            {error}
            <button onClick={reset} className="ml-3 underline">ลองใหม่</button>
          </div>
        )}

        {result && (
          <>
            <div className="flex items-center justify-between mb-4">
              <div>
                <span className="text-sm text-gray-500">ไฟล์: </span>
                <span className="text-sm font-medium text-gray-700">{filename}</span>
                <span className="ml-3 text-sm text-gray-400">
                  {result.rows.length} รายการ | {result.containerNames.length} ตู้
                </span>
              </div>
              <div className="flex gap-2">
                <button
                  onClick={handleExport}
                  className="px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors"
                >
                  Export Excel
                </button>
                <button
                  onClick={reset}
                  className="px-4 py-2 bg-gray-200 text-gray-700 text-sm rounded-lg hover:bg-gray-300 transition-colors"
                >
                  อัปโหลดใหม่
                </button>
              </div>
            </div>

            <div className="overflow-auto rounded-xl border border-gray-200 bg-white shadow-sm">
              <table className="text-sm w-full border-collapse">
                <thead>
                  <tr className="bg-gray-100 text-gray-700">
                    <th className="px-3 py-2 text-center border-b border-gray-200 whitespace-nowrap">N</th>
                    <th className="px-3 py-2 text-left border-b border-gray-200 whitespace-nowrap min-w-[180px]">Code</th>
                    <th className="px-3 py-2 text-left border-b border-gray-200 min-w-[300px]">Description</th>
                    <th className="px-3 py-2 text-left border-b border-gray-200 whitespace-nowrap min-w-[200px]">PO</th>
                    <th className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">QTY</th>
                    {result.containerNames.map(name => (
                      <th key={name} className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">{name}</th>
                    ))}
                    <th className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">LEFT</th>
                  </tr>
                </thead>
                <tbody>
                  {result.rows.map((row: ResultRow) => (
                    <tr key={row.no} className="hover:bg-gray-50 border-b border-gray-100 last:border-0">
                      <td className="px-3 py-2 text-center text-gray-500">{row.no}</td>
                      <td className="px-3 py-2 font-mono text-xs">{row.code}</td>
                      <td className="px-3 py-2 text-gray-700">{row.description}</td>
                      <td className="px-3 py-2 text-gray-600 whitespace-nowrap">{row.po}</td>
                      <td className="px-3 py-2 text-right font-medium">{row.qty.toLocaleString()}</td>
                      {result.containerNames.map(name => (
                        <td key={name} className="px-3 py-2 text-right text-gray-500">
                          {fmt(row.containers[name])}
                        </td>
                      ))}
                      <td className={`px-3 py-2 text-right font-medium ${row.left > 0 ? 'text-orange-600' : 'text-gray-400'}`}>
                        {fmt(row.left)}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-gray-50 font-semibold text-gray-700 border-t-2 border-gray-200">
                    <td colSpan={4} className="px-3 py-2 text-right">รวม</td>
                    <td className="px-3 py-2 text-right">
                      {result.rows.reduce((s, r) => s + r.qty, 0).toLocaleString()}
                    </td>
                    {result.containerNames.map(name => (
                      <td key={name} className="px-3 py-2 text-right">
                        {result.rows.reduce((s, r) => s + (r.containers[name] || 0), 0).toLocaleString()}
                      </td>
                    ))}
                    <td className="px-3 py-2 text-right">
                      {result.rows.reduce((s, r) => s + r.left, 0).toLocaleString()}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  )
}
