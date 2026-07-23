'use client'

import Link from 'next/link'
import LockButton from '@/components/LockButton'
import { useState } from 'react'

const sections = [
  {
    href: '/',
    label: 'PO Matching',
    color: 'bg-blue-50 border-blue-200',
    iconBg: 'bg-blue-600',
    icon: '📥',
    title: 'PO Matching — หน้าจัดการ Invoice',
    subtitle: 'หน้าหลักสำหรับบันทึก Invoice และ Match PO กับ Invoice สินค้าขาเข้า',
    steps: [
      { label: 'เตรียมไฟล์ Invoice', desc: 'ไฟล์ Excel ที่จะอัปโหลดต้องมี Sheet ชื่อ "CI" และ "PL" เท่านั้นจึงจะอัปโหลดได้ — หากระบบแจ้งว่าไม่พบ CI ให้ตรวจสอบว่าชื่อ Sheet เป็นตัว "ไอ (I)" หรือตัว "แอล (l)" เพราะหน้าตาคล้ายกันมาก' },
      { label: 'อัปโหลดไฟล์ Invoice', desc: 'กดปุ่มอัปโหลดแล้วเลือกไฟล์ Excel เมื่ออัปโหลดสำเร็จ ระบบจะ Match สินค้าแต่ละตู้กับเลข PO ให้อัตโนมัติ โดยไฟล์จะต้องมีคอลัมน์ที่ระบุเลข PO มาก่อนแล้ว' },
      { label: 'ตรวจสอบและบันทึก', desc: 'ตรวจสอบข้อมูลที่ระบบ Match ให้ว่าถูกต้อง จากนั้นกดบันทึกทุกครั้งเพื่อให้ข้อมูลถูกบันทึกเข้าระบบ — ทุกการบันทึกต้องใส่รหัสผ่านก่อนเสมอ' },
    ],
    tips: [
      'ตรวจชื่อ Sheet ให้ดีก่อนอัปโหลด: "CI" ต้องเป็น C + I (ไอ) ไม่ใช่ C + l (แอล)',
      'ไฟล์ที่อัปโหลดต้องมีคอลัมน์เลข PO ครบก่อน ระบบถึงจะ Match ได้ถูกต้อง',
    ],
  },
  {
    href: '/dashboard',
    label: 'Dashboard',
    color: 'bg-purple-50 border-purple-200',
    iconBg: 'bg-purple-600',
    icon: '📊',
    title: 'Dashboard — ภาพรวมทุก Shipment',
    subtitle: 'ดูสถานะสินค้าทุก Invoice แบบ Real-time และจัดการข้อมูล Shipment ทั้งหมดในหน้าเดียว',
    steps: [
      { label: 'กรอกข้อมูล Shipment', desc: 'หลังบันทึก Invoice เรียบร้อยแล้ว มาที่หน้า Dashboard กรอกชื่อ Supplier, สถานะ Shipment (On board / กำลังเข้าคลัง / เข้าคลังแล้ว), วันที่ถึงท่าเรือ (ETA) และประมาณการเข้าคลัง จากนั้นกดบันทึก' },
      { label: 'กรอก Exchange Rate', desc: 'กดที่ Invoice No. เพื่อเข้าหน้า Detail — ระบบจะแสดงยอดรวม Invoice อัตโนมัติ กดแก้ไขที่กล่อง "ยอดรวม Invoice" เพื่อกรอก Exchange Rate (CNY หรือ USD) แล้วกดบันทึก สามารถกรอกได้หลาย Exchange Rate ถ้าจ่ายหลายงวดด้วยอัตราต่างกัน' },
      { label: 'อัปโหลดเอกสาร', desc: 'อัปโหลดไฟล์ใบขน และเอกสาร B/L ในช่องที่กำหนด' },
      { label: 'บันทึก Cost Saving', desc: 'ที่ช่อง "Cost Saving" กดปุ่ม "เพิ่มข้อมูล" เพื่อกรอก Cost Saving (THB), % Cost Saving และอัปโหลดไฟล์เอกสาร Cost Saving แล้วกดบันทึก' },
      { label: 'บันทึก B/L Date', desc: 'กรอก B/L Date ในกล่องที่กำหนด แล้วกดบันทึก — Due Date (B/L Date + 30 วัน) จะขึ้นให้อัตโนมัติ' },
      { label: 'บันทึกการจ่ายเงิน', desc: 'กดปุ่ม "Submit Payment" ใส่วันที่จ่าย และอัปโหลดไฟล์หลักฐานการจ่ายเงิน (PDF หรือรูปภาพ) แล้วกดบันทึก หลังจากนั้นกลับไปที่กล่อง "ยอดรวม Invoice" กดแก้ไขเพื่อกรอก Exchange Rate จริงที่ใช้จ่าย แล้วกดบันทึก' },
      { label: 'ติดตามสินค้าด้วย Item Code', desc: 'ใส่ Item Code ในช่องค้นหา เพื่อติดตามว่าสินค้านั้นอยู่ใน Invoice ไหน เข้าคลังวันที่เท่าไหร่ และมีจำนวนเท่าใด' },
    ],
    tips: [
      'Dashboard จะแสดง Due Date และสถานะการจ่ายเงินให้อัตโนมัติหลังจากกรอกข้อมูลครบ',
      'ต้องปลดล็อก (ใส่รหัสผ่าน) ก่อนแก้ไขข้อมูลทุกครั้ง',
      'สามารถกรอก Exchange Rate ได้หลายอัตราในกรณีที่จ่ายเงินหลายงวดด้วยอัตราแลกเปลี่ยนต่างกัน',
    ],
  },
  {
    href: '/report',
    label: 'Report',
    color: 'bg-amber-50 border-amber-200',
    iconBg: 'bg-amber-500',
    icon: '📈',
    title: 'Report — สรุปยอด FOB และ Cost Saving รายเดือน',
    subtitle: 'ดูยอดรวม FOB, Cost Saving และวันที่จ่ายเงิน แยกตามเดือน',
    steps: [
      { label: 'เลือกวิธีจัดกลุ่ม', desc: 'กดปุ่ม "จัดกลุ่มตาม" เพื่อเลือกว่าจะแบ่งเดือนตาม: เข้าคลัง (วันที่สินค้าถึงไทย), Due Date (กำหนดชำระ), หรือ Payment (วันที่จ่ายจริง)' },
      { label: 'เลือกเดือนที่ต้องการ', desc: 'กดที่ chip สีเหลืองเพื่อเลือก/ยกเลิกเดือน หรือกด "เลือกทั้งหมด" / "ยกเลิกทั้งหมด"' },
      { label: 'ค้นหา Invoice', desc: 'พิมพ์ Invoice No. ในช่องค้นหาเพื่อกรองเฉพาะ Invoice ที่ต้องการ' },
      { label: 'ดู Exchange Rate Breakdown', desc: 'กดที่ตัวเลข "Actual FOB THB (Finance)" เพื่อดูรายละเอียดอัตราแลกเปลี่ยนและยอดแต่ละงวด' },
      { label: 'Export Excel', desc: 'กดปุ่ม "Export Excel" มุมบนขวา เพื่อดาวน์โหลดรายงานเป็นไฟล์ Excel' },
    ],
    tips: ['Commission Payment สามารถกรอกได้หลังจากปลดล็อก — วันที่ที่กรอกจะบันทึกในระบบทันที', 'ยอดในแถว GRAND TOTAL คำนวณจากเฉพาะเดือนที่เลือก'],
  },
  {
    href: '/compare',
    label: 'Cost Compare',
    color: 'bg-green-50 border-green-200',
    iconBg: 'bg-green-600',
    icon: '🔍',
    title: 'Cost Compare — เปรียบเทียบราคาจากทุก Supplier',
    subtitle: 'อัปโหลด PO ของแต่ละ Supplier แล้วเปรียบเทียบราคา FOB และ DDP ในตารางเดียว',
    steps: [
      { label: 'อัปโหลดไฟล์ PO', desc: 'กรอกชื่อ Project และ Supplier จากนั้นกด "+ อัปโหลดไฟล์ PO (Excel)" เลือกไฟล์ Excel PO ของ Supplier นั้น ระบบจะอ่าน Item Code, Description และราคาให้อัตโนมัติ' },
      { label: 'เลือก Project', desc: 'กดที่ชื่อ Project ทางขวา เพื่อแสดงตารางเปรียบเทียบราคาจากทุก Supplier ในโปรเจ็กต์นั้น' },
      { label: 'อ่านตาราง', desc: 'แต่ละ Supplier จะแสดง 3 คอลัมน์: FOB (สกุลเงินต้นทาง), FOB THB และ DDP THB — DDP ที่ถูกที่สุดจะ highlight สีเขียว' },
      { label: 'ดู Price History', desc: 'กดที่ตัวเลขราคา (FOB) เพื่อดูประวัติราคาย้อนหลังของ Item นั้นจาก Supplier นั้น' },
      { label: 'จัดการ / ลบข้อมูล', desc: 'กดปุ่ม ⋯ ถัดจากชื่อ Supplier เพื่อดูรายการไฟล์ที่อัปโหลด และลบแต่ละ Batch ได้ (ต้องปลดล็อกก่อน)' },
      { label: 'Export Excel', desc: 'กดปุ่ม "Export Excel" เพื่อดาวน์โหลดตารางเปรียบเทียบ — รองรับการกรองด้วย keyword ก่อน Export' },
    ],
    tips: [
      'ถ้าอัปโหลด Supplier เดิมซ้ำ ระบบจะเก็บทั้งสอง batch — ใช้ checkbox "แทนที่ข้อมูลเดิม" ถ้าต้องการอัปเดตราคาใหม่',
      'ราคา Description แสดงจากไฟล์ล่าสุดเสมอ',
      'Estimate Rates (CNY/THB, USD/THB, DDP ×) สามารถแก้ไขได้ที่ด้านบนหน้า (ต้องปลดล็อก)',
    ],
  },
  {
    href: '/po-builder',
    label: 'PO Builder',
    color: 'bg-rose-50 border-rose-200',
    iconBg: 'bg-rose-600',
    icon: '🏗️',
    title: 'PO Builder — สร้างใบสั่งซื้อ (PO) พร้อม Export Excel',
    subtitle: 'เลือก Supplier + ใส่ Item Code เพื่อดึงราคา FOB แล้ว Export เป็นไฟล์ PO ตาม Template',
    steps: [
      { label: 'เลือก Supplier', desc: 'เลือก Supplier จาก dropdown — ราคาจะดึงมาจากข้อมูลล่าสุดในระบบของ Supplier นั้น (ค้นหาข้ามทุก Project)' },
      { label: 'เพิ่ม Item Code', desc: 'พิมพ์ Item Code ในช่องแล้วกด Enter หรือกดปุ่ม "+ เพิ่ม" ระบบจะดึง Description และราคา FOB มาใส่ให้อัตโนมัติ' },
      { label: 'กรอก QTY', desc: 'ใส่จำนวน (QTY) แต่ละรายการ ระบบจะคำนวณ Total = QTY × FOB Price ให้อัตโนมัติ' },
      { label: 'Export PO Excel', desc: 'กดปุ่ม "Export PO Excel" เพื่อดาวน์โหลดไฟล์ Excel ที่มีรูปแบบตรงตาม PO Template — มี 30 rows, แถว TOTAL, Remark และลายเซ็นครบ' },
    ],
    tips: [
      'ถ้า Item Code ไม่พบในระบบ จะแสดงว่า "ไม่พบราคา" — ยังสามารถ Export ได้ แต่ช่อง UNIT PRICE จะว่าง',
      'การเปลี่ยน Supplier จะอัปเดตราคาของ Item ทั้งหมดในลิสต์ให้อัตโนมัติ',
      'กดปุ่ม ✕ เพื่อลบ Item ออกจากลิสต์',
    ],
  },
]

export default function GuidePage() {
  const [, setUnlocked] = useState(false)

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-6 text-sm sticky top-0 z-20 shadow-sm flex-wrap">
        <span className="font-bold text-gray-900">Import PO</span>
        <Link href="/" className="text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <div className="relative group">
          <span className="text-gray-500 cursor-default hover:text-gray-800">Summary ▾</span>
          <div className="absolute left-0 top-full pt-1 hidden group-hover:block z-50">
            <div className="bg-white border border-gray-200 rounded-lg shadow-lg py-1 min-w-[150px]">
              <Link href="/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">Item Summary</Link>
              <Link href="/qc/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">QC Summary</Link>
            </div>
          </div>
        </div>
        <Link href="/qc" className="text-gray-500 hover:text-gray-800 transition-colors">QC Report</Link>
        <Link href="/guide" className="text-blue-600">Guide</Link>
        <div className="ml-auto">
          <LockButton onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        </div>
      </nav>

      <div className="max-w-4xl mx-auto px-6 py-10">
        <div className="mb-8">
          <h1 className="text-2xl font-bold text-gray-900">คู่มือการใช้งาน</h1>
          <p className="text-sm text-gray-500 mt-1">อธิบายวิธีการใช้งานแต่ละหน้าของระบบ Import PO</p>
        </div>

        {/* Quick nav */}
        <div className="flex flex-wrap gap-2 mb-8">
          {sections.map(s => (
            <a key={s.href} href={`#${s.label}`}
              className="px-3 py-1.5 rounded-full border border-gray-200 bg-white text-xs text-gray-600 hover:border-blue-300 hover:text-blue-600 transition-colors">
              {s.label}
            </a>
          ))}
        </div>

        <div className="space-y-8">
          {sections.map(s => (
            <div key={s.href} id={s.label} className={`rounded-2xl border-2 ${s.color} p-6`}>
              {/* Header */}
              <div className="flex items-start gap-4 mb-5">
                <div className={`${s.iconBg} rounded-xl w-12 h-12 flex items-center justify-center text-2xl flex-shrink-0`}>
                  {s.icon}
                </div>
                <div className="flex-1">
                  <div className="flex items-center gap-3 flex-wrap">
                    <h2 className="text-base font-bold text-gray-900">{s.title}</h2>
                    <Link href={s.href}
                      className="text-xs text-blue-600 hover:underline border border-blue-200 rounded-full px-2 py-0.5 bg-white">
                      เปิดหน้า →
                    </Link>
                  </div>
                  <p className="text-sm text-gray-600 mt-1">{s.subtitle}</p>
                </div>
              </div>

              {/* Steps */}
              <div className="space-y-3 mb-4">
                {s.steps.map((step, i) => (
                  <div key={i} className="flex gap-3 bg-white/70 rounded-xl px-4 py-3">
                    <span className="flex-shrink-0 w-6 h-6 rounded-full bg-gray-200 text-gray-600 text-xs font-bold flex items-center justify-center mt-0.5">
                      {i + 1}
                    </span>
                    <div>
                      <p className="text-sm font-semibold text-gray-800">{step.label}</p>
                      <p className="text-xs text-gray-500 mt-0.5 leading-relaxed">{step.desc}</p>
                    </div>
                  </div>
                ))}
              </div>

              {/* Tips */}
              {s.tips.length > 0 && (
                <div className="bg-white/60 rounded-xl px-4 py-3 space-y-1">
                  <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1.5">เพิ่มเติม</p>
                  {s.tips.map((tip, i) => (
                    <p key={i} className="text-xs text-gray-600 leading-relaxed">
                      <span className="text-gray-400 mr-1">•</span>{tip}
                    </p>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>

        {/* Password note */}
        <div className="mt-8 rounded-xl border border-gray-200 bg-white p-5">
          <p className="text-sm font-semibold text-gray-800 mb-1">🔒 ระบบล็อก</p>
          <p className="text-xs text-gray-500 leading-relaxed">
            บางฟังก์ชัน เช่น การแก้ไขข้อมูล, ลบไฟล์, หรือเปลี่ยนค่า Exchange Rate จะต้องปลดล็อกก่อน
            กดปุ่ม <span className="font-medium text-gray-700">"ล็อก — คลิกเพื่อแก้ไข"</span> มุมขวาบนแล้วใส่รหัสผ่าน
            เมื่อปลดล็อกแล้ว ปุ่มจะเปลี่ยนเป็นสีเขียว และสามารถแก้ไขได้จนกว่าจะ refresh หน้า
          </p>
        </div>

        <p className="text-center text-xs text-gray-400 mt-8">Import PO System · พัฒนาสำหรับทีม Purchasing and Import</p>
      </div>
    </div>
  )
}
