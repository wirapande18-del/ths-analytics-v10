import React, { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { createClient } from '@supabase/supabase-js'
import { RefreshCw, Database, Users, UserCheck, UserX, UserPlus, Search, Download, Clock, Trash2, Upload, BarChart3, TrendingUp } from 'lucide-react'

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || ''
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY || ''
const supabase = supabaseUrl && supabaseKey ? createClient(supabaseUrl, supabaseKey) : null
const TYPE_ALLOWED = ['OIL', 'SBE']

function norm(v){ return String(v ?? '').trim() }
function upper(v){ return norm(v).toUpperCase() }
function excelDateToJS(v){
  if (!v) return null
  if (v instanceof Date) return v.toISOString().slice(0,10)
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v)
    if (!d) return null
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`
  }
  const d = new Date(v)
  if (!isNaN(d)) return d.toISOString().slice(0,10)
  return null
}
function toNumber(v){
  if (typeof v === 'number') return v
  const s = norm(v).replace(/[^0-9,-]/g,'').replace(',', '.')
  return Number(s) || 0
}
function keyClean(v){ return upper(v).replace(/[^A-Z0-9]/g,'') }
function getValue(row, keys){
  const all = Object.keys(row || {})
  for (const k of keys){
    const kk = keyClean(k)
    const found = all.find(x => keyClean(x).includes(kk))
    if (found) return row[found]
  }
  return ''
}
function addMonths(date, months){ const d = new Date(date); d.setMonth(d.getMonth()+months); return d }
function dateStr(d){ const x = new Date(d); return isNaN(x) ? '' : x.toISOString().slice(0,10) }
function daysBetween(a,b){ return Math.floor((new Date(b)-new Date(a))/(1000*60*60*24)) }
function monthsEarly(expected, actual){ const diff = daysBetween(actual, expected); return diff > 0 ? Math.max(1, Math.round(diff / 30)) : 0 }

function extractKmNumber(v){
  const n = parseInt(String(v ?? '').replace(/[^0-9]/g,''), 10)
  return Number.isFinite(n) ? n : null
}
function formatKmLabel(n){
  return n ? `${n.toLocaleString('id-ID')} KM` : ''
}
function getJobSbe(base){
  // Ambil dari kolom Excel "Job SBE" (contoh kolom P), bukan dari kilometer kendaraan.
  return norm(base?.job_sbe || getValue(base?.raw_data || {}, ['JOB SBE','JOBSBE','JOB SBE KM','SBE']))
}
function buildNextServiceInfo(base){
  const type = upper(base.repair_type)
  if(type === 'SBE'){
    const jobSbe = getJobSbe(base)
    const jobSbeNumber = extractKmNumber(jobSbe)
    return {
      last_sbe_km: jobSbe,
      next_service: jobSbeNumber ? `SBE ${formatKmLabel(jobSbeNumber + 10000)}` : 'SBE'
    }
  }
  return { last_sbe_km: '', next_service: 'OIL' }
}

function cleanPhoneKey(v){
  let digits = String(v ?? '').replace(/[^0-9]/g,'')
  if(digits.startsWith('0')) digits = '62' + digits.slice(1)
  return digits
}
function getWaCp(row){
  return norm(row?.wa_cp || getValue(row?.raw_data || {}, ['WA CP','WACP','WA','NO HP','NOHP','HP','PHONE','TELP','TELEPON','HANDPHONE']))
}
function mergePhones(...values){
  const out = []
  const seen = new Set()
  values.forEach(v => {
    const text = norm(v)
    if(!text) return
    const parts = text.split(/\n|;|,/).map(x=>norm(x)).filter(Boolean)
    parts.forEach(part => {
      const key = cleanPhoneKey(part)
      if(!key || key.length < 8) return
      if(seen.has(key)) return
      seen.add(key)
      out.push(part)
    })
  })
  return out.join(' / ')
}
function buildPhoneIndex(rows){
  const m = new Map()
  rows.forEach(r => {
    const key = vehicleKey(r)
    if(!key) return
    const wa = getWaCp(r)
    if(!wa) return
    m.set(key, mergePhones(m.get(key), wa))
  })
  return m
}
function firstPhoneForWa(v){
  const first = norm(v).split('/')[0] || ''
  const digits = cleanPhoneKey(first)
  return digits || ''
}


function monthLabel(dateText){
  if(!dateText) return '-'
  const d = new Date(dateText)
  if(isNaN(d)) return '-'
  return d.toLocaleDateString('id-ID', {month:'short', year:'2-digit'})
}
function groupCount(rows, keyFn){
  const m = new Map()
  rows.forEach(r => { const k = keyFn(r) || '-'; m.set(k, (m.get(k)||0)+1) })
  return Array.from(m.entries()).map(([label,value]) => ({label,value})).sort((a,b)=> b.value-a.value || String(a.label).localeCompare(String(b.label)))
}
function leadInfo(row, sampaiTanggal){
  const status = upper(row?.status)
  const telat = Number(row?.telat_hari || 0)
  const awal = Number(row?.early_months || 0)
  if(status.includes('CUSTOMER BARU')) return {key:'baru', label:'Customer Baru', color:'purple', sort:6}
  if(awal > 0 || status.includes('LEBIH AWAL')) return {key:'awal', label:`Lebih Awal${awal?` ${awal} bln`:''}`, color:'blue', sort:5}
  if(status.includes('TELAT') && telat > 0) return {key:'datang_telat', label:`Datang Telat ${telat} hari`, color:'orange', sort:4}
  if(status.includes('SUDAH DATANG')) return {key:'sudah', label:'Sudah Datang', color:'green', sort:4}
  if(status.includes('LOST')) return {key:'lost', label:'Lost Customer', color:'darkred', sort:0}
  if(status.includes('BELUM') || status.includes('DUE')){
    if(telat > 30) return {key:'telat30', label:'Telat >30 hari', color:'red', sort:1}
    if(telat > 0) return {key:'telat', label:`Telat ${telat} hari`, color:'orange', sort:2}
    return {key:'bulanini', label:'Due bulan ini', color:'yellow', sort:3}
  }
  return {key:'netral', label: row?.status || 'Netral', color:'gray', sort:9}
}
function leadBucket(row){
  const early = Number(row?.early_months || 0)
  const late = Number(row?.telat_hari || 0)
  if(early >= 3) return 'Lebih awal ≥3 bln'
  if(early >= 1) return 'Lebih awal 1-2 bln'
  if(late > 60) return 'Telat >60 hr'
  if(late > 30) return 'Telat 31-60 hr'
  if(late > 0) return 'Telat 1-30 hr'
  return 'Tepat / bulan ini'
}
function uniqueSignature(r){
  return [dateStr(r.repair_date || ''), keyClean(r.police_no), keyClean(r.chassis_no), upper(r.repair_type), upper(r.sa), upper(r.km)].join('|')
}

function parseAllSheets(file, onDone, onError){
  const reader = new FileReader()
  reader.onload = evt => {
    try{
      const wb = XLSX.read(evt.target.result)
      const json = wb.SheetNames.flatMap(sheetName => {
        const ws = wb.Sheets[sheetName]
        return XLSX.utils.sheet_to_json(ws, {defval:''}).map(row => ({...row, __sheet: sheetName}))
      })
      const payload = json.map(row => {
        const repairType = upper(getValue(row, ['REPAIRTYPE','RTYPE','TYPE','TYPESBE']))
        return {
          repair_date: excelDateToJS(getValue(row, ['REPAIRDATE','SERVICEDATE','DATE','TANGGAL'])),
          police_no: upper(getValue(row, ['POLICENO','NOPOL','PLAT','POLICE'])),
          chassis_no: upper(getValue(row, ['NORANGKA','NO RANGKA','CHASSIS','CHASSISNO','FRAMENO','VIN','NOKA'])),
          customer_name: norm(getValue(row, ['CUSTOMERNAME','CUSTOMER','NAMA'])),
          repair_type: repairType,
          sa: norm(getValue(row, ['SA','SERVICEADVISOR','ADVISOR'])),
          tts: norm(getValue(row, ['TTS','TECHNICIAN','TEKNISI','TECH'])),
          km: norm(getValue(row, ['KM','ODOMETER'])),
          omzet: toNumber(getValue(row, ['OMZET','AMOUNT','TOTAL','REVENUE','LABOR','PART'])),
          source_file: file.name,
          source_sheet: row.__sheet,
          raw_data: row
        }
      }).filter(r => r.repair_date && r.police_no && TYPE_ALLOWED.includes(r.repair_type))
      onDone(payload, wb.SheetNames.length)
    }catch(err){ onError(err) }
  }
  reader.onerror = () => onError(new Error('Gagal membaca file.'))
  reader.readAsArrayBuffer(file)
}
function vehicleKey(r){
  const plat = upper(r.police_no).replace(/[^A-Z0-9]/g,'')
  const rangka = upper(r.chassis_no).replace(/[^A-Z0-9]/g,'')
  return plat || (rangka ? `RANGKA:${rangka}` : '')
}
function latestByVehicle(rows){
  const m = new Map()
  rows.forEach(r => {
    const key = vehicleKey(r)
    if(!key) return
    if(!m.has(key) || new Date(m.get(key).repair_date) < new Date(r.repair_date)) m.set(key, r)
  })
  return m
}
function earliestByVehicle(rows){
  const m = new Map()
  rows.forEach(r => {
    const key = vehicleKey(r)
    if(!key) return
    if(!m.has(key) || new Date(m.get(key).repair_date) > new Date(r.repair_date)) m.set(key, r)
  })
  return m
}
function filterBase(rows, typeMode, saMode){
  return rows.filter(r => (typeMode === 'ALL' ? TYPE_ALLOWED : [typeMode]).includes(upper(r.repair_type)))
    .filter(r => saMode === 'ALL' || upper(r.sa) === upper(saMode))
    .filter(r => r.repair_date && (r.police_no || r.chassis_no))
}
function sortedServicesByVehicle(rows){
  const m = new Map()
  rows.forEach(r => {
    const key = vehicleKey(r)
    if(!key) return
    if(!m.has(key)) m.set(key, [])
    m.get(key).push(r)
  })
  m.forEach(list => list.sort((a,b)=> new Date(a.repair_date) - new Date(b.repair_date) || upper(a.repair_type).localeCompare(upper(b.repair_type)) || upper(a.km).localeCompare(upper(b.km))))
  return m
}
function buildAnalytics(oldRows, currentRows, period, sampaiTanggal, typeMode, saMode){
  const now = sampaiTanggal ? new Date(sampaiTanggal) : new Date()
  const oldFiltered = filterBase(oldRows, typeMode, saMode).filter(r => new Date(r.repair_date) <= now)
  const currentFiltered = filterBase(currentRows, typeMode, saMode).filter(r => new Date(r.repair_date) <= now)
  const lastOldMap = latestByVehicle(oldFiltered)
  const currentByVehicle = sortedServicesByVehicle(currentFiltered)
  const oldVehicleAllSA = latestByVehicle(filterBase(oldRows, typeMode, 'ALL'))
  const oldPhoneMap = buildPhoneIndex(oldRows)
  const currentPhoneMap = buildPhoneIndex(currentRows)
  const phoneForKey = key => mergePhones(currentPhoneMap.get(key), oldPhoneMap.get(key))
  const pembanding=[], sudahDatang=[], datangAwal=[], belumDatang=[], due=[], lost=[]

  ;[...lastOldMap.values()].forEach(base => {
    const key = vehicleKey(base)
    const expected = addMonths(new Date(base.repair_date), period)
    const expectedTxt = dateStr(expected)
    const serviceInfo = buildNextServiceInfo(base)
    const wa_cp = phoneForKey(key)
    const hariSejak = daysBetween(base.repair_date, now)
    const telatSampaiHariIni = daysBetween(expected, now)
    const isDue = expected <= now
    const currentList = (currentByVehicle.get(key) || []).filter(c => new Date(c.repair_date) >= new Date(base.repair_date) && new Date(c.repair_date) <= now)

    if(isDue) pembanding.push({...base, ...serviceInfo, expected_date: expectedTxt, hari: hariSejak, telat_hari: Math.max(0,telatSampaiHariIni)})

    if(currentList.length){
      currentList.forEach((current, idx) => {
        const early = monthsEarly(expected, current.repair_date)
        const lateAtCome = Math.max(0, daysBetween(expected, current.repair_date))
        let status = 'Sudah Datang'
        if(early > 0) status = `Datang Lebih Awal ${early} bln`
        else if(lateAtCome > 0) status = `Datang Telat ${lateAtCome} hari`
        else status = 'Sudah Datang Tepat Waktu'
        const currentWa = getWaCp(current)
        const row = {
          ...base,
          ...serviceInfo,
          wa_cp: mergePhones(currentWa, wa_cp),
          current_date: current.repair_date,
          current_repair_type: current.repair_type,
          current_sa: current.sa,
          current_km: current.km,
          current_source_file: current.source_file,
          current_source_sheet: current.source_sheet,
          repeat_no_bulan_ini: idx + 1,
          expected_date: expectedTxt,
          early_months: early,
          hari: hariSejak,
          telat_hari: lateAtCome,
          status
        }
        sudahDatang.push(row)
        if(early > 0) datangAwal.push(row)
      })
    } else if(isDue){
      const row = {...base, ...serviceInfo, wa_cp, expected_date: expectedTxt, hari: hariSejak, telat_hari: Math.max(0,telatSampaiHariIni), status:'Due Service / Belum Datang'}
      belumDatang.push(row); due.push(row)
      if(hariSejak >= 180) lost.push({...row, status:'Lost Customer'})
    }
  })

  const customerBaru = currentFiltered
    .filter(r => !oldVehicleAllSA.has(vehicleKey(r)))
    .map(r => ({
      ...r,
      wa_cp: phoneForKey(vehicleKey(r)) || getWaCp(r),
      current_date: r.repair_date,
      current_repair_type: r.repair_type,
      current_sa: r.sa,
      current_km: r.km,
      repeat_no_bulan_ini: 1,
      status:'Customer Baru'
    }))
  return {pembanding,sudahDatang,datangAwal,belumDatang,due,lost,customerBaru,currentFiltered,oldFiltered}
}

async function replaceTable(table, rows){
  if(!supabase) throw new Error('Supabase belum disetting. Isi file .env dulu.')
  const { error: delErr } = await supabase.from(table).delete().neq('id', 0)
  if(delErr) throw delErr
  // Insert semua baris, dibagi batch supaya file besar tidak gagal.
  for(let i=0;i<rows.length;i+=500){
    const { error } = await supabase.from(table).insert(rows.slice(i,i+500))
    if(error) throw error
  }
}

function normalizeComparable(r){
  const raw = r?.raw_data || {}
  return JSON.stringify({
    repair_date: dateStr(r?.repair_date || ''),
    police_no: upper(r?.police_no),
    chassis_no: upper(r?.chassis_no),
    customer_name: norm(r?.customer_name),
    repair_type: upper(r?.repair_type),
    sa: norm(r?.sa),
    tts: norm(r?.tts),
    km: norm(r?.km),
    omzet: Number(r?.omzet || 0),
    source_sheet: norm(r?.source_sheet),
    raw_data: raw
  })
}
function stripForSave(r){
  const {id, created_at, updated_at, lead, ...clean} = r || {}
  return clean
}
async function syncUpdateTable(table, rows){
  if(!supabase) throw new Error('Supabase belum disetting. Isi file .env dulu.')

  // Jika file yang sama terupload dua kali dalam satu import, baris terakhir dipakai.
  // Kunci dibuat dari tanggal service + plat/rangka + repair type + SA + KM.
  const incomingMap = new Map()
  rows.forEach(r => {
    const sig = uniqueSignature(r)
    if(sig.replace(/\|/g,'').trim()) incomingMap.set(sig, stripForSave(r))
  })
  const incoming = Array.from(incomingMap.entries())

  const existing = await fetchRows(table)
  const existingMap = new Map(existing.map(r => [uniqueSignature(r), r]))
  const inserts = []
  const updates = []
  let skipped = 0
  let duplicateInFile = rows.length - incoming.length

  incoming.forEach(([sig,row]) => {
    const old = existingMap.get(sig)
    if(!old){
      inserts.push(row)
      return
    }
    if(normalizeComparable(old) === normalizeComparable(row)){
      skipped += 1
      return
    }
    updates.push({id: old.id, row})
  })

  for(let i=0;i<inserts.length;i+=500){
    const { error } = await supabase.from(table).insert(inserts.slice(i,i+500))
    if(error) throw error
  }
  for(let i=0;i<updates.length;i+=100){
    const batch = updates.slice(i,i+100)
    await Promise.all(batch.map(u => supabase.from(table).update(u.row).eq('id', u.id).then(({error}) => { if(error) throw error })))
  }
  return {inserted:inserts.length, updated:updates.length, skipped, duplicateInFile, total:existing.length + inserts.length}
}

async function fetchRows(table){
  if(!supabase) return []
  // Supabase/PostgREST sering membatasi 1000 baris per request.
  // v7 mengambil data pakai pagination agar semua baris terbaca.
  const all = []
  const pageSize = 1000
  for(let from=0; ; from += pageSize){
    const to = from + pageSize - 1
    const { data, error } = await supabase.from(table).select('*').order('repair_date',{ascending:false}).range(from, to)
    if(error) throw error
    all.push(...(data || []))
    if(!data || data.length < pageSize) break
  }
  return all
}

export default function App(){
  const [oldRows,setOldRows] = useState([])
  const [currentRows,setCurrentRows] = useState([])
  const [period,setPeriod] = useState(6)
  const [typeMode,setTypeMode] = useState('ALL')
  const [saMode,setSaMode] = useState('ALL')
  const [sampaiTanggal,setSampaiTanggal] = useState(new Date().toISOString().slice(0,10))
  const [detail,setDetail] = useState('due')
  const [search,setSearch] = useState('')
  const [msg,setMsg] = useState('')
  const [loading,setLoading] = useState(false)

  async function loadData(){
    try{ setLoading(true); const [oldData, currentData] = await Promise.all([fetchRows('job_history_old'), fetchRows('job_history_current')]); setOldRows(oldData); setCurrentRows(currentData); setMsg(`Data dari Supabase terbaca: Data Lama ${oldData.length.toLocaleString('id-ID')} baris, Bulan Berjalan ${currentData.length.toLocaleString('id-ID')} baris.`) }
    catch(err){ setMsg(err.message) } finally{ setLoading(false) }
  }
  useEffect(()=>{ loadData() },[])

  function uploadTo(table, label, mode='append'){
    return e => {
      const file = e.target.files?.[0]
      if(!file) return
      const isAppend = mode === 'append'
      const warning = isAppend
        ? `${label} akan DISINKRONKAN ke tabel ${table}. Data lama tidak dihapus. Data baru ditambahkan, data yang berubah diperbarui, data sama dilewati. Lanjutkan?`
        : `${label} akan REPLACE data di tabel ${table}. Data Lama/Pembanding tetap aman. Lanjutkan?`
      if(!confirm(warning)) return
      setLoading(true); setMsg(`Membaca ${label}...`)
      parseAllSheets(file, async (rows, sheets)=>{
        try{
          if(isAppend){
            const res = await syncUpdateTable(table, rows)
            setMsg(`${label} berhasil disinkronkan: ${res.inserted.toLocaleString('id-ID')} data baru ditambahkan, ${res.updated.toLocaleString('id-ID')} data lama diperbarui, ${res.skipped.toLocaleString('id-ID')} data sama dilewati, ${res.duplicateInFile.toLocaleString('id-ID')} dobel dalam file dirapikan. Total tabel sekarang ${res.total.toLocaleString('id-ID')} baris.`)
          }else{
            await replaceTable(table, rows)
            setMsg(`${label} berhasil disimpan FULL ke Supabase: ${rows.length.toLocaleString('id-ID')} baris OIL/SBE dari ${sheets} sheet.`)
          }
          await loadData()
        }
        catch(err){ setMsg(err.message) } finally{ setLoading(false); e.target.value='' }
      }, err=>{ setMsg(err.message); setLoading(false); e.target.value='' })
    }
  }
  async function resetCurrent(){ if(!confirm('DELETE ALL data Bulan Berjalan di Supabase? Data Lama tetap aman.')) return; try{ setLoading(true); await replaceTable('job_history_current', []); setMsg('DELETE ALL Bulan Berjalan berhasil.'); await loadData() } catch(err){ setMsg(err.message) } finally{ setLoading(false) } }
  async function resetAll(){ if(!confirm('Hapus Data Lama dan Bulan Berjalan di Supabase?')) return; try{ setLoading(true); await replaceTable('job_history_current', []); await replaceTable('job_history_old', []); setMsg('Semua data Supabase sudah dikosongkan.'); await loadData() } catch(err){ setMsg(err.message) } finally{ setLoading(false) } }

  const saList = useMemo(()=> ['ALL', ...Array.from(new Set([...oldRows,...currentRows].map(r=>norm(r.sa)).filter(Boolean))).sort()], [oldRows,currentRows])
  const a = useMemo(()=> buildAnalytics(oldRows,currentRows,period,sampaiTanggal,typeMode,saMode), [oldRows,currentRows,period,sampaiTanggal,typeMode,saMode])
  const detailRows = {sudah:a.sudahDatang, awal:a.datangAwal, belum:a.belumDatang, baru:a.customerBaru, due:a.due, lost:a.lost}[detail] || []
  const shown = detailRows.filter(r => !search || upper(r.police_no).includes(upper(search)) || upper(r.chassis_no).includes(upper(search)) || upper(r.customer_name).includes(upper(search)) || upper(r.sa).includes(upper(search)) || upper(r.wa_cp).includes(upper(search)))
  const repeatRate = a.pembanding.length ? Math.round((a.sudahDatang.length/a.pembanding.length)*100) : 0
  const allDetail = useMemo(()=> [...a.sudahDatang, ...a.belumDatang, ...a.customerBaru].map(r => ({...r, lead: leadInfo(r, sampaiTanggal)})), [a, sampaiTanggal])
  const statusChart = useMemo(()=> groupCount(allDetail, r => leadInfo(r, sampaiTanggal).key).map(x => {
    const sample = allDetail.find(r => leadInfo(r, sampaiTanggal).key === x.label)
    return { ...x, label: sample ? leadInfo(sample, sampaiTanggal).label.replace(/ \d+ hari| \d+ bln/g,'') : x.label, detail: sample ? leadInfo(sample, sampaiTanggal).key : x.label }
  }), [allDetail, sampaiTanggal])
  const typeChart = useMemo(()=> groupCount(allDetail, r => upper(r.repair_type)), [allDetail])
  const monthChart = useMemo(()=> groupCount(allDetail, r => monthLabel(r.expected_date || r.current_date || r.repair_date)).slice(0,12).reverse(), [allDetail])
  const leadBucketChart = useMemo(()=> groupCount(allDetail, leadBucket), [allDetail])

  function exportExcel(){
    const makeDetail = rows => rows.map(r => {
      const info = leadInfo(r, sampaiTanggal)
      return {Plat:r.police_no,'No Rangka':r.chassis_no||'',Customer:r.customer_name,'WA CP':r.wa_cp||'',SA:r.current_sa||r.sa||'','Last Service / Tgl Pembanding':r.repair_date,'SBE Terakhir':r.last_sbe_km||'','Next Service':r.next_service||'','Datang Lagi Bulan Berjalan':r.current_date||'','Service Ke Bulan Ini':r.repeat_no_bulan_ini||'','Estimasi Jadwal':r.expected_date||'',Type:r.current_repair_type||r.repair_type,Status:r.status||'','Status Warna':info.label,'Lead Bucket':leadBucket(r),'Lebih Awal':r.early_months?`${r.early_months} bulan`:'','Hari Sejak Service':r.hari||'','Telat Hari':r.telat_hari||'','KM Datang':r.current_km||'', 'File Sumber':r.current_source_file||r.source_file||''}
    })
    const wb = XLSX.utils.book_new()
    const summary = [
      ['Dashboard THS Analytics', ''],
      ['Sampai tanggal', sampaiTanggal],
      ['Periode service', `${period} bulan`],
      ['Repair type', typeMode],
      ['SA', saMode === 'ALL' ? 'Semua SA' : saMode],
      [],
      ['Indikator', 'Jumlah'],
      ['Customer Pembanding Due', a.pembanding.length],
      ['Sudah Datang', a.sudahDatang.length],
      ['Datang Lebih Awal', a.datangAwal.length],
      ['Belum Datang', a.belumDatang.length],
      ['Customer Baru', a.customerBaru.length],
      ['Lost Customer', a.lost.length],
      ['Repeat Rate', `${repeatRate}%`],
      [],
      ['Status Warna / Lead', 'Jumlah'],
      ...statusChart.map(x => [x.label, x.value]),
      [],
      ['Lead Bucket', 'Jumlah'],
      ...leadBucketChart.map(x => [x.label, x.value]),
      [],
      ['Repair Type', 'Jumlah'],
      ...typeChart.map(x => [x.label, x.value])
    ]
    const wsSummary = XLSX.utils.aoa_to_sheet(summary)
    wsSummary['!cols'] = [{wch:28},{wch:18}]
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Dashboard')
    const wsDetail = XLSX.utils.json_to_sheet(makeDetail(shown))
    wsDetail['!autofilter'] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(wsDetail['!ref'] || 'A1:Q1')) }
    wsDetail['!cols'] = [{wch:13},{wch:20},{wch:24},{wch:20},{wch:14},{wch:15},{wch:16},{wch:18},{wch:15},{wch:12},{wch:15},{wch:10},{wch:24},{wch:18},{wch:18},{wch:12},{wch:12},{wch:10},{wch:12},{wch:24}]
    XLSX.utils.book_append_sheet(wb, wsDetail, 'Detail Filter')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(makeDetail(allDetail)), 'Semua Data')
    XLSX.writeFile(wb, `THS_Dashboard_Lead_${period}bulan_${sampaiTanggal}.xlsx`)
  }

  function setDetailByStatusKey(key){
    if(['awal'].includes(key)) setDetail('awal')
    else if(['telat30','telat','bulanini'].includes(key)) setDetail('belum')
    else if(key==='sudah') setDetail('sudah')
    else if(key==='baru') setDetail('baru')
    else if(key==='lost') setDetail('lost')
  }

  return <div className="page">
    <header><div><h1>THS Analytics Online v9 Smart Update</h1><p>Supabase + data lama aman + smart update + warna lead + grafik + export Excel filter</p></div><button onClick={loadData}><RefreshCw size={16}/> Refresh</button></header>
    <section className="panel uploadGrid">
      <div className="uploadBox"><h3><Database size={18}/> Upload Data Lama / Pembanding</h3><p>Isi histori sampai bulan lalu. Mode ini <b>smart update</b>: data lama tidak dihapus, data baru ditambah, data berubah diperbarui, duplikat dilewati.</p><input type="file" accept=".xlsx,.xls" onChange={uploadTo('job_history_old','Data Lama / Pembanding','append')}/><b>{oldRows.length.toLocaleString('id-ID')} baris aktif</b></div>
      <div className="uploadBox current"><h3><Upload size={18}/> Upload Bulan Berjalan</h3><p>Khusus bulan aktif. Mode ini juga <b>smart update</b>, jadi upload ulang tidak membuat data dobel.</p><input type="file" accept=".xlsx,.xls" onChange={uploadTo('job_history_current','Bulan Berjalan')}/><b>{currentRows.length.toLocaleString('id-ID')} baris aktif</b><button className="danger" onClick={resetCurrent}><Trash2 size={15}/> Delete All Bulan Berjalan</button></div>
    </section>
    <section className="panel controls"><label>Periode Service <select value={period} onChange={e=>setPeriod(Number(e.target.value))}><option value={6}>6 Bulan</option><option value={12}>12 Bulan / 1 Tahun</option></select></label><label>Sampai Tanggal <input type="date" value={sampaiTanggal} onChange={e=>setSampaiTanggal(e.target.value)}/></label><label>Repair Type <select value={typeMode} onChange={e=>setTypeMode(e.target.value)}><option value="ALL">OIL + SBE</option><option value="OIL">OIL</option><option value="SBE">SBE</option></select></label><label>SA <select value={saMode} onChange={e=>setSaMode(e.target.value)}>{saList.map(sa=><option key={sa} value={sa}>{sa==='ALL'?'Semua SA':sa}</option>)}</select></label><button className="danger secondary" onClick={resetAll}>Reset Semua Data Supabase</button></section>
    {msg && <div className="msg">{msg}</div>}{loading && <div className="msg">Sedang proses, tunggu sebentar...</div>}
    <section className="cards"><Card icon={<Database/>} title="Customer Pembanding Due" value={a.pembanding.length}/><Card icon={<UserCheck/>} title="Sudah Datang" value={a.sudahDatang.length} onClick={()=>setDetail('sudah')}/><Card icon={<Clock/>} title="Datang Lebih Awal" value={a.datangAwal.length} onClick={()=>setDetail('awal')}/><Card icon={<UserX/>} title="Belum Datang" value={a.belumDatang.length} onClick={()=>setDetail('belum')}/><Card icon={<UserPlus/>} title="Customer Baru" value={a.customerBaru.length} onClick={()=>setDetail('baru')}/><Card icon={<Users/>} title="Repeat Rate" value={repeatRate+'%'}/><Card icon={<UserX/>} title="Lost Customer" value={a.lost.length} onClick={()=>setDetail('lost')}/></section>
    <section className="panel chartPanel"><div className="chartTitle"><h2><BarChart3 size={20}/> Grafik Dashboard</h2><p>Klik batang grafik status untuk membuka detail customer.</p></div><div className="chartGrid"><MiniBar title="Status Lead / Warna" data={statusChart} onBarClick={x=>setDetailByStatusKey(x.detail)}/><MiniBar title="Lead Time" data={leadBucketChart}/><MiniBar title="Repair Type" data={typeChart}/><MiniBar title="Estimasi per Bulan" data={monthChart}/></div></section>
    <section className="panel summary"><h2>Jawaban untuk Atasan</h2><p>Per tanggal <b>{new Date(sampaiTanggal).getDate()}</b>, dari Data Lama ditemukan <b>{a.pembanding.length}</b> customer yang masuk jadwal service {period} bulan. Yang sudah datang di Bulan Berjalan <b>{a.sudahDatang.length}</b> record service, datang lebih awal <b>{a.datangAwal.length}</b> record, belum datang/due <b>{a.belumDatang.length}</b>, customer baru <b>{a.customerBaru.length}</b>. Repeat rate saat ini <b>{repeatRate}%</b>.</p>{currentRows.length===0 && <p className="note">Catatan: Data Bulan Berjalan masih kosong, jadi dashboard hanya menghitung customer yang waktunya service dari Data Lama.</p>}</section>
    <section className="panel"><div className="detailHead"><div><h2>Detail Data</h2><div className="tabs"><button className={detail==='sudah'?'active':''} onClick={()=>setDetail('sudah')}>Sudah Datang</button><button className={detail==='awal'?'active':''} onClick={()=>setDetail('awal')}>Datang Lebih Awal</button><button className={detail==='belum'?'active':''} onClick={()=>setDetail('belum')}>Belum Datang</button><button className={detail==='baru'?'active':''} onClick={()=>setDetail('baru')}>Customer Baru</button><button className={detail==='due'?'active':''} onClick={()=>setDetail('due')}>Due Service</button><button className={detail==='lost'?'active':''} onClick={()=>setDetail('lost')}>Lost Customer</button></div></div><div className="actionRight"><button className="exportBtn" onClick={exportExcel}><Download size={16}/> Export Excel</button><div className="search"><Search size={16}/><input placeholder="Cari plat / no rangka / customer / SA / WA" value={search} onChange={e=>setSearch(e.target.value)}/></div></div></div><div className="tableWrap"><table><thead><tr><th>Plat</th><th>No Rangka</th><th>Customer</th><th>WA CP</th><th>SA</th><th>Last Service</th><th>SBE Terakhir</th><th>Next Service</th><th>Datang Lagi</th><th>Ke</th><th>Estimasi Jadwal</th><th>Type</th><th>Status</th><th>Awal</th><th>Hari</th><th>Telat</th></tr></thead><tbody>{shown.map((r,i)=>{ const info = leadInfo(r, sampaiTanggal); return <tr key={i} className={`leadRow lead-${info.color}`}><td>{r.police_no}</td><td>{r.chassis_no||''}</td><td>{r.customer_name}</td><td>{r.wa_cp||''}{firstPhoneForWa(r.wa_cp) && <a className="waBtn" href={`https://wa.me/${firstPhoneForWa(r.wa_cp)}`} target="_blank" rel="noreferrer">WA</a>}</td><td>{r.current_sa||r.sa}</td><td>{r.repair_date}</td><td>{r.last_sbe_km||''}</td><td>{r.next_service||''}</td><td>{r.current_date||''}</td><td>{r.repeat_no_bulan_ini||''}</td><td>{r.expected_date||''}</td><td>{r.current_repair_type||r.repair_type}</td><td><span className={`badge ${info.color}`}>{info.label}</span><small className="statusText">{r.status||''}</small></td><td>{r.early_months?`${r.early_months} bln`:''}</td><td>{r.hari||''}</td><td>{r.telat_hari?`${r.telat_hari} hr`:''}</td></tr>})}</tbody></table></div></section>
  </div>
}
function MiniBar({title,data,onBarClick}){
  const max = Math.max(1, ...data.map(d=>d.value || 0))
  return <div className="miniChart"><h3><TrendingUp size={16}/> {title}</h3>{data.length===0 && <p className="emptyChart">Belum ada data</p>}{data.map((d,i)=><button key={i} className="barLine" onClick={()=>onBarClick && onBarClick(d)}><span className="barLabel">{d.label}</span><span className="barTrack"><i style={{width:`${Math.max(5, (d.value/max)*100)}%`}}/></span><b>{d.value.toLocaleString('id-ID')}</b></button>)}</div>
}
function Card({icon,title,value,onClick}){ return <button className="card" onClick={onClick}><span>{icon}</span><small>{title}</small><b>{value}</b></button> }
