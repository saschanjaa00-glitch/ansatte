import { useMemo, useState } from 'react'
import {
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import { jsPDF } from 'jspdf'
import * as XLSX from 'xlsx'
import './App.css'

type Employee = {
  id: string
  navn: string
  telefon: string
  fodselsdato: Date
  avdeling: string
  stillingskode: string
  hovedstilling: string
  skolearsplanlegging: string
}

type BirthdayRow = Employee & {
  alderDetteAret: number
  milestone: boolean
}

type BirthdayExportFormat = 'ics' | 'pdf'

type ParsedEmployeeRow = {
  navn: string
  telefon: string
  fodselsdato: Date
  avdeling: string
  stillingskode: string
  arbeidstittel: string
  hovedstilling: string
  skolearsplanlegging: string
  stillingskoder: string[]
}

type ImportStillingskodeOption = {
  id: string
  stillingskode: string
  arbeidstittel: string
}

function isRoundBirthdayAge(age: number): boolean {
  return age >= 20 && age % 10 === 0
}
const ALLOWED_PREFIXES = ['65', '66', '69', '72', '75', '76', '79', '80', '84']
const ALLOWED_EXACT_CODES = ['995401', '995405', '995101']
const REQUIRED_COLUMN_MATCHERS: Array<{ label: string; keys: string[] }> = [
  { label: 'navn', keys: ['navn', 'name'] },
  { label: 'telefon/mobil', keys: ['telefon', 'mobil', 'mobile'] },
  {
    label: 'fødselsdato',
    keys: ['fødselsdato', 'fodselsdato', 'fodsels', 'født', 'fodt', 'birth'],
  },
  { label: 'avdeling', keys: ['avdeling'] },
  { label: 'stillingskoder', keys: ['stillingskode', 'stillingskoder'] },
  {
    label: 'hovedjobb/hovedstilling/arbeidstittel',
    keys: ['hovedstilling', 'hovedjobb', 'hovedjob', 'arbeidstittel', 'stillingstittel'],
  },
  {
    label: 'inkluder i timeplanlegging/skoleårsplanlegging',
    keys: [
      'inkludert i skoleårsplanlegging',
      'inkludertiskolearsplanlegging',
      'skoleårsplanlegging',
      'skolearsplanlegging',
      'inkluder i timeplanlegging',
      'inkluderitimeplanlegging',
      'inkludertimeplanlegging',
      'timeplanlegging',
    ],
  },
]

function normalizeHeader(value: string): string {
  return value
    .toLowerCase()
    .replace(/æ/g, 'ae')
    .replace(/ø/g, 'o')
    .replace(/å/g, 'a')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '')
}

function isValidDateParts(year: number, month: number, day: number): boolean {
  if (year < 1900 || year > 2100) {
    return false
  }
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return false
  }
  const candidate = new Date(year, month - 1, day)
  return (
    candidate.getFullYear() === year &&
    candidate.getMonth() === month - 1 &&
    candidate.getDate() === day
  )
}

function buildDate(year: number, month: number, day: number): Date | null {
  if (!isValidDateParts(year, month, day)) {
    return null
  }
  return new Date(year, month - 1, day)
}

function expandTwoDigitYear(twoDigitYear: number): number {
  return twoDigitYear >= 30 ? 1900 + twoDigitYear : 2000 + twoDigitYear
}

function toDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value
  }

  if (typeof value === 'number') {
    if (Number.isFinite(value) && value > 1_000_000_000 && value < 99_999_999_999) {
      return new Date(value * 1000)
    }
    if (Number.isFinite(value) && value >= 100_000_000_000) {
      return new Date(value)
    }

    const parsed = XLSX.SSF.parse_date_code(value)
    if (!parsed) {
      return null
    }
    return buildDate(parsed.y, parsed.m, parsed.d)
  }

  if (typeof value === 'string') {
    const trimmed = value.trim().replace(/\s+/g, ' ')
    if (!trimmed) {
      return null
    }

    if (/^\d+(?:[.,]\d+)?$/.test(trimmed)) {
      const numeric = Number(trimmed.replace(',', '.'))
      if (!Number.isNaN(numeric)) {
        const fromNumber = toDate(numeric)
        if (fromNumber) {
          return fromNumber
        }
      }
    }

    const compactYmd = trimmed.match(/^(\d{4})(\d{2})(\d{2})$/)
    if (compactYmd) {
      const year = Number(compactYmd[1])
      const month = Number(compactYmd[2])
      const day = Number(compactYmd[3])
      const date = buildDate(year, month, day)
      if (date) {
        return date
      }
    }

    const compactDmy = trimmed.match(/^(\d{2})(\d{2})(\d{4})$/)
    if (compactDmy) {
      const day = Number(compactDmy[1])
      const month = Number(compactDmy[2])
      const year = Number(compactDmy[3])
      const date = buildDate(year, month, day)
      if (date) {
        return date
      }
    }

    const slashDmy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s.*)?$/)
    if (slashDmy) {
      const day = Number(slashDmy[1])
      const month = Number(slashDmy[2])
      const yearRaw = Number(slashDmy[3])
      const year = yearRaw < 100 ? expandTwoDigitYear(yearRaw) : yearRaw
      const date = buildDate(year, month, day)
      if (date) {
        return date
      }
    }

    const dmy = trimmed.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})(?:\s.*)?$/)
    if (dmy) {
      const day = Number(dmy[1])
      const month = Number(dmy[2])
      const yearRaw = Number(dmy[3])
      const year = yearRaw < 100 ? expandTwoDigitYear(yearRaw) : yearRaw
      const date = buildDate(year, month, day)
      if (date) {
        return date
      }
    }

    const ymd = trimmed.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})(?:\s.*)?$/)
    if (ymd) {
      const year = Number(ymd[1])
      const month = Number(ymd[2])
      const day = Number(ymd[3])
      const date = buildDate(year, month, day)
      if (date) {
        return date
      }
    }

    const monthNameDate = trimmed
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .match(/^(\d{1,2})\s+([a-z]+)\s+(\d{2,4})$/)

    if (monthNameDate) {
      const monthMap: Record<string, number> = {
        jan: 1,
        januar: 1,
        feb: 2,
        februar: 2,
        mar: 3,
        mars: 3,
        apr: 4,
        april: 4,
        mai: 5,
        jun: 6,
        juni: 6,
        jul: 7,
        juli: 7,
        aug: 8,
        august: 8,
        sep: 9,
        sept: 9,
        september: 9,
        okt: 10,
        oktober: 10,
        nov: 11,
        november: 11,
        des: 12,
        desember: 12,
      }

      const day = Number(monthNameDate[1])
      const month = monthMap[monthNameDate[2]]
      const yearRaw = Number(monthNameDate[3])
      const year = yearRaw < 100 ? expandTwoDigitYear(yearRaw) : yearRaw
      if (month) {
        const date = buildDate(year, month, day)
        if (date) {
          return date
        }
      }
    }

    const isoLike = trimmed.includes('T') ? trimmed : trimmed.replace(' ', 'T')
    const isoDate = new Date(isoLike)
    if (!Number.isNaN(isoDate.getTime())) {
      return new Date(isoDate.getFullYear(), isoDate.getMonth(), isoDate.getDate())
    }

    const parsed = new Date(trimmed)
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
    }
  }

  return null
}

function extractStillingskoder(raw: string): string[] {
  const inParens = [...raw.matchAll(/\((\d+)\)/g)].map((m) => m[1])
  const candidates =
    inParens.length > 0
      ? inParens
      : [...raw.matchAll(/\b(\d{2,})\b/g)].map((m) => m[1])

  return [...new Set(candidates)]
}

function extractArbeidstittelText(raw: string): string {
  const normalized = raw.trim().replace(/\s+/g, ' ')
  if (!normalized) {
    return ''
  }

  const textOnly = normalized
    .replace(/\(\s*\d+\s*\)/g, ' ')
    .replace(/\b\d{2,}\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

  return textOnly || normalized
}

function toImportOptionId(stillingskode: string, arbeidstittel: string): string {
  return `${stillingskode}|||${arbeidstittel}`
}

function isDefaultStillingskode(code: string): boolean {
  return (
    ALLOWED_PREFIXES.some((prefix) => code.startsWith(prefix)) ||
    ALLOWED_EXACT_CODES.includes(code)
  )
}

function isSpecialStillingskode(raw: string): boolean {
  const inParens = [...raw.matchAll(/\((\d+)\)/g)].map((m) => m[1])
  const candidates =
    inParens.length > 0
      ? inParens
      : [...raw.matchAll(/\b(\d{2,})\b/g)].map((m) => m[1])

  return candidates.some((code) => ALLOWED_EXACT_CODES.includes(code))
}

function extractAvdeling(raw: unknown): string {
  if (typeof raw !== 'string') {
    return ''
  }

  const value = raw.trim().replace(/\s+/g, ' ')
  if (!value) {
    return ''
  }

  const byLeader = value.match(/avdelingsleder\s+(.+)$/i)
  if (byLeader?.[1]) {
    return byLeader[1].trim()
  }

  const byAdminLeader = value.match(/administrasjonsleder\s+(.+)$/i)
  if (byAdminLeader?.[1]) {
    return byAdminLeader[1].trim()
  }

  const byAvdeling = value.match(/avdeling\s+(.+)$/i)
  if (byAvdeling?.[1]) {
    return byAvdeling[1].trim()
  }

  return value
}

function safeValue(row: Record<string, unknown>, keys: string[]): unknown {
  const normalizedKeys = keys.map((key) => normalizeHeader(key))
  const entries = Object.entries(row)
  for (const [header, value] of entries) {
    const normalized = normalizeHeader(header)
    if (normalizedKeys.some((key) => normalized.includes(key))) {
      return value
    }
  }
  return ''
}

function getByExactHeader(row: Record<string, unknown>, headers: string[]): unknown {
  const normalizedTargets = headers.map((header) => normalizeHeader(header))
  for (const [header, value] of Object.entries(row)) {
    const normalizedHeader = normalizeHeader(header)
    if (normalizedTargets.includes(normalizedHeader)) {
      return value
    }
  }
  return ''
}

function normalizeDisplayName(rawName: string): string {
  const cleaned = rawName.replace(/\s+/g, ' ').trim()
  if (!cleaned.includes(',')) {
    return cleaned
  }

  const parts = cleaned.split(',').map((part) => part.trim()).filter(Boolean)
  if (parts.length < 2) {
    return cleaned
  }

  const [lastName, ...firstNameParts] = parts
  return `${firstNameParts.join(' ')} ${lastName}`.replace(/\s+/g, ' ').trim()
}

function extractName(row: Record<string, unknown>): string {
  const directName = String(
    getByExactHeader(row, [
      'navn',
      'fulltnavn',
      'fullt navn',
      'fullstendig navn',
      'display name',
      'displayname',
      'name',
    ]) || '',
  ).trim()

  if (directName) {
    return normalizeDisplayName(directName)
  }

  const fornavn = String(getByExactHeader(row, ['fornavn', 'first name', 'firstname']) || '').trim()
  const etternavn = String(getByExactHeader(row, ['etternavn', 'last name', 'lastname', 'surname']) || '').trim()

  const combined = `${fornavn} ${etternavn}`.trim()
  if (combined) {
    return normalizeDisplayName(combined)
  }

  // Fallback only if no better option exists.
  return normalizeDisplayName(String(safeValue(row, ['navn', 'name']) || '').trim())
}

function findMissingRequiredColumns(headers: string[]): string[] {
  const normalizedHeaders = headers.map((header) => normalizeHeader(header))
  return REQUIRED_COLUMN_MATCHERS
    .filter(
      (required) =>
        !required.keys
          .map((key) => normalizeHeader(key))
          .some((key) => normalizedHeaders.some((header) => header.includes(key))),
    )
    .map((required) => required.label)
}

function formatNorwegianDate(date: Date): string {
  return new Intl.DateTimeFormat('nb-NO', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(date)
}

function formatDayMonth(date: Date): string {
  return new Intl.DateTimeFormat('nb-NO', {
    day: '2-digit',
    month: 'long',
  }).format(date)
}

function downloadBlob(blob: Blob, filename: string): void {
  const link = document.createElement('a')
  const url = URL.createObjectURL(blob)
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

function toICSDate(date: Date): string {
  const y = date.getFullYear().toString().padStart(4, '0')
  const m = (date.getMonth() + 1).toString().padStart(2, '0')
  const d = date.getDate().toString().padStart(2, '0')
  return `${y}${m}${d}`
}

function sanitizeICS(text: string): string {
  return text.replace(/,/g, '\\,').replace(/;/g, '\\;').replace(/\n/g, ' ')
}

function toTwoColumnPairs<T>(items: T[]): Array<[T | null, T | null]> {
  const pairs: Array<[T | null, T | null]> = []
  for (let i = 0; i < items.length; i += 2) {
    pairs.push([items[i] ?? null, items[i + 1] ?? null])
  }
  return pairs
}

function App() {
  const [employees, setEmployees] = useState<Employee[]>([])
  const [fileName, setFileName] = useState('')
  const [error, setError] = useState('')
  const [isDragging, setIsDragging] = useState(false)
  const [uploadHint, setUploadHint] = useState('')
  const [inactiveAvdelingslisteIds, setInactiveAvdelingslisteIds] = useState<Set<string>>(new Set())
  const [birthdayExportFormat, setBirthdayExportFormat] = useState<BirthdayExportFormat | null>(null)
  const [selectedBirthdayAvdelinger, setSelectedBirthdayAvdelinger] = useState<Set<string>>(new Set())
  const [inactiveBirthdayExportIds, setInactiveBirthdayExportIds] = useState<Set<string>>(new Set())
  const [pendingImportRows, setPendingImportRows] = useState<ParsedEmployeeRow[]>([])
  const [pendingImportFileName, setPendingImportFileName] = useState('')
  const [defaultStillingskodeOptions, setDefaultStillingskodeOptions] = useState<ImportStillingskodeOption[]>([])
  const [otherStillingskodeOptions, setOtherStillingskodeOptions] = useState<ImportStillingskodeOption[]>([])
  const [selectedImportStillingskoder, setSelectedImportStillingskoder] = useState<Set<string>>(new Set())

  const birthdays = useMemo<BirthdayRow[]>(() => {
    const year = new Date().getFullYear()
    return [...employees]
      .map((employee) => {
        const alderDetteAret = year - employee.fodselsdato.getFullYear()
        return {
          ...employee,
          alderDetteAret,
          milestone: isRoundBirthdayAge(alderDetteAret),
        }
      })
      .sort((a, b) => {
        const monthDiff = a.fodselsdato.getMonth() - b.fodselsdato.getMonth()
        if (monthDiff !== 0) {
          return monthDiff
        }
        const dayDiff = a.fodselsdato.getDate() - b.fodselsdato.getDate()
        if (dayDiff !== 0) {
          return dayDiff
        }
        return a.navn.localeCompare(b.navn, 'nb')
      })
  }, [employees])

  const milestones = useMemo(
    () => birthdays.filter((entry) => entry.milestone),
    [birthdays],
  )

  const reduction57 = useMemo(
    () =>
      birthdays
        .filter((entry) => entry.alderDetteAret >= 57 && entry.alderDetteAret <= 59)
        .sort((a, b) => a.navn.localeCompare(b.navn, 'nb')),
    [birthdays],
  )

  const reduction60 = useMemo(
    () =>
      birthdays
        .filter((entry) => entry.alderDetteAret >= 60)
        .sort((a, b) => a.navn.localeCompare(b.navn, 'nb')),
    [birthdays],
  )

  const birthdayByAvdeling = useMemo(() => {
    const grouped = new Map<string, BirthdayRow[]>()
    for (const birthday of birthdays) {
      const key = birthday.avdeling || 'Ukjent avdeling'
      const existing = grouped.get(key) ?? []
      existing.push(birthday)
      grouped.set(key, existing)
    }

    for (const list of grouped.values()) {
      list.sort((a, b) => a.navn.localeCompare(b.navn, 'nb'))
    }

    return [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], 'nb'))
  }, [birthdays])

  const birthdayAvdelingNames = useMemo(
    () => birthdayByAvdeling.map(([name]) => name),
    [birthdayByAvdeling],
  )

  const isAllBirthdayAvdelingerSelected =
    birthdayAvdelingNames.length > 0 && selectedBirthdayAvdelinger.size === birthdayAvdelingNames.length

  const birthdayAvdelingPeople = useMemo(() => {
    const selectedIds = new Set(selectedBirthdayAvdelinger)
    return birthdayByAvdeling
      .filter(([name]) => selectedIds.has(name))
      .flatMap(([, people]) => people)
  }, [birthdayByAvdeling, selectedBirthdayAvdelinger])

  const birthdayExportPeople = useMemo(
    () => birthdayAvdelingPeople.filter((person) => !inactiveBirthdayExportIds.has(person.id)),
    [birthdayAvdelingPeople, inactiveBirthdayExportIds],
  )

  const byAvdeling = useMemo(() => {
    const grouped = new Map<string, Employee[]>()
    for (const employee of employees) {
      const key = employee.avdeling || 'Ukjent avdeling'
      const existing = grouped.get(key) ?? []
      existing.push(employee)
      grouped.set(key, existing)
    }

    for (const list of grouped.values()) {
      list.sort((a, b) => a.navn.localeCompare(b.navn, 'nb'))
    }

    return [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], 'nb'))
  }, [employees])

  const exportByAvdeling = useMemo(() => {
    const grouped = new Map<string, Employee[]>()
    for (const employee of employees) {
      if (inactiveAvdelingslisteIds.has(employee.id)) {
        continue
      }
      const key = employee.avdeling || 'Ukjent avdeling'
      const existing = grouped.get(key) ?? []
      existing.push(employee)
      grouped.set(key, existing)
    }

    for (const list of grouped.values()) {
      list.sort((a, b) => a.navn.localeCompare(b.navn, 'nb'))
    }

    return [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], 'nb'))
  }, [employees, inactiveAvdelingslisteIds])

  const importPreviewCount = useMemo(
    () =>
      pendingImportRows.filter((row) =>
        selectedImportStillingskoder.has(toImportOptionId(row.stillingskode, row.arbeidstittel)),
      ).length,
    [pendingImportRows, selectedImportStillingskoder],
  )

  const allStillingskodeOptions = useMemo(
    () => [...defaultStillingskodeOptions, ...otherStillingskodeOptions],
    [defaultStillingskodeOptions, otherStillingskodeOptions],
  )

  function resetPendingImport(): void {
    setPendingImportRows([])
    setPendingImportFileName('')
    setDefaultStillingskodeOptions([])
    setOtherStillingskodeOptions([])
    setSelectedImportStillingskoder(new Set())
  }

  async function handleExcelFile(file: File): Promise<void> {
    setError('')
    setUploadHint('')
    resetPendingImport()

    try {
      if (!/\.(xlsx|xls)$/i.test(file.name)) {
        throw new Error('Ugyldig filtype. Velg en Excel-fil (.xlsx eller .xls).')
      }

      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
      const firstSheetName = workbook.SheetNames[0]

      if (!firstSheetName) {
        throw new Error('Fant ingen ark i Excel-filen.')
      }

      const sheet = workbook.Sheets[firstSheetName]
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
        defval: '',
      })

      const headers = Object.keys(rows[0] ?? {})
      const missingColumns = findMissingRequiredColumns(headers)
      if (missingColumns.length > 0) {
        const detectedHeaders = headers.length > 0 ? headers.join(', ') : '(ingen kolonner funnet)'
        throw new Error(
          `Mangler kolonner: ${missingColumns.join(', ')}. Fant kolonner: ${detectedHeaders}.`,
        )
      }

      const mapped: ParsedEmployeeRow[] = []

      for (const row of rows) {
        const navn = extractName(row)
        const telefon = String(
          safeValue(row, ['telefon', 'mobil', 'mobile']) || '',
        ).trim()
        const fodselsdatoRaw = safeValue(row, ['fødselsdato', 'fodselsdato', 'født', 'fodt', 'birth'])
        const avdelingRaw = safeValue(row, ['avdeling'])
        const stillingskode = String(
          safeValue(row, ['stillingskode', 'stillingskoder']) || '',
        ).trim()
        const stillingskoder = extractStillingskoder(stillingskode)
        const arbeidstittelRaw = String(
          safeValue(row, ['arbeidstittel', 'stillingstittel', 'hovedstilling', 'hovedjobb', 'hovedjob']) || '',
        ).trim()
        const arbeidstittel = extractArbeidstittelText(arbeidstittelRaw)
        const hovedstilling = arbeidstittelRaw
        const skolearsplanlegging = String(
          safeValue(row, [
            'inkludert i skoleårsplanlegging',
            'inkludertiskolearsplanlegging',
            'skoleårsplanlegging',
            'skolearsplanlegging',
            'inkluder i timeplanlegging',
            'inkluderitimeplanlegging',
            'timeplanlegging',
          ]) || '',
        ).trim()

        const fodselsdato = toDate(fodselsdatoRaw)

        if (!navn || !fodselsdato || stillingskoder.length === 0) {
          continue
        }

        mapped.push({
          navn,
          telefon,
          fodselsdato,
          avdeling: extractAvdeling(avdelingRaw),
          stillingskode,
          arbeidstittel,
          hovedstilling,
          skolearsplanlegging,
          stillingskoder,
        })
      }

      if (rows.length > 0 && mapped.length === 0) {
        throw new Error(
          'Ingen gyldige ansatte funnet. Kontroller at navn, fødselsdato og stillingskoder er utfylt.',
        )
      }

      const optionsById = new Map<string, ImportStillingskodeOption>()
      for (const employee of mapped) {
        const arbeidstittelValue = employee.arbeidstittel || 'Ukjent'
        const id = toImportOptionId(employee.stillingskode, arbeidstittelValue)
        if (!optionsById.has(id)) {
          optionsById.set(id, {
            id,
            stillingskode: employee.stillingskode,
            arbeidstittel: arbeidstittelValue,
          })
        }
      }

      const allOptions = [...optionsById.values()].sort((a, b) => {
        const titleCompare = a.arbeidstittel.localeCompare(b.arbeidstittel, 'nb')
        if (titleCompare !== 0) {
          return titleCompare
        }
        return a.stillingskode.localeCompare(b.stillingskode, 'nb')
      })

      const defaultCodes = allOptions.filter((option) =>
        extractStillingskoder(option.stillingskode).some((code) => isDefaultStillingskode(code)),
      )
      const otherCodes = allOptions.filter(
        (option) => !extractStillingskoder(option.stillingskode).some((code) => isDefaultStillingskode(code)),
      )

      setPendingImportRows(mapped)
      setPendingImportFileName(file.name)
      setDefaultStillingskodeOptions(defaultCodes)
      setOtherStillingskodeOptions(otherCodes)
      setSelectedImportStillingskoder(new Set(defaultCodes.map((option) => option.id)))
      setUploadHint(`Fil klar: ${file.name}. Velg stillingskoder i popupen for å fullføre import.`)
    } catch (caught) {
      const message =
        caught instanceof Error ? caught.message : 'Ukjent feil ved lesing av fil.'
      setError(message)
      setUploadHint('')
      setEmployees([])
      setInactiveAvdelingslisteIds(new Set())
      setFileName('')
      resetPendingImport()
    }
  }

  function toggleImportStillingskode(optionId: string): void {
    setSelectedImportStillingskoder((prev) => {
      const next = new Set(prev)
      if (next.has(optionId)) {
        next.delete(optionId)
      } else {
        next.add(optionId)
      }
      return next
    })
  }

  function cancelImportStillingskoder(): void {
    resetPendingImport()
    setUploadHint('')
  }

  function confirmImportStillingskoder(): void {
    if (pendingImportRows.length === 0) {
      return
    }

    const filtered: Employee[] = pendingImportRows
      .filter((employee) =>
        selectedImportStillingskoder.has(toImportOptionId(employee.stillingskode, employee.arbeidstittel || 'Ukjent')),
      )
      .map((employee, index) => ({
        id: `${employee.navn}-${employee.fodselsdato.getTime()}-${employee.telefon}-${index}`,
        navn: employee.navn,
        telefon: employee.telefon,
        fodselsdato: employee.fodselsdato,
        avdeling: employee.avdeling,
        stillingskode: employee.stillingskode,
        hovedstilling: employee.hovedstilling,
        skolearsplanlegging: employee.skolearsplanlegging,
      }))

    if (filtered.length === 0) {
      setError('Ingen ansatte matcher valgte stillingskoder.')
      return
    }

    setEmployees(filtered)
    setInactiveAvdelingslisteIds(new Set())
    setFileName(pendingImportFileName)
    setUploadHint(`Fil registrert: ${pendingImportFileName}`)
    setError('')
    resetPendingImport()
  }

  function toggleAvdelingslistePerson(employeeId: string): void {
    setInactiveAvdelingslisteIds((prev) => {
      const next = new Set(prev)
      if (next.has(employeeId)) {
        next.delete(employeeId)
      } else {
        next.add(employeeId)
      }
      return next
    })
  }

  function setAvdelingActiveState(avdeling: string, shouldBeActive: boolean): void {
    const people = byAvdeling.find(([name]) => name === avdeling)?.[1] ?? []
    setInactiveAvdelingslisteIds((prev) => {
      const next = new Set(prev)
      for (const person of people) {
        if (shouldBeActive) {
          next.delete(person.id)
        } else {
          next.add(person.id)
        }
      }
      return next
    })
  }

  function setAllActive(shouldBeActive: boolean): void {
    if (shouldBeActive) {
      setInactiveAvdelingslisteIds(new Set())
      return
    }
    setInactiveAvdelingslisteIds(new Set(employees.map((employee) => employee.id)))
  }

  function onFileChange(event: React.ChangeEvent<HTMLInputElement>): void {
    const file = event.target.files?.[0]
    if (file) {
      void handleExcelFile(file)
    }
  }

  function onDropZoneDragOver(event: React.DragEvent<HTMLLabelElement>): void {
    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
    setIsDragging(true)
  }

  function onDropZoneDragLeave(event: React.DragEvent<HTMLLabelElement>): void {
    if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
      setIsDragging(false)
    }
  }

  function onDropZoneDrop(event: React.DragEvent<HTMLLabelElement>): void {
    event.preventDefault()
    setIsDragging(false)

    const file = event.dataTransfer.files?.[0]
    if (file) {
      void handleExcelFile(file)
    }
  }

  function exportBirthdayPDF(items: BirthdayRow[]): void {
    const doc = new jsPDF()
    const pageWidth = doc.internal.pageSize.getWidth()
    const pageHeight = doc.internal.pageSize.getHeight()
    const margin = 15
    let yPos = margin

    // Title
    doc.setFontSize(18)
    doc.text('Bursdagskalender', margin, yPos)
    yPos += 10

    // Date
    doc.setFontSize(10)
    doc.text(`Generert: ${formatNorwegianDate(new Date())}`, margin, yPos)
    yPos += 8

    // Table header
    doc.setFontSize(9)
    doc.setFont('helvetica', 'bold')
    const col1 = margin
    const col2 = margin + 30
    const col3 = margin + 85
    const col4 = margin + 150
    
    doc.text('Dato', col1, yPos)
    doc.text('Navn', col2, yPos)
    doc.text('Avdeling', col3, yPos)
    doc.text('Alder', col4, yPos)
    yPos += 2
    doc.line(margin, yPos, pageWidth - margin, yPos)
    yPos += 5

    // Table rows
    doc.setFont('helvetica', 'normal')
    for (const person of items) {
      if (yPos > pageHeight - 20) {
        doc.addPage()
        yPos = margin
      }

      if (person.milestone) {
        doc.setFont('helvetica', 'bold')
      } else {
        doc.setFont('helvetica', 'normal')
      }

      const dateStr = formatDayMonth(person.fodselsdato)
      const name = person.navn.length > 25 ? person.navn.substring(0, 22) + '...' : person.navn
      const avdeling = (person.avdeling || 'Ukjent').length > 30 ? 
        (person.avdeling || 'Ukjent').substring(0, 27) + '...' : 
        (person.avdeling || 'Ukjent')
      const age = `${person.alderDetteAret} år`

      doc.text(dateStr, col1, yPos)
      doc.text(name, col2, yPos)
      doc.text(avdeling, col3, yPos)
      doc.text(age, col4, yPos)
      yPos += 6
    }

    doc.save('bursdagskalender.pdf')
  }

  function exportBirthdayICS(items: BirthdayRow[]): void {
    const today = new Date()
    const year = today.getFullYear()
    const dtstamp = toICSDate(today)

    const events = items
      .map((employee, index) => {
        const date = new Date(
          year,
          employee.fodselsdato.getMonth(),
          employee.fodselsdato.getDate(),
        )
        const title = employee.milestone
          ? `${employee.navn} fyller ${employee.alderDetteAret} (milepæl)`
          : `${employee.navn} har bursdag`

        return [
          'BEGIN:VEVENT',
          `UID:${sanitizeICS(`${employee.navn}-${index}@ansatte`)}`,
          `DTSTAMP:${dtstamp}T080000Z`,
          `DTSTART;VALUE=DATE:${toICSDate(date)}`,
          'RRULE:FREQ=YEARLY',
          `SUMMARY:${sanitizeICS(title)}`,
          `DESCRIPTION:${sanitizeICS(`Avdeling: ${employee.avdeling || 'Ukjent'}`)}`,
          'END:VEVENT',
        ].join('\r\n')
      })
      .join('\r\n')

    const ics = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Ansatteverktøy//NO',
      'CALSCALE:GREGORIAN',
      events,
      'END:VCALENDAR',
      '',
    ].join('\r\n')

    downloadBlob(
      new Blob([ics], { type: 'text/calendar;charset=utf-8;' }),
      'ansatte-bursdager.ics',
    )
  }

  function exportRundeDagerPDF(): void {
    const doc = new jsPDF()
    const pageWidth = doc.internal.pageSize.getWidth()
    const pageHeight = doc.internal.pageSize.getHeight()
    const margin = 15
    let yPos = margin

    doc.setFontSize(18)
    doc.text('Runde dager', margin, yPos)
    yPos += 10

    doc.setFontSize(10)
    doc.text(`Generert: ${formatNorwegianDate(new Date())}`, margin, yPos)
    yPos += 8

    doc.setFontSize(9)
    doc.setFont('helvetica', 'bold')
    const col1 = margin
    const col2 = margin + 30
    const col3 = margin + 85
    const col4 = margin + 150

    doc.text('Dato', col1, yPos)
    doc.text('Navn', col2, yPos)
    doc.text('Avdeling', col3, yPos)
    doc.text('Alder', col4, yPos)
    yPos += 2
    doc.line(margin, yPos, pageWidth - margin, yPos)
    yPos += 5

    doc.setFont('helvetica', 'normal')
    if (milestones.length === 0) {
      doc.text('Ingen funnet.', margin, yPos)
    } else {
      for (const person of milestones) {
        if (yPos > pageHeight - 20) {
          doc.addPage()
          yPos = margin
        }

        const dateStr = formatDayMonth(person.fodselsdato)
        const name = person.navn.length > 25 ? `${person.navn.substring(0, 22)}...` : person.navn
        const avdeling = (person.avdeling || 'Ukjent').length > 30
          ? `${(person.avdeling || 'Ukjent').substring(0, 27)}...`
          : (person.avdeling || 'Ukjent')
        const age = `${person.alderDetteAret} år`

        doc.text(dateStr, col1, yPos)
        doc.text(name, col2, yPos)
        doc.text(avdeling, col3, yPos)
        doc.text(age, col4, yPos)
        yPos += 6
      }
    }

    doc.save('runde-dager.pdf')
  }

  function exportSeniorreduksjonPDF(): void {
    const doc = new jsPDF()
    const pageWidth = doc.internal.pageSize.getWidth()
    const pageHeight = doc.internal.pageSize.getHeight()
    const margin = 15
    let yPos = margin

    const drawSection = (title: string, rows: BirthdayRow[]): void => {
      if (yPos > pageHeight - 30) {
        doc.addPage()
        yPos = margin
      }

      doc.setFontSize(12)
      doc.setFont('helvetica', 'bold')
      doc.text(title, margin, yPos)
      yPos += 6

      doc.setFontSize(9)
      doc.text('Navn', margin, yPos)
      doc.text('Avdeling', margin + 85, yPos)
      doc.text('Alder', margin + 155, yPos)
      yPos += 2
      doc.line(margin, yPos, pageWidth - margin, yPos)
      yPos += 5

      doc.setFont('helvetica', 'normal')
      if (rows.length === 0) {
        doc.text('Ingen funnet.', margin, yPos)
        yPos += 8
        return
      }

      for (const person of rows) {
        if (yPos > pageHeight - 20) {
          doc.addPage()
          yPos = margin
        }

        const name = person.navn.length > 36 ? `${person.navn.substring(0, 33)}...` : person.navn
        const avdeling = (person.avdeling || 'Ukjent').length > 30
          ? `${(person.avdeling || 'Ukjent').substring(0, 27)}...`
          : (person.avdeling || 'Ukjent')
        const age = `${person.alderDetteAret} år`

        doc.text(name, margin, yPos)
        doc.text(avdeling, margin + 85, yPos)
        doc.text(age, margin + 155, yPos)
        yPos += 6
      }

      yPos += 4
    }

    doc.setFontSize(18)
    doc.text('Seniorreduksjon', margin, yPos)
    yPos += 10

    doc.setFontSize(10)
    doc.setFont('helvetica', 'normal')
    doc.text(`Generert: ${formatNorwegianDate(new Date())}`, margin, yPos)
    yPos += 10

    drawSection('57års-reduksjon (6%)', reduction57)
    drawSection('60års-reduksjon (12,5%)', reduction60)

    doc.save('seniorreduksjon.pdf')
  }

  function openBirthdayExport(format: BirthdayExportFormat): void {
    if (birthdays.length === 0) {
      return
    }

    setBirthdayExportFormat(format)
    setSelectedBirthdayAvdelinger(new Set(birthdayByAvdeling.map(([name]) => name)))
    setInactiveBirthdayExportIds(new Set())
  }

  function closeBirthdayExport(): void {
    setBirthdayExportFormat(null)
    setInactiveBirthdayExportIds(new Set())
  }

  function toggleBirthdayExportPerson(employeeId: string): void {
    setInactiveBirthdayExportIds((prev) => {
      const next = new Set(prev)
      if (next.has(employeeId)) {
        next.delete(employeeId)
      } else {
        next.add(employeeId)
      }
      return next
    })
  }

  function toggleAllBirthdayAvdelinger(): void {
    setSelectedBirthdayAvdelinger((prev) => {
      if (birthdayAvdelingNames.length === 0) {
        return new Set()
      }
      if (prev.size === birthdayAvdelingNames.length) {
        return new Set()
      }
      return new Set(birthdayAvdelingNames)
    })
    setInactiveBirthdayExportIds(new Set())
  }

  function toggleBirthdayAvdeling(avdeling: string): void {
    setSelectedBirthdayAvdelinger((prev) => {
      const next = new Set(prev)
      if (next.has(avdeling)) {
        next.delete(avdeling)
      } else {
        next.add(avdeling)
      }
      return next
    })
    setInactiveBirthdayExportIds(new Set())
  }

  function setBirthdayAvdelingActiveState(shouldBeActive: boolean): void {
    setInactiveBirthdayExportIds((prev) => {
      const next = new Set(prev)
      for (const person of birthdayAvdelingPeople) {
        if (shouldBeActive) {
          next.delete(person.id)
        } else {
          next.add(person.id)
        }
      }
      return next
    })
  }

  function confirmBirthdayExport(): void {
    if (!birthdayExportFormat || birthdayExportPeople.length === 0) {
      return
    }

    if (birthdayExportFormat === 'pdf') {
      exportBirthdayPDF(birthdayExportPeople)
    } else {
      exportBirthdayICS(birthdayExportPeople)
    }

    closeBirthdayExport()
  }

  async function exportAvdelingWord(): Promise<void> {
    const children: Array<Paragraph | Table> = [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 240 },
        children: [new TextRun({ text: 'Avdelingsoversikt', font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { after: 240 },
        children: [new TextRun({ text: `Generert: ${formatNorwegianDate(new Date())}`, font: 'Calibri' })],
      }),
    ]

    const rows: TableRow[] = []

    for (const [avdeling, list] of exportByAvdeling) {
      const specialEmployees = list.filter(emp => isSpecialStillingskode(emp.stillingskode))
      const regularEmployees = list.filter(emp => !isSpecialStillingskode(emp.stillingskode))
      
      const specialNames = specialEmployees.map(emp => emp.navn).join(', ')
      const headerText = specialNames ? `${avdeling}: ${specialNames}` : avdeling
      
      rows.push(
        new TableRow({
          children: [
            new TableCell({
              columnSpan: 2,
              shading: {
                fill: 'E8F5E9',
                color: 'auto',
              },
              children: [
                new Paragraph({
                  spacing: { after: 60 },
                  children: [
                    new TextRun({
                      text: headerText,
                      bold: true,
                      font: 'Calibri',
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      )

      const pairs = toTwoColumnPairs(regularEmployees)
      for (const [left, right] of pairs) {
        rows.push(
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: left?.navn ?? '', font: 'Calibri' })],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: right?.navn ?? '', font: 'Calibri' })],
                  }),
                ],
              }),
            ],
          }),
        )
      }
    }

    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows,
        borders: {
          top: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
          bottom: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
          left: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
          right: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'E5E5E5' },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'E5E5E5' },
        },
      }),
    )

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: {
                width: 11906,
                height: 16838,
              },
              margin: {
                top: 720,
                bottom: 720,
                left: 1440,
                right: 1440,
              },
            },
          },
          children,
        },
      ],
    })

    const blob = await Packer.toBlob(doc)
    downloadBlob(blob, 'avdelingsliste-a4.docx')
  }

  return (
    <main className="app-shell">
      <header className="hero">
        <p className="eyebrow">Ansatteverktøy</p>
        <h1>Bursdager og avdelingslister fra Excel</h1>
        <p className="lead">
          Last opp en Excel-fil. Data behandles lokalt i nettleseren. Ved import får du velge
          hvilke stillingskoder som skal inkluderes.
        </p>
      </header>

      <section className="panel upload-panel">
        <label
          className={`upload-box ${isDragging ? 'dragover' : ''}`}
          htmlFor="excel-upload"
          onDragOver={onDropZoneDragOver}
          onDragEnter={() => setIsDragging(true)}
          onDragLeave={onDropZoneDragLeave}
          onDrop={onDropZoneDrop}
        >
          <span>{isDragging ? 'Slipp filen her' : 'Klikk for å velge eller dra Excel-fil hit'}</span>
          <small>Støtter .xlsx og .xls</small>
        </label>
        <input id="excel-upload" type="file" accept=".xlsx,.xls" onChange={onFileChange} />
        {uploadHint ? <p className="status strong">{uploadHint}</p> : null}
        {fileName ? <p className="status">Fil: {fileName}</p> : null}
        {error ? <p className="error">Feil: {error}</p> : null}
      </section>

      <section className="panel actions-panel">
        <button disabled={birthdays.length === 0} onClick={() => openBirthdayExport('ics')}>
          Last ned bursdagskalender (.ics)
        </button>
        <button disabled={birthdays.length === 0} onClick={() => openBirthdayExport('pdf')}>
          Last ned bursdagskalender (.pdf)
        </button>
        <button disabled={milestones.length === 0} onClick={exportRundeDagerPDF}>
          Last ned runde dager (.pdf)
        </button>
        <button
          disabled={reduction57.length === 0 && reduction60.length === 0}
          onClick={exportSeniorreduksjonPDF}
        >
          Last ned seniorreduksjon (.pdf)
        </button>
      </section>

      {pendingImportRows.length > 0 ? (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="stillingskode-popup-title">
          <section className="modal-panel">
            <h2 id="stillingskode-popup-title">Velg stillingskoder som skal importeres</h2>
            <p className="muted">
              Hele stillingskode-feltet vises. Standardkoder er forhåndsvalgt øverst.
            </p>

            {allStillingskodeOptions.length === 0 ? (
              <p className="error">Fant ingen stillingskoder i filen.</p>
            ) : (
              <>
                <div className="stillingskode-section">
                  <h3>Forhåndsvalgte standardkoder</h3>
                  {defaultStillingskodeOptions.length === 0 ? (
                    <p className="muted">Ingen standardkoder ble funnet i filen.</p>
                  ) : (
                    <>
                      <div className="stillingskode-list-header" aria-hidden="true">
                        <span>Stillingskode</span>
                        <span>Arbeidstittel</span>
                      </div>
                      <ul className="stillingskode-list">
                        {defaultStillingskodeOptions.map((option) => {
                          const checked = selectedImportStillingskoder.has(option.id)
                          return (
                            <li key={`default-${option.id}`}>
                              <label>
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => toggleImportStillingskode(option.id)}
                                />
                                <span className="stillingskode-option-row">
                                  <span className="stillingskode-option-code">{option.stillingskode}</span>
                                  <span className="stillingskode-option-title">{option.arbeidstittel}</span>
                                </span>
                              </label>
                            </li>
                          )
                        })}
                      </ul>
                    </>
                  )}
                </div>

                <hr className="stillingskode-divider" />

                <div className="stillingskode-section">
                  <h3>Andre stillingskoder</h3>
                  {otherStillingskodeOptions.length === 0 ? (
                    <p className="muted">Ingen andre stillingskoder funnet.</p>
                  ) : (
                    <>
                      <div className="stillingskode-list-header" aria-hidden="true">
                        <span>Stillingskode</span>
                        <span>Arbeidstittel</span>
                      </div>
                      <ul className="stillingskode-list">
                        {otherStillingskodeOptions.map((option) => {
                          const checked = selectedImportStillingskoder.has(option.id)
                          return (
                            <li key={`other-${option.id}`}>
                              <label>
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => toggleImportStillingskode(option.id)}
                                />
                                <span className="stillingskode-option-row">
                                  <span className="stillingskode-option-code">{option.stillingskode}</span>
                                  <span className="stillingskode-option-title">{option.arbeidstittel}</span>
                                </span>
                              </label>
                            </li>
                          )
                        })}
                      </ul>
                    </>
                  )}
                </div>
              </>
            )}

            <div className="birthday-export-footer">
              <p className="muted">Valgt ansatte: {importPreviewCount}</p>
              <div className="birthday-export-buttons">
                <button
                  disabled={importPreviewCount === 0 || allStillingskodeOptions.length === 0}
                  onClick={confirmImportStillingskoder}
                >
                  Importer valgt
                </button>
                <button onClick={cancelImportStillingskoder}>Avbryt</button>
              </div>
            </div>
          </section>
        </div>
      ) : null}

      {birthdayExportFormat ? (
        <section className="panel birthday-export-panel">
          <h2>Velg avdeling og ansatte før nedlasting</h2>
          <div className="birthday-export-controls">
            <div className="birthday-avdeling-checklist">
              <p className="muted">1. Avdelinger</p>
              <div className="birthday-toggle-group" role="group" aria-label="Velg avdelinger">
                <button
                  type="button"
                  className={`toggle-pill ${isAllBirthdayAvdelingerSelected ? 'is-on' : ''}`}
                  aria-pressed={isAllBirthdayAvdelingerSelected}
                  onClick={toggleAllBirthdayAvdelinger}
                >
                  Alle
                </button>
                {birthdayByAvdeling.map(([avdeling]) => {
                  const checked = selectedBirthdayAvdelinger.has(avdeling)
                  return (
                    <button
                      key={avdeling}
                      type="button"
                      className={`toggle-pill ${checked ? 'is-on' : ''}`}
                      aria-pressed={checked}
                      onClick={() => toggleBirthdayAvdeling(avdeling)}
                    >
                      {avdeling}
                    </button>
                  )
                })}
              </div>
            </div>
          </div>

          {birthdayAvdelingPeople.length === 0 ? (
            <p className="muted">Ingen ansatte i valgt avdeling.</p>
          ) : (
            <div className="birthday-person-toggle-wrapper">
              <p className="muted">2. Ansatte i valgte avdelinger</p>
              <ul className="birthday-export-list">
                {birthdayAvdelingPeople.map((person) => {
                  const checked = !inactiveBirthdayExportIds.has(person.id)
                  return (
                    <li key={`${person.id}-birthday-export`}>
                      <label>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleBirthdayExportPerson(person.id)}
                        />
                        <span>{person.navn}</span>
                      </label>
                    </li>
                  )
                })}
              </ul>
              <div className="birthday-export-buttons">
                <button onClick={() => setBirthdayAvdelingActiveState(true)}>Marker alle</button>
                <button onClick={() => setBirthdayAvdelingActiveState(false)}>Fjern alle</button>
              </div>
            </div>
          )}

          <div className="birthday-export-footer">
            <p className="muted">
              Valgt: {birthdayExportPeople.length} av {birthdayAvdelingPeople.length}
            </p>
            <div className="birthday-export-buttons">
              <button disabled={birthdayExportPeople.length === 0} onClick={confirmBirthdayExport}>
                Last ned valgt kalender ({birthdayExportFormat === 'pdf' ? '.pdf' : '.ics'})
              </button>
              <button onClick={closeBirthdayExport}>Avbryt</button>
            </div>
          </div>
        </section>
      ) : null}

      <section className="stats-grid">
        <article className="metric panel">
          <h2>Filtrerte ansatte</h2>
          <p>{employees.length}</p>
        </article>
        <article className="metric panel">
          <h2>Runde dager i år</h2>
          <p>{milestones.length}</p>
        </article>
      </section>

      <section className="reductions-grid">
        <article className="panel">
          <h2>57års-reduksjon (6%)</h2>
          {reduction57.length === 0 ? (
            <p className="muted">Ingen ansatte mellom 57 og 59 år.</p>
          ) : (
            <ul className="plain-list">
              {reduction57.map((person) => (
                <li key={`${person.navn}-${person.fodselsdato.toISOString()}-57`}>
                  <strong>{person.navn}</strong> - {person.avdeling || 'Ukjent avdeling'} -{' '}
                  {person.alderDetteAret} år
                </li>
              ))}
            </ul>
          )}
        </article>

        <article className="panel">
          <h2>60års-reduksjon (12,5%)</h2>
          {reduction60.length === 0 ? (
            <p className="muted">Ingen ansatte som fyller eller har fylt 60 år.</p>
          ) : (
            <ul className="plain-list">
              {reduction60.map((person) => (
                <li key={`${person.navn}-${person.fodselsdato.toISOString()}-60`}>
                  <strong>{person.navn}</strong> - {person.avdeling || 'Ukjent avdeling'} -{' '}
                  {person.alderDetteAret} år
                </li>
              ))}
            </ul>
          )}
        </article>
      </section>

      <section className="panel">
        <h2>Runde dager</h2>
        {milestones.length === 0 ? (
          <p className="muted">Ingen funnet enda.</p>
        ) : (
          <ul className="plain-list">
            {milestones.map((person) => (
              <li key={`${person.navn}-${person.fodselsdato.toISOString()}`}>
                <strong>{person.navn}</strong> - {person.avdeling || 'Ukjent avdeling'} -{' '}
                {formatDayMonth(person.fodselsdato)} ({person.alderDetteAret} år)
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="panel">
        <h2>Bursdagskalender</h2>
        {birthdays.length === 0 ? (
          <p className="muted">Last opp fil for å se kalender.</p>
        ) : (
          <ul className="plain-list calendar-list">
            {birthdays.map((person) => (
              <li
                key={`${person.navn}-${person.fodselsdato.toISOString()}`}
                className={person.milestone ? 'highlight' : ''}
              >
                <span>{formatDayMonth(person.fodselsdato)}</span>
                <span>{person.navn}</span>
                <span>{person.avdeling || 'Ukjent avdeling'}</span>
                <span>{person.alderDetteAret} år</span>
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="panel avdelingsliste-panel">
        <h2>Avdelingsliste</h2>
        <p className="muted">Velg hvem som skal være med i avdelingslisten før eksport.</p>
        {employees.length > 0 ? (
          <div className="avdelingsliste-controls">
            <button onClick={() => setAllActive(true)}>Aktiver alle</button>
            <button onClick={() => setAllActive(false)}>Deaktiver alle</button>
            <button disabled={exportByAvdeling.length === 0} onClick={() => void exportAvdelingWord()}>
              Last ned avdelingsliste A4 (.docx)
            </button>
          </div>
        ) : null}

        {byAvdeling.length === 0 ? (
          <p className="muted">Ingen avdelinger enda.</p>
        ) : (
          <>
            <div className="avdelingsliste-picker">
              {byAvdeling.map(([avdeling, list]) => (
                <div className="avdeling-card" key={avdeling}>
                  <div className="avdeling-card-header">
                    <h3>{avdeling}</h3>
                    <div>
                      <button onClick={() => setAvdelingActiveState(avdeling, true)}>Alle på</button>
                      <button onClick={() => setAvdelingActiveState(avdeling, false)}>Alle av</button>
                    </div>
                  </div>
                  <ul>
                    {list.map((employee) => {
                      const checked = !inactiveAvdelingslisteIds.has(employee.id)
                      return (
                        <li key={employee.id}>
                          <label>
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={() => toggleAvdelingslistePerson(employee.id)}
                            />
                            <span>{employee.navn}</span>
                          </label>
                        </li>
                      )
                    })}
                  </ul>
                </div>
              ))}
            </div>

            <div className="avdeling">
              <table className="avdeling-table">
                <tbody>
                  {exportByAvdeling.flatMap(([avdeling, list]) => {
                    const specialEmployees = list.filter(emp => isSpecialStillingskode(emp.stillingskode))
                    const regularEmployees = list.filter(emp => !isSpecialStillingskode(emp.stillingskode))
                    
                    const specialNames = specialEmployees.map(emp => emp.navn).join(', ')
                    const headerText = specialNames ? `${avdeling}: ${specialNames}` : avdeling
                    
                    const pairs = toTwoColumnPairs(regularEmployees)
                    return [
                      <tr key={`${avdeling}-header`}>
                        <th colSpan={2}>{headerText}</th>
                      </tr>,
                      ...pairs.map(([left, right], index) => (
                        <tr key={`${avdeling}-${index}`}>
                          <td>{left?.navn ?? ''}</td>
                          <td>{right?.navn ?? ''}</td>
                        </tr>
                      )),
                    ]
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}
      </section>

      <footer className="panel privacy">
        <h2>Personvern</h2>
        <p>All behandling skjer lokalt i nettleseren. Data lagres ikke på server.</p>
      </footer>
    </main>
  )
}

export default App
