import './style.css'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'

const app = document.querySelector<HTMLDivElement>('#app')
const SESSION_KEY = 'labsoft-session'
const ISSUE_DRAFT_KEY = 'labsoft-issue-draft'
const DRAWN_DRAFT_KEY = 'labsoft-drawn-draft'
const SOFT_DELETE_TIMEOUT_MS = 8000

if (!app) {
  throw new Error('App root element not found')
}

type UserRole = 'admin' | 'staff'
type ModuleKey = 'issue-entry' | 'issue-records' | 'drawn-entry' | 'drawn-records' | 'admin-panel'
type RecordStatus = 'Pending' | 'In Progress' | 'Reported'

type LoginResponse = {
  token: string
  user: {
    email: string
    role: UserRole
  }
}

type MeResponse = {
  user: {
    email: string
    role: UserRole
  }
}

type Session = {
  token: string
  email: string
  role: UserRole
}

type IssueRecord = {
  id?: string
  createdAt?: string
  status: RecordStatus
  srNo: string
  codeNo: string
  sampleDescription: string
  parameterToBeTested: string
  issuedOn: string
  issuedBy: string
  issuedTo: string
  reportDueOn: string
  receivedBy: string
  reportedOn: string
  reportedByRemarks: string
}

type DrawnRecord = {
  id?: string
  createdAt?: string
  status: RecordStatus
  srNo: string
  sampleDescription: string
  sampleDrawnOn: string
  sampleDrawnBy: string
  customerNameAddress: string
  parameterToBeTested: string
  reportDueOn: string
  sampleReceivedBy: string
}

type RegistersResponse = {
  issueRecords: IssueRecord[]
  drawnRecords: DrawnRecord[]
}

type TestMasterTest = {
  id: string
  testName: string
  description: string
  displayOrder?: number
}

type TestMasterParameter = {
  id: string
  testId: string
  parameterName: string
  displayOrder?: number
}

type TestMasterResponse = {
  tests: TestMasterTest[]
  parameters: TestMasterParameter[]
}

const FALLBACK_TESTS: TestMasterTest[] = [
  { id: 't1', testName: 'Ambient Air Quality Monitoring & Analysis (Extended)', description: 'Ambient Air Quality Monitoring & Analysis', displayOrder: 1 },
  {
    id: 't2',
    testName: 'Ambient Air Quality Monitoring & Analysis (Basic)',
    description: 'Ambient Air Quality Monitoring & Analysis (Basic)',
    displayOrder: 2
  },
  { id: 't3', testName: 'Indoor Air', description: 'Indoor Air', displayOrder: 3 },
  { id: 't4', testName: 'Ambient Noise', description: 'Ambient Noise', displayOrder: 4 },
  { id: 't5', testName: 'Indoor Noise', description: 'Indoor Noise', displayOrder: 5 },
  { id: 't6', testName: 'DG Stack Emission', description: 'DG Stack Emission', displayOrder: 6 },
  { id: 't7', testName: 'DG Noise', description: 'DG Noise', displayOrder: 7 },
  { id: 't8', testName: 'ETP Inlet', description: 'ETP Inlet', displayOrder: 8 },
  { id: 't8b', testName: 'ETP Outlet', description: 'ETP Outlet', displayOrder: 9 },
  { id: 't8c', testName: 'STP Inlet', description: 'STP Inlet', displayOrder: 10 },
  { id: 't8d', testName: 'STP Outlet', description: 'STP Outlet', displayOrder: 11 },
  { id: 't8e', testName: 'Waste Water', description: 'Waste Water', displayOrder: 12 },
  { id: 't9', testName: 'Drinking Water Testing', description: 'Drinking Water Testing', displayOrder: 13 },
  { id: 't10', testName: 'Ground Water Quality', description: 'Ground Water Quality', displayOrder: 14 },
  { id: 't11', testName: 'Surface Water Bodies', description: 'Surface Water Bodies', displayOrder: 15 },
  { id: 't12', testName: 'Soil Quality Test', description: 'Soil Quality Test', displayOrder: 16 }
]

const FALLBACK_PARAMETERS: TestMasterParameter[] = [
  {
    id: 'p1',
    testId: 't1',
    parameterName: 'PM10, PM2.5, SO2, NO2, CO, Ammonia, Arsenic, Benzene, Lead, Nickel, Benzo(a)pyrene',
    displayOrder: 1
  },
  { id: 'p2', testId: 't2', parameterName: 'PM10, PM2.5, SO2, NO2, CO', displayOrder: 1 },
  { id: 'p3', testId: 't3', parameterName: 'PM, SO2, NO2, CO', displayOrder: 1 },
  { id: 'p4', testId: 't4', parameterName: 'Leq', displayOrder: 1 },
  { id: 'p5', testId: 't5', parameterName: 'Leq', displayOrder: 1 },
  { id: 'p6', testId: 't6', parameterName: 'PM, SOx, NOx, CO', displayOrder: 1 },
  { id: 'p7', testId: 't7', parameterName: 'Leq', displayOrder: 1 },
  { id: 'p8', testId: 't8', parameterName: 'pH, COD, BOD, TSS, Oil & Grease', displayOrder: 1 },
  { id: 'p8b', testId: 't8b', parameterName: 'pH, COD, BOD, TSS, Oil & Grease', displayOrder: 1 },
  { id: 'p8c', testId: 't8c', parameterName: 'pH, COD, BOD, TSS, Oil & Grease', displayOrder: 1 },
  { id: 'p8d', testId: 't8d', parameterName: 'pH, COD, BOD, TSS, Oil & Grease', displayOrder: 1 },
  { id: 'p8e', testId: 't8e', parameterName: 'pH, COD, BOD, TSS, Oil & Grease', displayOrder: 1 },
  {
    id: 'p9',
    testId: 't9',
    parameterName: 'pH, Colour, Odour, Taste, Turbidity, TDS, Calcium (as Ca), Chloride (as Cl), Fluoride (as F), Iron (as Fe), Magnesium (as Mg), Total Hardness (as CaCO3), Sulphate',
    displayOrder: 1
  },
  {
    id: 'p10',
    testId: 't10',
    parameterName: 'pH, Value, Colour, Odour, Taste, Turbidity, TDS, Total Hardness (as CaCO3), Calcium (as Ca), Magnesium (as Mg), Chloride (as Cl), Iron (as Fe), Fluoride (as F), Free Residual Chlorine, Phenolic Compound, Anionic Surface Detergents (as MBAS), Sulphate (as SO4), Nitrate (as NO3), Alkalinity (as CaCO3), Copper (as Cu), Total Ammonia, Sulphide (as H2S), Zinc (as Zn), Manganese (as Mn), Boron (as B), Selenium (as Se), Cadmium (as Cd), Lead (as Pb), Total Chromium (as Cr), Nickel (as Ni), Arsenic (as As)',
    displayOrder: 1
  },
  {
    id: 'p11',
    testId: 't11',
    parameterName: 'pH, Temperature, Turbidity, Conductivity, Total Suspended Solid, Total Alkalinity, BOD, DO, Calcium, Magnesium, Chlorides, Iron, Fluorides, Total Dissolved Solids, Total Hardness, Sulphate (SO4), Phosphate, Sodium, Manganese, Total Chromium, Zinc, Potassium, Nitrates, Cadmium, Lead, Copper, COD, Arsenic',
    displayOrder: 1
  },
  {
    id: 'p12',
    testId: 't12',
    parameterName: 'Texture, Sand %, Clay %, Moisture %, Silt %, pH, Electrical Conductivity, Potassium, Sodium, Calcium, Magnesium, Sodium Absorption Ratio, Water Holding Capacity, Total Kjeldahl Nitrogen, Bulk Density, Available Phosphorus, Organic Matter, Porosity',
    displayOrder: 1
  }
]

type AdminUser = {
  id: string
  email: string
  role: UserRole
  isActive: boolean
  createdAt: string
}

type AdminAlert = {
  type: 'overdue' | 'due-soon'
  source: 'issue' | 'drawn'
  recordId: string
  srNo: string
  dueOn: string
  message: string
}

type AuditEntry = {
  id: string
  actor: string
  action: string
  target: string
  details: string
  createdAt: string
}

type RegisterHistoryEntry = {
  id: string
  action: string
  source: 'issue' | 'drawn'
  actor: string
  recordId: string
  srNo: string
  createdAt: string
}

type RouteView = 'login' | ModuleKey

type PendingDelete = {
  source: 'issue' | 'drawn'
  module: ModuleKey
  index: number
  timeoutId: ReturnType<typeof setTimeout>
  record: IssueRecord | DrawnRecord
}

const issueRecords: IssueRecord[] = []
const drawnRecords: DrawnRecord[] = []
let issueSearch = ''
let issueFromDate = ''
let issueToDate = ''
let issueIssuedByFilter = ''
let issueIssuedToFilter = ''
let issueParameterFilter = ''
let issueCategoryFilter = 'All'
let drawnSearch = ''
let drawnFromDate = ''
let drawnToDate = ''
let drawnByFilter = ''
let drawnCustomerFilter = ''
let drawnParameterFilter = ''
let drawnCategoryFilter = 'All'
let issueEditingId = ''
let drawnEditingId = ''
const adminUsers: AdminUser[] = []
const adminAlerts: AdminAlert[] = []
const adminAuditEntries: AuditEntry[] = []
const adminRegisterHistoryEntries: RegisterHistoryEntry[] = []
const adminBackups: string[] = []
const testMasterTests: TestMasterTest[] = []
const testMasterParameters: TestMasterParameter[] = []
let adminMessage = ''
let adminMessageState: '' | 'error' = ''
let activeSession: Session | null = null
let currentView: RouteView | null = null
let pendingDelete: PendingDelete | null = null
let lastKnownDayKey = ''
let lastKnownOverdueCount = 0

const moduleRoutes: ModuleKey[] = ['issue-entry', 'issue-records', 'drawn-entry', 'drawn-records', 'admin-panel']

const isModuleKey = (value: string): value is ModuleKey => moduleRoutes.includes(value as ModuleKey)

const getHashView = (): RouteView | null => {
  const route = window.location.hash.replace(/^#/, '')

  if (!route) {
    return null
  }

  if (route === 'login') {
    return 'login'
  }

  return isModuleKey(route) ? route : null
}

const updateUrlView = (view: RouteView, mode: 'push' | 'replace'): void => {
  const nextHash = `#${view}`
  if (window.location.hash === nextHash) {
    return
  }

  const nextUrl = `${window.location.pathname}${window.location.search}${nextHash}`
  if (mode === 'replace') {
    window.history.replaceState(null, '', nextUrl)
    return
  }

  window.history.pushState(null, '', nextUrl)
}

const normalizeRole = (role: unknown, email: string): UserRole => {
  if (role === 'admin' || role === 'staff') {
    return role
  }

  return email === 'admin@labsoft.dev' ? 'admin' : 'staff'
}

const setIssueRecords = (records: IssueRecord[]): void => {
  issueRecords.splice(0, issueRecords.length, ...records)
}

const setDrawnRecords = (records: DrawnRecord[]): void => {
  drawnRecords.splice(0, drawnRecords.length, ...records)
}

const setAdminUsers = (users: AdminUser[]): void => {
  adminUsers.splice(0, adminUsers.length, ...users)
}

const setAdminAlerts = (alerts: AdminAlert[]): void => {
  adminAlerts.splice(0, adminAlerts.length, ...alerts)
}

const setAdminAuditEntries = (entries: AuditEntry[]): void => {
  adminAuditEntries.splice(0, adminAuditEntries.length, ...entries)
}

const setAdminRegisterHistoryEntries = (entries: RegisterHistoryEntry[]): void => {
  adminRegisterHistoryEntries.splice(0, adminRegisterHistoryEntries.length, ...entries)
}

const setAdminBackups = (backups: string[]): void => {
  adminBackups.splice(0, adminBackups.length, ...backups)
}

const setTestMaster = (tests: TestMasterTest[], parameters: TestMasterParameter[]): void => {
  const hasTests = Array.isArray(tests) && tests.length > 0
  const hasParameters = Array.isArray(parameters) && parameters.length > 0

  testMasterTests.splice(0, testMasterTests.length, ...(hasTests ? tests : FALLBACK_TESTS))
  testMasterParameters.splice(0, testMasterParameters.length, ...(hasParameters ? parameters : FALLBACK_PARAMETERS))
}

const normalizeMasterKey = (value: string): string => value.trim().toLowerCase()

const getParameterOptions = (description: string): string[] => {
  const match = testMasterTests.find((item) => normalizeMasterKey(item.description) === normalizeMasterKey(description))
  if (!match) {
    return []
  }

  const raw = testMasterParameters
    .filter((item) => item.testId === match.id)
    .sort((first, second) => (first.displayOrder ?? 0) - (second.displayOrder ?? 0))
    .map((item) => item.parameterName.trim())
    .filter(Boolean)

  const normalizedSeen = new Set<string>()
  const options: string[] = []

  raw.forEach((entry) => {
    const key = normalizeMasterKey(entry)
    if (normalizedSeen.has(key)) {
      return
    }

    normalizedSeen.add(key)
    options.push(entry)
  })

  return options
}

const renderSampleDescriptionOptions = (selectedDescription: string): string => {
  const normalizedSelected = normalizeMasterKey(selectedDescription)
  const sortedTests = testMasterTests
    .slice()
    .sort((first, second) => (first.displayOrder ?? 0) - (second.displayOrder ?? 0))
  const hasSelectedInMaster = sortedTests.some((item) => normalizeMasterKey(item.description) === normalizedSelected)

  const options = sortedTests
    .map((item) => {
      const selected = normalizeMasterKey(item.description) === normalizedSelected ? 'selected' : ''
      return `<option value="${escapeHtml(item.description)}" ${selected}>${escapeHtml(item.description)}</option>`
    })
    .join('')

  if (selectedDescription && !hasSelectedInMaster) {
    return `<option value="${escapeHtml(selectedDescription)}" selected>${escapeHtml(selectedDescription)}</option>${options}`
  }

  return options
}

const readNumericPart = (value: string): number | null => {
  const match = String(value).match(/(\d+)/)
  if (!match) {
    return null
  }

  const parsed = Number.parseInt(match[1], 10)
  return Number.isNaN(parsed) ? null : parsed
}

const getNextSerialNumber = (records: Array<{ srNo: string }>): string => {
  let maxSerial = 0

  records.forEach((record) => {
    const numeric = readNumericPart(record.srNo)
    if (numeric && numeric > maxSerial) {
      maxSerial = numeric
    }
  })

  return String(maxSerial + 1)
}

const setInputValue = (input: HTMLInputElement | null, value: string): void => {
  if (!input) {
    return
  }

  input.value = value
}

const MANUAL_PARAMETER_VALUE = '__manual__'

const hasParameterInMaster = (description: string, value: string): boolean => {
  if (!value) {
    return false
  }

  const normalizedValue = normalizeMasterKey(value)
  return getParameterOptions(description).some((item) => normalizeMasterKey(item) === normalizedValue)
}

const getInitialParameterValue = (description: string, fallback: string): string => {
  const options = getParameterOptions(description)
  if (fallback && options.some((item) => normalizeMasterKey(item) === normalizeMasterKey(fallback))) {
    return fallback
  }

  if (fallback && options.length === 0) {
    return fallback
  }

  return options[0] ?? ''
}

const renderParameterOptions = (description: string, selectedValue: string): string => {
  const options = getParameterOptions(description)
  const normalizedSelected = normalizeMasterKey(selectedValue)
  const hasSelected = options.some((item) => normalizeMasterKey(item) === normalizedSelected)

  const mapped = options
    .map((item) => {
      const selected = normalizeMasterKey(item) === normalizedSelected ? 'selected' : ''
      return `<option value="${escapeHtml(item)}" ${selected}>${escapeHtml(item)}</option>`
    })
    .join('')

  if (selectedValue && !hasSelected) {
    return `<option value="${escapeHtml(selectedValue)}" selected>${escapeHtml(selectedValue)}</option>${mapped}`
  }

  return mapped
}

const syncParameterSelectOptions = (
  descriptionSelect: HTMLSelectElement | null,
  parameterSelect: HTMLSelectElement | null,
  preserveExisting: boolean
): void => {
  if (!descriptionSelect || !parameterSelect) {
    return
  }

  const currentValue = preserveExisting ? parameterSelect.value : ''
  const manualInput = parameterSelect
    .closest('form')
    ?.querySelector<HTMLInputElement>('input[name="parameterManual"]')
  const manualValue = manualInput?.value.trim() ?? ''
  const keepManualSelection =
    preserveExisting &&
    (currentValue === MANUAL_PARAMETER_VALUE ||
      (currentValue && !hasParameterInMaster(descriptionSelect.value, currentValue) && Boolean(manualValue)))
  const nextValue = getInitialParameterValue(descriptionSelect.value, currentValue)
  const optionsMarkup = renderParameterOptions(descriptionSelect.value, nextValue)

  parameterSelect.innerHTML = `<option value="">Select parameter</option>${optionsMarkup}<option value="${MANUAL_PARAMETER_VALUE}">Manual Entry</option>`

  if (keepManualSelection) {
    parameterSelect.value = MANUAL_PARAMETER_VALUE
    return
  }

  if (nextValue) {
    parameterSelect.value = nextValue
  }
}

const getInitialIssueSrNo = (editing: IssueRecord | undefined): string => {
  if (editing?.srNo) {
    return editing.srNo
  }

  return getNextSerialNumber(issueRecords)
}

const getInitialIssueCodeNo = (
  editing: IssueRecord | undefined,
  draftCodeNo: string | undefined
): string => {
  if (editing?.codeNo) {
    return editing.codeNo
  }

  return draftCodeNo ?? ''
}

const getInitialDrawnSrNo = (editing: DrawnRecord | undefined): string => {
  if (editing?.srNo) {
    return editing.srNo
  }

  return getNextSerialNumber(drawnRecords)
}

const renderSampleDescriptionSelect = (selectedDescription: string, label: string): string => {
  return `
    <label class="field-group"><span>${label}</span>
      <select name="sampleDescription" required>
        <option value="">Select sample type</option>
        ${renderSampleDescriptionOptions(selectedDescription)}
      </select>
    </label>
  `
}

const renderAutoParameterField = (description: string, value: string, label: string): string => {
  const manualSelected = Boolean(value) && !hasParameterInMaster(description, value)
  const selectedValue = manualSelected ? '' : value
  return `
    <label class="field-group"><span>${label}</span>
      <select name="parameterToBeTested" required>
        <option value="">Select parameter</option>
        ${renderParameterOptions(description, selectedValue)}
        <option value="${MANUAL_PARAMETER_VALUE}" ${manualSelected ? 'selected' : ''}>Manual Entry</option>
      </select>
    </label>
    <label class="field-group manual-parameter ${manualSelected ? '' : 'hidden'}"><span>Manual Parameter</span>
      <input name="parameterManual" value="${escapeHtml(manualSelected ? value : '')}" placeholder="Enter full test parameter set" ${manualSelected ? 'required' : ''} />
    </label>
  `
}

const renderAutoSerialField = (value: string, readonly: boolean, label = 'Sr.No.'): string => {
  return `<label class="field-group"><span>${label}</span><input name="srNo" value="${escapeHtml(value)}" required ${readonly ? 'readonly' : ''} /></label>`
}

const renderAutoCodeField = (value: string): string => {
  return `<label class="field-group"><span>Code No.</span><input name="codeNo" value="${escapeHtml(value)}" required /></label>`
}

const renderSectionHint = (text: string): string => `<p class="draft-note">${text}</p>`

const getIssueDescriptionValue = (editing: IssueRecord | undefined, draft: Partial<IssueRecord>): string => {
  return editing?.sampleDescription ?? draft.sampleDescription ?? ''
}

const getDrawnDescriptionValue = (editing: DrawnRecord | undefined, draft: Partial<DrawnRecord>): string => {
  return editing?.sampleDescription ?? draft.sampleDescription ?? ''
}

const getIssueParameterValue = (editing: IssueRecord | undefined, draft: Partial<IssueRecord>, description: string): string => {
  return getInitialParameterValue(description, editing?.parameterToBeTested ?? draft.parameterToBeTested ?? '')
}

const getDrawnParameterValue = (editing: DrawnRecord | undefined, draft: Partial<DrawnRecord>, description: string): string => {
  return getInitialParameterValue(description, editing?.parameterToBeTested ?? draft.parameterToBeTested ?? '')
}

const shouldReadonlyAutoField = (isEditing: boolean): boolean => !isEditing

const syncAutoIssueFields = (form: HTMLFormElement, isEditing: boolean): void => {
  const srNoInput = form.querySelector<HTMLInputElement>('input[name="srNo"]')
  const descriptionSelect = form.querySelector<HTMLSelectElement>('select[name="sampleDescription"]')
  const parameterSelect = form.querySelector<HTMLSelectElement>('select[name="parameterToBeTested"]')

  if (!isEditing && !srNoInput?.value.trim()) {
    setInputValue(srNoInput, getNextSerialNumber(issueRecords))
  }

  syncParameterSelectOptions(descriptionSelect, parameterSelect, true)
}

const syncAutoDrawnFields = (form: HTMLFormElement, isEditing: boolean): void => {
  const srNoInput = form.querySelector<HTMLInputElement>('input[name="srNo"]')
  const descriptionSelect = form.querySelector<HTMLSelectElement>('select[name="sampleDescription"]')
  const parameterSelect = form.querySelector<HTMLSelectElement>('select[name="parameterToBeTested"]')

  if (!isEditing && !srNoInput?.value.trim()) {
    setInputValue(srNoInput, getNextSerialNumber(drawnRecords))
  }

  syncParameterSelectOptions(descriptionSelect, parameterSelect, true)
}

const bindIssueAutoEvents = (form: HTMLFormElement, isEditing: boolean): void => {
  void isEditing
  const descriptionSelect = form.querySelector<HTMLSelectElement>('select[name="sampleDescription"]')
  const parameterSelect = form.querySelector<HTMLSelectElement>('select[name="parameterToBeTested"]')

  descriptionSelect?.addEventListener('change', () => {
    syncParameterSelectOptions(descriptionSelect, parameterSelect, false)
    syncManualParameterUi(form)
  })

  parameterSelect?.addEventListener('change', () => {
    syncManualParameterUi(form)
  })
}

const bindDrawnAutoEvents = (form: HTMLFormElement, isEditing: boolean): void => {
  void isEditing
  const descriptionSelect = form.querySelector<HTMLSelectElement>('select[name="sampleDescription"]')
  const parameterSelect = form.querySelector<HTMLSelectElement>('select[name="parameterToBeTested"]')

  descriptionSelect?.addEventListener('change', () => {
    syncParameterSelectOptions(descriptionSelect, parameterSelect, false)
    syncManualParameterUi(form)
  })

  parameterSelect?.addEventListener('change', () => {
    syncManualParameterUi(form)
  })
}

const syncManualParameterUi = (form: HTMLFormElement): void => {
  const parameterSelect = form.querySelector<HTMLSelectElement>('select[name="parameterToBeTested"]')
  const manualWrapper = form.querySelector<HTMLLabelElement>('.manual-parameter')
  const manualInput = form.querySelector<HTMLInputElement>('input[name="parameterManual"]')

  if (!parameterSelect || !manualWrapper || !manualInput) {
    return
  }

  const manualMode = parameterSelect.value === MANUAL_PARAMETER_VALUE
  manualWrapper.classList.toggle('hidden', !manualMode)

  if (manualMode) {
    manualInput.required = true
    return
  }

  manualInput.required = false
  manualInput.value = ''
}

const getResolvedParameterValue = (formData: FormData, getter: (formData: FormData, name: string) => string): string => {
  const selectedParameter = getter(formData, 'parameterToBeTested')
  if (selectedParameter === MANUAL_PARAMETER_VALUE) {
    return getter(formData, 'parameterManual')
  }

  return selectedParameter
}

const syncIssueDraftFromAutoFields = (form: HTMLFormElement): void => {
  const srNoInput = form.querySelector<HTMLInputElement>('input[name="srNo"]')

  if (!srNoInput) {
    return
  }

  if (!srNoInput.value.trim()) {
    srNoInput.value = getNextSerialNumber(issueRecords)
  }
}

const syncDrawnDraftFromAutoFields = (form: HTMLFormElement): void => {
  const srNoInput = form.querySelector<HTMLInputElement>('input[name="srNo"]')
  if (!srNoInput) {
    return
  }

  if (!srNoInput.value.trim()) {
    srNoInput.value = getNextSerialNumber(drawnRecords)
  }
}

const getIssueFieldValue = (formData: FormData, name: string): string => String(formData.get(name) ?? '').trim()

const getDrawnFieldValue = (formData: FormData, name: string): string => String(formData.get(name) ?? '').trim()

const ensureIssueAutoDefaults = (payload: IssueRecord): IssueRecord => {
  const nextPayload = { ...payload }

  if (!nextPayload.srNo) {
    nextPayload.srNo = getNextSerialNumber(issueRecords)
  }

  if (nextPayload.sampleDescription && !nextPayload.parameterToBeTested) {
    nextPayload.parameterToBeTested = getInitialParameterValue(nextPayload.sampleDescription, '')
  }

  return nextPayload
}

const ensureDrawnAutoDefaults = (payload: DrawnRecord): DrawnRecord => {
  const nextPayload = { ...payload }

  if (!nextPayload.srNo) {
    nextPayload.srNo = getNextSerialNumber(drawnRecords)
  }

  if (nextPayload.sampleDescription && !nextPayload.parameterToBeTested) {
    nextPayload.parameterToBeTested = getInitialParameterValue(nextPayload.sampleDescription, '')
  }

  return nextPayload
}

const getIssueDraftOrEditing = (editing: IssueRecord | undefined): Partial<IssueRecord> => (editing ? {} : readDraft<IssueRecord>(ISSUE_DRAFT_KEY))

const getDrawnDraftOrEditing = (editing: DrawnRecord | undefined): Partial<DrawnRecord> => (editing ? {} : readDraft<DrawnRecord>(DRAWN_DRAFT_KEY))

const getIssueFormInitialValues = (editing: IssueRecord | undefined): {
  issueDraft: Partial<IssueRecord>
  srNoValue: string
  codeNoValue: string
  descriptionValue: string
  parameterValue: string
} => {
  const issueDraft = getIssueDraftOrEditing(editing)
  const srNoValue = getInitialIssueSrNo(editing)
  const codeNoValue = getInitialIssueCodeNo(editing, issueDraft.codeNo)
  const descriptionValue = getIssueDescriptionValue(editing, issueDraft)
  const parameterValue = getIssueParameterValue(editing, issueDraft, descriptionValue)

  return { issueDraft, srNoValue, codeNoValue, descriptionValue, parameterValue }
}

const getDrawnFormInitialValues = (editing: DrawnRecord | undefined): {
  drawnDraft: Partial<DrawnRecord>
  srNoValue: string
  descriptionValue: string
  parameterValue: string
} => {
  const drawnDraft = getDrawnDraftOrEditing(editing)
  const srNoValue = getInitialDrawnSrNo(editing)
  const descriptionValue = getDrawnDescriptionValue(editing, drawnDraft)
  const parameterValue = getDrawnParameterValue(editing, drawnDraft, descriptionValue)

  return { drawnDraft, srNoValue, descriptionValue, parameterValue }
}

const renderIssueAutoFields = (
  srNoValue: string,
  codeNoValue: string,
  descriptionValue: string,
  parameterValue: string,
  isEditing: boolean
): string => {
  return [
    renderAutoSerialField(srNoValue, shouldReadonlyAutoField(isEditing)),
    renderAutoCodeField(codeNoValue),
    renderSampleDescriptionSelect(descriptionValue, 'Sample Description'),
    renderAutoParameterField(descriptionValue, parameterValue, 'Parameter to be tested')
  ].join('')
}

const renderDrawnAutoFields = (
  srNoValue: string,
  descriptionValue: string,
  parameterValue: string,
  isEditing: boolean
): string => {
  return [
    renderAutoSerialField(srNoValue, shouldReadonlyAutoField(isEditing)),
    renderSampleDescriptionSelect(descriptionValue, 'Sample Description'),
    renderAutoParameterField(descriptionValue, parameterValue, 'Parameter to be Tested')
  ].join('')
}

const renderDraftHint = (isEditing: boolean): string => (isEditing ? '' : renderSectionHint('Draft auto-save is on.'))

const initializeIssueAutoUi = (form: HTMLFormElement, isEditing: boolean): void => {
  syncAutoIssueFields(form, isEditing)
  syncManualParameterUi(form)
  bindIssueAutoEvents(form, isEditing)
}

const initializeDrawnAutoUi = (form: HTMLFormElement, isEditing: boolean): void => {
  syncAutoDrawnFields(form, isEditing)
  syncManualParameterUi(form)
  bindDrawnAutoEvents(form, isEditing)
}

const readIssuePayloadFromForm = (formData: FormData): IssueRecord => {
  return ensureIssueAutoDefaults({
    srNo: getIssueFieldValue(formData, 'srNo'),
    codeNo: getIssueFieldValue(formData, 'codeNo'),
    status: normalizeRecordStatus(formData.get('status')),
    sampleDescription: getIssueFieldValue(formData, 'sampleDescription'),
    parameterToBeTested: getResolvedParameterValue(formData, getIssueFieldValue),
    issuedOn: getIssueFieldValue(formData, 'issuedOn'),
    issuedBy: getIssueFieldValue(formData, 'issuedBy'),
    issuedTo: getIssueFieldValue(formData, 'issuedTo'),
    reportDueOn: getIssueFieldValue(formData, 'reportDueOn'),
    receivedBy: getIssueFieldValue(formData, 'receivedBy'),
    reportedOn: getIssueFieldValue(formData, 'reportedOn'),
    reportedByRemarks: getIssueFieldValue(formData, 'reportedByRemarks')
  })
}

const readDrawnPayloadFromForm = (formData: FormData): DrawnRecord => {
  return ensureDrawnAutoDefaults({
    srNo: getDrawnFieldValue(formData, 'srNo'),
    status: normalizeRecordStatus(formData.get('status')),
    sampleDescription: getDrawnFieldValue(formData, 'sampleDescription'),
    sampleDrawnOn: getDrawnFieldValue(formData, 'sampleDrawnOn'),
    sampleDrawnBy: getDrawnFieldValue(formData, 'sampleDrawnBy'),
    customerNameAddress: getDrawnFieldValue(formData, 'customerNameAddress'),
    parameterToBeTested: getResolvedParameterValue(formData, getDrawnFieldValue),
    reportDueOn: getDrawnFieldValue(formData, 'reportDueOn'),
    sampleReceivedBy: getDrawnFieldValue(formData, 'sampleReceivedBy')
  })
}

const syncIssueDraftPayload = (form: HTMLFormElement): Record<string, string> => {
  syncIssueDraftFromAutoFields(form)
  const draftData = new FormData(form)
  return {
    srNo: getIssueFieldValue(draftData, 'srNo'),
    codeNo: getIssueFieldValue(draftData, 'codeNo'),
    status: getIssueFieldValue(draftData, 'status'),
    sampleDescription: getIssueFieldValue(draftData, 'sampleDescription'),
    parameterToBeTested: getResolvedParameterValue(draftData, getIssueFieldValue),
    issuedOn: getIssueFieldValue(draftData, 'issuedOn'),
    issuedBy: getIssueFieldValue(draftData, 'issuedBy'),
    issuedTo: getIssueFieldValue(draftData, 'issuedTo'),
    reportDueOn: getIssueFieldValue(draftData, 'reportDueOn'),
    receivedBy: getIssueFieldValue(draftData, 'receivedBy'),
    reportedOn: getIssueFieldValue(draftData, 'reportedOn'),
    reportedByRemarks: getIssueFieldValue(draftData, 'reportedByRemarks')
  }
}

const syncDrawnDraftPayload = (form: HTMLFormElement): Record<string, string> => {
  syncDrawnDraftFromAutoFields(form)
  const draftData = new FormData(form)
  return {
    srNo: getDrawnFieldValue(draftData, 'srNo'),
    status: getDrawnFieldValue(draftData, 'status'),
    sampleDescription: getDrawnFieldValue(draftData, 'sampleDescription'),
    sampleDrawnOn: getDrawnFieldValue(draftData, 'sampleDrawnOn'),
    sampleDrawnBy: getDrawnFieldValue(draftData, 'sampleDrawnBy'),
    customerNameAddress: getDrawnFieldValue(draftData, 'customerNameAddress'),
    parameterToBeTested: getResolvedParameterValue(draftData, getDrawnFieldValue),
    reportDueOn: getDrawnFieldValue(draftData, 'reportDueOn'),
    sampleReceivedBy: getDrawnFieldValue(draftData, 'sampleReceivedBy')
  }
}

const renderIssueDefaultDescriptionMessage = (descriptionValue: string): string =>
  descriptionValue ? '' : '<p class="draft-note">Select Sample Description to auto-load parameters.</p>'

const renderDrawnDefaultDescriptionMessage = (descriptionValue: string): string =>
  descriptionValue ? '' : '<p class="draft-note">Select Sample Description to auto-load parameters.</p>'

const toDateValue = (value: string): number => {
  const parsed = Date.parse(value)
  return Number.isNaN(parsed) ? 0 : parsed
}

const escapeHtml = (value: string): string =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')

const escapeCsv = (value: string): string => `"${String(value).replace(/"/g, '""')}"`

const readDraft = <T extends Record<string, unknown>>(key: string): Partial<T> => {
  const raw = localStorage.getItem(key)
  if (!raw) {
    return {}
  }

  try {
    const parsed = JSON.parse(raw) as Partial<T>
    return typeof parsed === 'object' && parsed ? parsed : {}
  } catch {
    return {}
  }
}

const saveDraft = (key: string, payload: Record<string, string>): void => {
  localStorage.setItem(key, JSON.stringify(payload))
}

const clearDraft = (key: string): void => {
  localStorage.removeItem(key)
}

const toLocalDateValue = (value: string): number => {
  if (!value) {
    return 0
  }

  const date = new Date(`${value}T00:00:00`)
  return Number.isNaN(date.getTime()) ? 0 : date.getTime()
}

const toLocalDateKey = (value: string): string => {
  if (!value) {
    return ''
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value
  }

  const date = new Date(value)
  if (Number.isNaN(date.getTime())) {
    return ''
  }

  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

const getTodayLocalDateKey = (): string => {
  const date = new Date()
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

const normalizeRecordStatus = (value: unknown, fallback: RecordStatus = 'Pending'): RecordStatus => {
  if (value === 'Pending' || value === 'In Progress' || value === 'Reported') {
    return value
  }

  return fallback
}

const getIssueStatus = (record: IssueRecord): RecordStatus => {
  return normalizeRecordStatus(record.status, record.reportedOn.trim() ? 'Reported' : 'Pending')
}

const getDrawnStatus = (record: DrawnRecord): RecordStatus => {
  return normalizeRecordStatus(record.status, 'Pending')
}

const isOverdue = (dueOn: string, status: RecordStatus, completedOn = ''): boolean => {
  if (!dueOn || status === 'Reported' || completedOn.trim()) {
    return false
  }

  const dueTime = toLocalDateValue(dueOn)
  if (!dueTime) {
    return false
  }

  const today = new Date()
  today.setHours(0, 0, 0, 0)
  return dueTime < today.getTime()
}

const getActivityStats = (): { totalEntries: number; overdueEntries: number; todayEntries: number } => {
  const today = getTodayLocalDateKey()
  const issueToday = issueRecords.filter((record) => toLocalDateKey(record.createdAt ?? record.issuedOn) === today).length
  const drawnToday = drawnRecords.filter((record) => toLocalDateKey(record.createdAt ?? record.sampleDrawnOn) === today).length
  const issueOverdue = issueRecords.filter((record) => isOverdue(record.reportDueOn, getIssueStatus(record), record.reportedOn)).length
  const drawnOverdue = drawnRecords.filter((record) => isOverdue(record.reportDueOn, getDrawnStatus(record))).length

  return {
    totalEntries: issueRecords.length + drawnRecords.length,
    overdueEntries: issueOverdue + drawnOverdue,
    todayEntries: issueToday + drawnToday
  }
}

const hasIssueDuplicate = (record: IssueRecord, excludeId = ''): string | null => {
  const srNo = record.srNo.trim().toLowerCase()
  const codeNo = record.codeNo.trim().toLowerCase()

  const duplicate = issueRecords.find(
    (entry) =>
      entry.id !== excludeId &&
      (entry.srNo.trim().toLowerCase() === srNo || entry.codeNo.trim().toLowerCase() === codeNo)
  )

  if (!duplicate) {
    return null
  }

  if (duplicate.srNo.trim().toLowerCase() === srNo) {
    return `Duplicate Sr.No. found: ${record.srNo}`
  }

  return `Duplicate Code No. found: ${record.codeNo}`
}

const hasDrawnDuplicate = (record: DrawnRecord, excludeId = ''): string | null => {
  const srNo = record.srNo.trim().toLowerCase()

  const duplicate = drawnRecords.find((entry) => entry.id !== excludeId && entry.srNo.trim().toLowerCase() === srNo)
  if (!duplicate) {
    return null
  }

  return `Duplicate Sr.No. found: ${record.srNo}`
}

const pdfLogoPath = '/ultra-lab-logo.png'
const ISSUE_PDF_TABLE_WIDTH = 198
const DRAWN_PDF_TABLE_WIDTH = 184
const pdfLogoCache = new Map<string, { dataUrl: string; width: number; height: number }>()

const getPdfLogoAsset = async (alpha = 1): Promise<{ dataUrl: string; width: number; height: number }> => {
  const key = alpha.toFixed(2)
  const cached = pdfLogoCache.get(key)
  if (cached) {
    return cached
  }

  const image = new Image()
  image.src = `${pdfLogoPath}?t=${Date.now()}`

  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve()
    image.onerror = () => reject(new Error('Logo load failed for PDF.'))
  })

  const canvas = document.createElement('canvas')
  canvas.width = image.width
  canvas.height = image.height
  const context = canvas.getContext('2d')
  if (!context) {
    throw new Error('Logo canvas init failed for PDF.')
  }

  context.clearRect(0, 0, canvas.width, canvas.height)
  context.globalAlpha = Math.max(0, Math.min(alpha, 1))
  context.drawImage(image, 0, 0)
  const asset = {
    dataUrl: canvas.toDataURL('image/png'),
    width: image.width,
    height: image.height
  }

  pdfLogoCache.set(key, asset)
  return asset
}

const drawPdfHeaderLogo = async (pdf: jsPDF): Promise<void> => {
  const logo = await getPdfLogoAsset(1)
  const boxX = 12
  const boxY = 1
  const boxWidth = 52
  const boxHeight = 36
  const ratio = logo.width / logo.height

  let drawWidth = boxWidth
  let drawHeight = drawWidth / ratio
  if (drawHeight > boxHeight) {
    drawHeight = boxHeight
    drawWidth = drawHeight * ratio
  }

  const drawX = boxX + (boxWidth - drawWidth) / 2
  const drawY = boxY + (boxHeight - drawHeight) / 2
  pdf.addImage(logo.dataUrl, 'PNG', drawX, drawY, drawWidth, drawHeight)
}

const drawPdfBottomBackgroundLogo = async (pdf: jsPDF): Promise<void> => {
  const logo = await getPdfLogoAsset(0.2)
  const pageWidth = pdf.internal.pageSize.getWidth()
  const pageHeight = pdf.internal.pageSize.getHeight()
  const targetHeight = 62
  const targetWidth = targetHeight * (logo.width / logo.height)
  const drawX = pageWidth / 2 - targetWidth / 2
  const drawY = pageHeight - targetHeight - 10
  pdf.addImage(logo.dataUrl, 'PNG', drawX, drawY, targetWidth, targetHeight)
}

const downloadIssueRegisterPdf = async (records: IssueRecord[]): Promise<void> => {
  const pdf = new jsPDF({ orientation: 'landscape', format: 'a4' })
  const pageWidth = pdf.internal.pageSize.getWidth()
  const marginLeft = (pageWidth - ISSUE_PDF_TABLE_WIDTH) / 2

  await drawPdfBottomBackgroundLogo(pdf)
  await drawPdfHeaderLogo(pdf)

  pdf.setTextColor(20, 20, 20)
  pdf.setFont('times', 'bold')
  pdf.setFontSize(12)
  pdf.text('ULTRA TESTING & RESEARCH LABORATORY', 148.5, 13, { align: 'center' })
  pdf.setFontSize(11)
  pdf.text('SAMPLE ISSUE REGISTER', 148.5, 19, { align: 'center' })

  const headers = [
    'Sr.No.',
    'Code No.',
    'Sample Description',
    'Parameter to be tested',
    'Issued On',
    'Issued By',
    'Issued To',
    'Report Due On',
    'Received By',
    'Reported On',
    'ReportedBy/R\nremarks'
  ]

  const rows = records.map((record) => [
    record.srNo,
    record.codeNo,
    record.sampleDescription,
    record.parameterToBeTested,
    record.issuedOn,
    record.issuedBy,
    record.issuedTo,
    record.reportDueOn,
    record.receivedBy,
    record.reportedOn,
    record.reportedByRemarks
  ])

  autoTable(pdf, {
    head: [headers],
    body: rows,
    startY: 34,
    margin: { top: 34, right: marginLeft, bottom: 8, left: marginLeft },
    tableWidth: ISSUE_PDF_TABLE_WIDTH,
    theme: 'grid',
    styles: {
      font: 'times',
      fontSize: 7,
      cellPadding: 1.2,
      lineColor: [70, 70, 70],
      lineWidth: 0.1,
      textColor: [20, 20, 20],
      valign: 'middle'
    },
    headStyles: {
      font: 'times',
      fontStyle: 'bold',
      fillColor: [255, 255, 255],
      textColor: [20, 20, 20],
      lineColor: [70, 70, 70],
      lineWidth: 0.1
    },
    columnStyles: {
      0: { cellWidth: 12 },
      1: { cellWidth: 14 },
      2: { cellWidth: 24 },
      3: { cellWidth: 36 },
      4: { cellWidth: 14 },
      5: { cellWidth: 13 },
      6: { cellWidth: 13 },
      7: { cellWidth: 16 },
      8: { cellWidth: 14 },
      9: { cellWidth: 14 },
      10: { cellWidth: 28 }
    }
  })

  pdf.save('issue-register.pdf')
}

const downloadDrawnRegisterPdf = async (records: DrawnRecord[]): Promise<void> => {
  const pdf = new jsPDF({ orientation: 'landscape', format: 'a4' })
  const pageWidth = pdf.internal.pageSize.getWidth()
  const marginLeft = (pageWidth - DRAWN_PDF_TABLE_WIDTH) / 2

  await drawPdfBottomBackgroundLogo(pdf)
  await drawPdfHeaderLogo(pdf)

  pdf.setTextColor(20, 20, 20)
  pdf.setFont('times', 'bold')
  pdf.setFontSize(12)
  pdf.text('ULTRA TESTING & RESEARCH LABORATORY', 148.5, 13, { align: 'center' })
  pdf.setFontSize(11)
  pdf.text('SAMPLE RECEIVING REGISTER', 148.5, 19, { align: 'center' })

  const headers = [
    'Sr.No.',
    'Sample Description',
    'Sample Drawn on',
    'Sample Drawn By',
    'Customer Name &\nAddress',
    'Parameter to be\nTested',
    'Report Due On',
    'Sample Received\nBy'
  ]

  const rows = records.map((record) => [
    record.srNo,
    record.sampleDescription,
    record.sampleDrawnOn,
    record.sampleDrawnBy,
    record.customerNameAddress,
    record.parameterToBeTested,
    record.reportDueOn,
    record.sampleReceivedBy
  ])

  autoTable(pdf, {
    head: [headers],
    body: rows,
    startY: 34,
    margin: { top: 34, right: marginLeft, bottom: 8, left: marginLeft },
    tableWidth: DRAWN_PDF_TABLE_WIDTH,
    theme: 'grid',
    styles: {
      font: 'times',
      fontSize: 7.2,
      cellPadding: 1.2,
      lineColor: [70, 70, 70],
      lineWidth: 0.1,
      textColor: [20, 20, 20],
      valign: 'middle'
    },
    headStyles: {
      font: 'times',
      fontStyle: 'bold',
      fillColor: [255, 255, 255],
      textColor: [20, 20, 20],
      lineColor: [70, 70, 70],
      lineWidth: 0.1
    },
    columnStyles: {
      0: { cellWidth: 12 },
      1: { cellWidth: 24 },
      2: { cellWidth: 20 },
      3: { cellWidth: 18 },
      4: { cellWidth: 34 },
      5: { cellWidth: 34 },
      6: { cellWidth: 18 },
      7: { cellWidth: 24 }
    }
  })

  pdf.save('drawn-sample-register.pdf')
}

const toIssueCreatePayload = (record: IssueRecord): IssueRecord => ({
  status: getIssueStatus(record),
  srNo: record.srNo,
  codeNo: record.codeNo,
  sampleDescription: record.sampleDescription,
  parameterToBeTested: record.parameterToBeTested,
  issuedOn: record.issuedOn,
  issuedBy: record.issuedBy,
  issuedTo: record.issuedTo,
  reportDueOn: record.reportDueOn,
  receivedBy: record.receivedBy,
  reportedOn: record.reportedOn,
  reportedByRemarks: record.reportedByRemarks
})

const toDrawnCreatePayload = (record: DrawnRecord): DrawnRecord => ({
  status: getDrawnStatus(record),
  srNo: record.srNo,
  sampleDescription: record.sampleDescription,
  sampleDrawnOn: record.sampleDrawnOn,
  sampleDrawnBy: record.sampleDrawnBy,
  customerNameAddress: record.customerNameAddress,
  parameterToBeTested: record.parameterToBeTested,
  reportDueOn: record.reportDueOn,
  sampleReceivedBy: record.sampleReceivedBy
})

const normalizeCategoryFilter = (value: string): string => value.trim().toLowerCase()

const matchesCategory = (sampleDescription: string, selectedCategory: string): boolean => {
  const category = normalizeCategoryFilter(selectedCategory)
  if (!category || category === 'all') {
    return true
  }

  const description = sampleDescription.trim().toLowerCase()
  return description.includes(category)
}

const resetIssueFilters = (): void => {
  issueSearch = ''
  issueFromDate = ''
  issueToDate = ''
  issueIssuedByFilter = ''
  issueIssuedToFilter = ''
  issueParameterFilter = ''
  issueCategoryFilter = 'All'
}

const resetDrawnFilters = (): void => {
  drawnSearch = ''
  drawnFromDate = ''
  drawnToDate = ''
  drawnByFilter = ''
  drawnCustomerFilter = ''
  drawnParameterFilter = ''
  drawnCategoryFilter = 'All'
}

const downloadCsv = (fileName: string, headers: string[], rows: string[][]): void => {
  const csv = [headers, ...rows].map((row) => row.map((item) => escapeCsv(item)).join(',')).join('\n')
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
  URL.revokeObjectURL(url)
}

const getFilteredIssueRecords = (): IssueRecord[] => {
  const query = issueSearch.trim().toLowerCase()
  const issuedByFilter = issueIssuedByFilter.trim().toLowerCase()
  const issuedToFilter = issueIssuedToFilter.trim().toLowerCase()
  const parameterFilter = issueParameterFilter.trim().toLowerCase()
  const categoryFilter = issueCategoryFilter.trim().toLowerCase()

  return issueRecords.filter((record) => {
    const dateValue = toDateValue(record.issuedOn)
    const from = issueFromDate ? toDateValue(issueFromDate) : 0
    const to = issueToDate ? toDateValue(issueToDate) : Number.MAX_SAFE_INTEGER
    const inDateRange = dateValue >= from && dateValue <= to
    const matchesIssuedBy = !issuedByFilter || record.issuedBy.toLowerCase().includes(issuedByFilter)
    const matchesIssuedTo = !issuedToFilter || record.issuedTo.toLowerCase().includes(issuedToFilter)
    const matchesParameter = !parameterFilter || record.parameterToBeTested.toLowerCase().includes(parameterFilter)
    const matchesSelectedCategory = matchesCategory(record.sampleDescription, categoryFilter)

    if (!query) {
      return inDateRange && matchesIssuedBy && matchesIssuedTo && matchesParameter && matchesSelectedCategory
    }

    const serial = record.srNo.toLowerCase()
    return inDateRange && matchesIssuedBy && matchesIssuedTo && matchesParameter && matchesSelectedCategory && serial.includes(query)
  })
}

const getFilteredDrawnRecords = (): DrawnRecord[] => {
  const query = drawnSearch.trim().toLowerCase()
  const sampleByFilter = drawnByFilter.trim().toLowerCase()
  const customerFilter = drawnCustomerFilter.trim().toLowerCase()
  const parameterFilter = drawnParameterFilter.trim().toLowerCase()
  const categoryFilter = drawnCategoryFilter.trim().toLowerCase()

  return drawnRecords.filter((record) => {
    const dateValue = toDateValue(record.sampleDrawnOn)
    const from = drawnFromDate ? toDateValue(drawnFromDate) : 0
    const to = drawnToDate ? toDateValue(drawnToDate) : Number.MAX_SAFE_INTEGER
    const inDateRange = dateValue >= from && dateValue <= to
    const matchesDrawnBy = !sampleByFilter || record.sampleDrawnBy.toLowerCase().includes(sampleByFilter)
    const matchesCustomer = !customerFilter || record.customerNameAddress.toLowerCase().includes(customerFilter)
    const matchesParameter = !parameterFilter || record.parameterToBeTested.toLowerCase().includes(parameterFilter)
    const matchesSelectedCategory = matchesCategory(record.sampleDescription, categoryFilter)

    if (!query) {
      return inDateRange && matchesDrawnBy && matchesCustomer && matchesParameter && matchesSelectedCategory
    }

    const serial = record.srNo.toLowerCase()
    return inDateRange && matchesDrawnBy && matchesCustomer && matchesParameter && matchesSelectedCategory && serial.includes(query)
  })
}

const saveSession = (session: Session): void => {
  localStorage.setItem(SESSION_KEY, JSON.stringify(session))
}

const clearSession = (): void => {
  localStorage.removeItem(SESSION_KEY)
}

const getSession = (): Session | null => {
  const rawSession = localStorage.getItem(SESSION_KEY)
  if (!rawSession) {
    return null
  }

  try {
    const parsed = JSON.parse(rawSession) as Session
    if (!parsed.token || !parsed.email) {
      return null
    }
    return parsed
  } catch {
    return null
  }
}

const readJsonSafe = async <T>(response: Response): Promise<T> => {
  const raw = await response.text()
  if (!raw) {
    return {} as T
  }

  try {
    return JSON.parse(raw) as T
  } catch {
    return {} as T
  }
}

const authenticate = async (email: string, password: string): Promise<LoginResponse> => {
  const response = await fetch('/api/login', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ email, password })
  })

  const body = await readJsonSafe<Partial<LoginResponse> & { message?: string }>(response)

  if (!response.ok) {
    throw new Error(body.message ?? 'Login failed. Try again.')
  }

  if (!body.token || !body.user?.email) {
    throw new Error('Invalid server response.')
  }

  return {
    token: body.token,
    user: {
      email: body.user.email,
      role: normalizeRole(body.user.role, body.user.email)
    }
  }
}

const fetchCurrentUser = async (token: string): Promise<MeResponse> => {
  const response = await fetch('/api/me', {
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  const body = await readJsonSafe<Partial<MeResponse> & { message?: string }>(response)
  if (!response.ok || !body.user?.email) {
    throw new Error(body.message ?? 'Session expired. Please login again.')
  }

  return {
    user: {
      email: body.user.email,
      role: normalizeRole(body.user.role, body.user.email)
    }
  }
}

const loadRegisters = async (token: string): Promise<void> => {
  const response = await fetch('/api/registers', {
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  const body = await readJsonSafe<Partial<RegistersResponse> & { message?: string }>(response)
  if (!response.ok) {
    throw new Error(body.message ?? 'Failed to load registers.')
  }

  const nextIssueRecords = (Array.isArray(body.issueRecords) ? body.issueRecords : []).map((record) => ({
    ...record,
    status: normalizeRecordStatus(record.status, String(record.reportedOn ?? '').trim() ? 'Reported' : 'Pending')
  }))

  const nextDrawnRecords = (Array.isArray(body.drawnRecords) ? body.drawnRecords : []).map((record) => ({
    ...record,
    status: normalizeRecordStatus(record.status, 'Pending')
  }))

  setIssueRecords(nextIssueRecords)
  setDrawnRecords(nextDrawnRecords)
}

const loadTestMaster = async (token: string): Promise<void> => {
  const response = await fetch('/api/test-master', {
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  const body = await readJsonSafe<Partial<TestMasterResponse> & { message?: string }>(response)

  if (!response.ok) {
    throw new Error(body.message ?? 'Failed to load test master.')
  }

  setTestMaster(Array.isArray(body.tests) ? body.tests : [], Array.isArray(body.parameters) ? body.parameters : [])
}

const createIssueRecord = async (token: string, record: IssueRecord): Promise<IssueRecord> => {
  const response = await fetch('/api/registers/issue', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(record)
  })

  const body = await readJsonSafe<{ message?: string; record?: IssueRecord }>(response)
  if (!response.ok || !body.record) {
    throw new Error(body.message ?? 'Failed to save issue register entry.')
  }

  return body.record
}

const updateIssueRecord = async (token: string, recordId: string, record: IssueRecord): Promise<IssueRecord> => {
  const response = await fetch(`/api/registers/issue/${recordId}`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(record)
  })

  const body = await readJsonSafe<{ message?: string; record?: IssueRecord }>(response)
  if (!response.ok || !body.record) {
    throw new Error(body.message ?? 'Failed to update issue register entry.')
  }

  return body.record
}

const deleteIssueRecordApi = async (token: string, recordId: string): Promise<void> => {
  const response = await fetch(`/api/registers/issue/${recordId}`, {
    method: 'DELETE',
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  if (!response.ok) {
    throw new Error('Failed to delete issue register entry.')
  }
}

const createDrawnRecord = async (token: string, record: DrawnRecord): Promise<DrawnRecord> => {
  const response = await fetch('/api/registers/drawn', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(record)
  })

  const body = await readJsonSafe<{ message?: string; record?: DrawnRecord }>(response)
  if (!response.ok || !body.record) {
    throw new Error(body.message ?? 'Failed to save drawn register entry.')
  }

  return body.record
}

const updateDrawnRecord = async (token: string, recordId: string, record: DrawnRecord): Promise<DrawnRecord> => {
  const response = await fetch(`/api/registers/drawn/${recordId}`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(record)
  })

  const body = await readJsonSafe<{ message?: string; record?: DrawnRecord }>(response)
  if (!response.ok || !body.record) {
    throw new Error(body.message ?? 'Failed to update drawn register entry.')
  }

  return body.record
}

const deleteDrawnRecordApi = async (token: string, recordId: string): Promise<void> => {
  const response = await fetch(`/api/registers/drawn/${recordId}`, {
    method: 'DELETE',
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  if (!response.ok) {
    throw new Error('Failed to delete drawn register entry.')
  }
}

const loadAdminPanelData = async (token: string): Promise<void> => {
  const [usersResponse, alertsResponse, auditResponse, historyResponse, backupResponse] = await Promise.all([
    fetch('/api/admin/users', { headers: { Authorization: `Bearer ${token}` } }),
    fetch('/api/admin/alerts', { headers: { Authorization: `Bearer ${token}` } }),
    fetch('/api/admin/audit?limit=25', { headers: { Authorization: `Bearer ${token}` } }),
    fetch('/api/admin/register-history?limit=25', { headers: { Authorization: `Bearer ${token}` } }),
    fetch('/api/admin/backups', { headers: { Authorization: `Bearer ${token}` } })
  ])

  const [usersBody, alertsBody, auditBody, historyBody, backupBody] = await Promise.all([
    readJsonSafe<{ users?: AdminUser[]; message?: string }>(usersResponse),
    readJsonSafe<{ alerts?: AdminAlert[]; message?: string }>(alertsResponse),
    readJsonSafe<{ entries?: AuditEntry[]; message?: string }>(auditResponse),
    readJsonSafe<{ entries?: RegisterHistoryEntry[]; message?: string }>(historyResponse),
    readJsonSafe<{ backups?: string[]; message?: string }>(backupResponse)
  ])

  if (!usersResponse.ok) {
    throw new Error(usersBody.message ?? 'Failed to load admin users.')
  }

  if (!alertsResponse.ok) {
    throw new Error(alertsBody.message ?? 'Failed to load alerts.')
  }

  if (!auditResponse.ok) {
    throw new Error(auditBody.message ?? 'Failed to load audit entries.')
  }

  if (!historyResponse.ok) {
    throw new Error(historyBody.message ?? 'Failed to load register history entries.')
  }

  if (!backupResponse.ok) {
    throw new Error(backupBody.message ?? 'Failed to load backups.')
  }

  setAdminUsers(Array.isArray(usersBody.users) ? usersBody.users : [])
  setAdminAlerts(Array.isArray(alertsBody.alerts) ? alertsBody.alerts : [])
  setAdminAuditEntries(Array.isArray(auditBody.entries) ? auditBody.entries : [])
  setAdminRegisterHistoryEntries(Array.isArray(historyBody.entries) ? historyBody.entries : [])
  setAdminBackups(Array.isArray(backupBody.backups) ? backupBody.backups : [])
}

const createAdminUser = async (token: string, payload: { email: string; password: string; role: UserRole }): Promise<void> => {
  const response = await fetch('/api/admin/users', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  })

  if (!response.ok) {
    const body = await readJsonSafe<{ message?: string }>(response)
    throw new Error(body.message ?? 'Failed to create user.')
  }
}

const updateAdminUserStatus = async (token: string, userId: string, isActive: boolean): Promise<void> => {
  const response = await fetch(`/api/admin/users/${userId}/status`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ isActive })
  })

  if (!response.ok) {
    const body = await readJsonSafe<{ message?: string }>(response)
    throw new Error(body.message ?? 'Failed to update user status.')
  }
}

const resetAdminUserPassword = async (token: string, userId: string, password: string): Promise<void> => {
  const response = await fetch(`/api/admin/users/${userId}/reset-password`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ password })
  })

  if (!response.ok) {
    const body = await readJsonSafe<{ message?: string }>(response)
    throw new Error(body.message ?? 'Failed to reset user password.')
  }
}

const createBackup = async (token: string): Promise<string> => {
  const response = await fetch('/api/admin/backup', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  const body = await readJsonSafe<{ fileName?: string; message?: string }>(response)
  if (!response.ok || !body.fileName) {
    throw new Error(body.message ?? 'Failed to create backup.')
  }

  return body.fileName
}

const restoreBackup = async (token: string, fileName: string): Promise<void> => {
  const response = await fetch('/api/admin/restore', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ fileName })
  })

  if (!response.ok) {
    const body = await readJsonSafe<{ message?: string }>(response)
    throw new Error(body.message ?? 'Failed to restore backup.')
  }
}

const getModuleLabel = (module: ModuleKey): string => {
  if (module === 'issue-entry') {
    return 'Sample Issue Register (Entry)'
  }

  if (module === 'issue-records') {
    return 'Sample Issue Register (Records)'
  }

  if (module === 'drawn-entry') {
    return 'Sample Receiving Register (Entry)'
  }

  if (module === 'drawn-records') {
    return 'Sample Receiving Register (Records)'
  }

  return 'Admin Panel'
}

const getMenuItems = (role: UserRole): ModuleKey[] => {
  if (role === 'admin') {
    return ['issue-entry', 'issue-records', 'drawn-entry', 'drawn-records', 'admin-panel']
  }

  return ['issue-entry', 'issue-records', 'drawn-entry', 'drawn-records']
}

const renderIssueTable = (canDelete: boolean): string => {
  const filtered = getFilteredIssueRecords()

  if (filtered.length === 0) {
    return '<p class="empty-state">No entries yet.</p>'
  }

  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Sr.No.</th>
            <th>Code No.</th>
            <th>Status</th>
            <th>Sample Description</th>
            <th>Parameter to be tested</th>
            <th>Issued On</th>
            <th>Issued By</th>
            <th>Issued To</th>
            <th>Report Due On</th>
            <th>Received By</th>
            <th>Reported On</th>
            <th>ReportedBy/Remarks</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${filtered
            .map(
              (item, index) => {
                const status = getIssueStatus(item)
                const overdue = isOverdue(item.reportDueOn, status, item.reportedOn)
                return `
            <tr class="${overdue ? 'row-overdue' : ''}">
              <td>${index + 1}</td>
              <td>${escapeHtml(item.codeNo)}</td>
              <td><span class="status-chip status-${status.toLowerCase().replace(/\s+/g, '-')}">${escapeHtml(status)}</span></td>
              <td>${escapeHtml(item.sampleDescription)}</td>
              <td>${escapeHtml(item.parameterToBeTested)}</td>
              <td>${escapeHtml(item.issuedOn)}</td>
              <td>${escapeHtml(item.issuedBy)}</td>
              <td>${escapeHtml(item.issuedTo)}</td>
              <td><span class="due-chip ${overdue ? 'overdue' : ''}">${escapeHtml(item.reportDueOn)}${overdue ? ' • Overdue' : ''}</span></td>
              <td>${escapeHtml(item.receivedBy)}</td>
              <td>${escapeHtml(item.reportedOn)}</td>
              <td>${escapeHtml(item.reportedByRemarks)}</td>
              <td class="actions-col">
                <button class="table-action edit" data-issue-edit="${escapeHtml(item.id ?? '')}" type="button">Edit</button>
                ${canDelete ? `<button class="table-action delete" data-issue-delete="${escapeHtml(item.id ?? '')}" type="button">Delete</button>` : ''}
              </td>
            </tr>
          `
              }
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `
}

const renderDrawnTable = (canDelete: boolean): string => {
  const filtered = getFilteredDrawnRecords()

  if (filtered.length === 0) {
    return '<p class="empty-state">No entries yet.</p>'
  }

  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Sr.No.</th>
            <th>Status</th>
            <th>Sample Description</th>
            <th>Sample Drawn on</th>
            <th>Sample Drawn By</th>
            <th>Customer Name & Address</th>
            <th>Parameter to be Tested</th>
            <th>Report Due On</th>
            <th>Sample Received By</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${filtered
            .map(
              (item, index) => {
                const status = getDrawnStatus(item)
                const overdue = isOverdue(item.reportDueOn, status)
                return `
            <tr class="${overdue ? 'row-overdue' : ''}">
              <td>${index + 1}</td>
              <td><span class="status-chip status-${status.toLowerCase().replace(/\s+/g, '-')}">${escapeHtml(status)}</span></td>
              <td>${escapeHtml(item.sampleDescription)}</td>
              <td>${escapeHtml(item.sampleDrawnOn)}</td>
              <td>${escapeHtml(item.sampleDrawnBy)}</td>
              <td>${escapeHtml(item.customerNameAddress)}</td>
              <td>${escapeHtml(item.parameterToBeTested)}</td>
              <td><span class="due-chip ${overdue ? 'overdue' : ''}">${escapeHtml(item.reportDueOn)}${overdue ? ' • Overdue' : ''}</span></td>
              <td>${escapeHtml(item.sampleReceivedBy)}</td>
              <td class="actions-col">
                <button class="table-action edit" data-drawn-edit="${escapeHtml(item.id ?? '')}" type="button">Edit</button>
                ${canDelete ? `<button class="table-action delete" data-drawn-delete="${escapeHtml(item.id ?? '')}" type="button">Delete</button>` : ''}
              </td>
            </tr>
          `
              }
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `
}

const renderIssueEntryModule = (): string => {
  const editing = issueEditingId ? issueRecords.find((item) => item.id === issueEditingId) : undefined
  const { issueDraft, srNoValue, codeNoValue, descriptionValue, parameterValue } = getIssueFormInitialValues(editing)
  const issueStatus = editing
    ? getIssueStatus(editing)
    : normalizeRecordStatus(issueDraft.status, issueDraft.reportedOn ? 'Reported' : 'Pending')

  return `
    <section class="module-card">
      <div class="register-head">
        <p class="register-lab">ULTRA TESTING & RESEARCH LABORATORY</p>
        <h3>SAMPLE ISSUE REGISTER</h3>
        <p class="register-note">Maintain issue, due, and reporting trail for each sample entry.</p>
      </div>
      <form id="issue-form" class="data-form" novalidate>
        ${renderIssueAutoFields(srNoValue, codeNoValue, descriptionValue, parameterValue, Boolean(editing))}
        ${renderIssueDefaultDescriptionMessage(descriptionValue)}
        <label class="field-group"><span>Issued On</span><input name="issuedOn" type="date" value="${escapeHtml(editing?.issuedOn ?? issueDraft.issuedOn ?? '')}" required /></label>
        <label class="field-group"><span>Issued By</span><input name="issuedBy" value="${escapeHtml(editing?.issuedBy ?? issueDraft.issuedBy ?? '')}" required /></label>
        <label class="field-group"><span>Issued To</span><input name="issuedTo" value="${escapeHtml(editing?.issuedTo ?? issueDraft.issuedTo ?? '')}" required /></label>
        <label class="field-group"><span>Status</span>
          <select name="status" required>
            <option value="Pending" ${issueStatus === 'Pending' ? 'selected' : ''}>Pending</option>
            <option value="In Progress" ${issueStatus === 'In Progress' ? 'selected' : ''}>In Progress</option>
            <option value="Reported" ${issueStatus === 'Reported' ? 'selected' : ''}>Reported</option>
          </select>
        </label>
        <label class="field-group"><span>Report Due On</span><input name="reportDueOn" type="date" value="${escapeHtml(editing?.reportDueOn ?? issueDraft.reportDueOn ?? '')}" required /></label>
        <label class="field-group"><span>Received By</span><input name="receivedBy" value="${escapeHtml(editing?.receivedBy ?? issueDraft.receivedBy ?? '')}" required /></label>
        <label class="field-group"><span>Reported On</span><input name="reportedOn" type="date" value="${escapeHtml(editing?.reportedOn ?? issueDraft.reportedOn ?? '')}" /></label>
        <label class="field-group"><span>ReportedBy/Remarks</span><input name="reportedByRemarks" value="${escapeHtml(editing?.reportedByRemarks ?? issueDraft.reportedByRemarks ?? '')}" /></label>
        <div class="form-actions">
          <button class="primary-btn" type="submit">${editing ? 'Update Entry' : 'Add Entry'}</button>
          ${editing ? '<button id="issue-cancel-edit" class="secondary-btn light" type="button">Cancel Edit</button>' : ''}
        </div>
      </form>
      ${renderDraftHint(Boolean(editing))}
    </section>
  `
}

const renderIssueRecordsModule = (): string => {
  return `
    <section class="module-card records-page">
      <div class="register-head">
        <p class="register-lab">ULTRA TESTING & RESEARCH LABORATORY</p>
        <h3>SAMPLE ISSUE REGISTER RECORDS</h3>
        <p class="register-note">All issued sample entries with full status trail.</p>
      </div>
      <div class="module-toolbar">
        <input id="issue-search" placeholder="Search by Sr.No." value="${escapeHtml(issueSearch)}" />
        <label class="toolbar-filter" for="issue-filter-category">Filter by Category
          <select id="issue-filter-category">
            <option value="All" ${issueCategoryFilter === 'All' ? 'selected' : ''}>All</option>
            <option value="Air" ${issueCategoryFilter === 'Air' ? 'selected' : ''}>Air</option>
            <option value="Water" ${issueCategoryFilter === 'Water' ? 'selected' : ''}>Water</option>
            <option value="Soil" ${issueCategoryFilter === 'Soil' ? 'selected' : ''}>Soil</option>
            <option value="Noise" ${issueCategoryFilter === 'Noise' ? 'selected' : ''}>Noise</option>
          </select>
        </label>
        <input id="issue-from" type="date" value="${escapeHtml(issueFromDate)}" />
        <input id="issue-to" type="date" value="${escapeHtml(issueToDate)}" />
        <input id="issue-filter-issued-by" placeholder="Filter: Issued By" value="${escapeHtml(issueIssuedByFilter)}" />
        <input id="issue-filter-issued-to" placeholder="Filter: Issued To" value="${escapeHtml(issueIssuedToFilter)}" />
        <input id="issue-filter-parameter" placeholder="Filter: Parameter" value="${escapeHtml(issueParameterFilter)}" />
        <button id="issue-filter-reset" class="secondary-btn light" type="button">Reset Filters</button>
        <button id="issue-export" class="secondary-btn light" type="button">Export CSV</button>
        <button id="issue-export-pdf" class="secondary-btn light" type="button">Export PDF</button>
      </div>
      ${renderIssueTable(activeSession?.role === 'admin')}
    </section>
  `
}

const renderDrawnEntryModule = (): string => {
  const editing = drawnEditingId ? drawnRecords.find((item) => item.id === drawnEditingId) : undefined
  const { drawnDraft, srNoValue, descriptionValue, parameterValue } = getDrawnFormInitialValues(editing)
  const drawnStatus = editing ? getDrawnStatus(editing) : normalizeRecordStatus(drawnDraft.status, 'Pending')

  return `
    <section class="module-card">
      <div class="register-head">
        <p class="register-lab">ULTRA TESTING & RESEARCH LABORATORY</p>
        <h3>SAMPLE RECEIVING REGISTER</h3>
        <p class="register-note">Capture receiving details for drawn samples with due-date tracking.</p>
      </div>
      <form id="drawn-form" class="data-form" novalidate>
        ${renderDrawnAutoFields(srNoValue, descriptionValue, parameterValue, Boolean(editing))}
        ${renderDrawnDefaultDescriptionMessage(descriptionValue)}
        <label class="field-group"><span>Sample Drawn on</span><input name="sampleDrawnOn" type="date" value="${escapeHtml(editing?.sampleDrawnOn ?? drawnDraft.sampleDrawnOn ?? '')}" required /></label>
        <label class="field-group"><span>Sample Drawn By</span><input name="sampleDrawnBy" value="${escapeHtml(editing?.sampleDrawnBy ?? drawnDraft.sampleDrawnBy ?? '')}" required /></label>
        <label class="field-group"><span>Customer Name & Address</span><input name="customerNameAddress" value="${escapeHtml(editing?.customerNameAddress ?? drawnDraft.customerNameAddress ?? '')}" required /></label>
        <label class="field-group"><span>Status</span>
          <select name="status" required>
            <option value="Pending" ${drawnStatus === 'Pending' ? 'selected' : ''}>Pending</option>
            <option value="In Progress" ${drawnStatus === 'In Progress' ? 'selected' : ''}>In Progress</option>
            <option value="Reported" ${drawnStatus === 'Reported' ? 'selected' : ''}>Reported</option>
          </select>
        </label>
        <label class="field-group"><span>Report Due On</span><input name="reportDueOn" type="date" value="${escapeHtml(editing?.reportDueOn ?? drawnDraft.reportDueOn ?? '')}" required /></label>
        <label class="field-group"><span>Sample Received By</span><input name="sampleReceivedBy" value="${escapeHtml(editing?.sampleReceivedBy ?? drawnDraft.sampleReceivedBy ?? '')}" required /></label>
        <div class="form-actions">
          <button class="primary-btn" type="submit">${editing ? 'Update Entry' : 'Add Entry'}</button>
          ${editing ? '<button id="drawn-cancel-edit" class="secondary-btn light" type="button">Cancel Edit</button>' : ''}
        </div>
      </form>
      ${renderDraftHint(Boolean(editing))}
    </section>
  `
}

const renderDrawnRecordsModule = (): string => {
  return `
    <section class="module-card records-page">
      <div class="register-head">
        <p class="register-lab">ULTRA TESTING & RESEARCH LABORATORY</p>
        <h3>SAMPLE RECEIVING REGISTER RECORDS</h3>
        <p class="register-note">All received sample entries with drawing and due-date details.</p>
      </div>
      <div class="module-toolbar">
        <input id="drawn-search" placeholder="Search by Sr.No." value="${escapeHtml(drawnSearch)}" />
        <label class="toolbar-filter" for="drawn-filter-category">Filter by Category
          <select id="drawn-filter-category">
            <option value="All" ${drawnCategoryFilter === 'All' ? 'selected' : ''}>All</option>
            <option value="Air" ${drawnCategoryFilter === 'Air' ? 'selected' : ''}>Air</option>
            <option value="Water" ${drawnCategoryFilter === 'Water' ? 'selected' : ''}>Water</option>
            <option value="Soil" ${drawnCategoryFilter === 'Soil' ? 'selected' : ''}>Soil</option>
            <option value="Noise" ${drawnCategoryFilter === 'Noise' ? 'selected' : ''}>Noise</option>
          </select>
        </label>
        <input id="drawn-from" type="date" value="${escapeHtml(drawnFromDate)}" />
        <input id="drawn-to" type="date" value="${escapeHtml(drawnToDate)}" />
        <input id="drawn-filter-by" placeholder="Filter: Drawn By" value="${escapeHtml(drawnByFilter)}" />
        <input id="drawn-filter-customer" placeholder="Filter: Customer" value="${escapeHtml(drawnCustomerFilter)}" />
        <input id="drawn-filter-parameter" placeholder="Filter: Parameter" value="${escapeHtml(drawnParameterFilter)}" />
        <button id="drawn-filter-reset" class="secondary-btn light" type="button">Reset Filters</button>
        <button id="drawn-export" class="secondary-btn light" type="button">Export CSV</button>
        <button id="drawn-export-pdf" class="secondary-btn light" type="button">Export PDF</button>
      </div>
      ${renderDrawnTable(activeSession?.role === 'admin')}
    </section>
  `
}

const renderActivityCards = (): string => {
  const activity = getActivityStats()

  return `
    <section class="activity-cards">
      <article class="activity-card">
        <h4>Total Entries</h4>
        <p>${activity.totalEntries}</p>
      </article>
      <article class="activity-card ${activity.overdueEntries ? 'attention' : ''}">
        <h4>Overdue Entries</h4>
        <p>${activity.overdueEntries}</p>
      </article>
      <article class="activity-card">
        <h4>Today's New Entries</h4>
        <p>${activity.todayEntries}</p>
      </article>
    </section>
  `
}

const renderAdminModule = (): string => {
  const issueCount = issueRecords.length
  const drawnCount = drawnRecords.length
  const totalCount = issueCount + drawnCount
  const today = new Date().toISOString().slice(0, 10)

  const issueToday = issueRecords.filter((record) => record.issuedOn === today).length
  const drawnToday = drawnRecords.filter((record) => record.sampleDrawnOn === today).length

  const pendingIssue = issueRecords.filter((record) => getIssueStatus(record) !== 'Reported').length
  const recentIssue = [...issueRecords]
    .sort((first, second) => toDateValue(second.createdAt ?? second.issuedOn) - toDateValue(first.createdAt ?? first.issuedOn))
    .slice(0, 5)

  const recentDrawn = [...drawnRecords]
    .sort((first, second) => toDateValue(second.createdAt ?? second.sampleDrawnOn) - toDateValue(first.createdAt ?? first.sampleDrawnOn))
    .slice(0, 5)

  const activeUsers = adminUsers.filter((user) => user.isActive).length
  const disabledUsers = adminUsers.filter((user) => !user.isActive).length

  return `
    <section class="module-card">
      <div class="register-head">
        <p class="register-lab">ULTRA TESTING & RESEARCH LABORATORY</p>
        <h3>ADMIN PANEL</h3>
        <p class="register-note">Manage users, alerts, backups and monitor activity.</p>
      </div>

      <div class="admin-stats">
        <article class="admin-stat-card">
          <h4>Total Records</h4>
          <p>${totalCount}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Issue Register</h4>
          <p>${issueCount}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Receiving Register</h4>
          <p>${drawnCount}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Pending Reports</h4>
          <p>${pendingIssue}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Issued Today</h4>
          <p>${issueToday}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Received Today</h4>
          <p>${drawnToday}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Active Users</h4>
          <p>${activeUsers}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Disabled Users</h4>
          <p>${disabledUsers}</p>
        </article>
        <article class="admin-stat-card">
          <h4>Due Alerts</h4>
          <p>${adminAlerts.length}</p>
        </article>
      </div>

      ${adminMessage ? `<p class="message" data-state="${adminMessageState}">${escapeHtml(adminMessage)}</p>` : ''}

      <div class="form-actions admin-actions">
        <button id="admin-refresh" class="secondary-btn light" type="button">Refresh Data</button>
        <button id="admin-export-issue" class="secondary-btn light" type="button">Export Issue CSV</button>
        <button id="admin-export-drawn" class="secondary-btn light" type="button">Export Receiving CSV</button>
        <button id="admin-backup-create" class="secondary-btn light" type="button">Create Backup</button>
      </div>

      <section class="admin-users">
        <h4>User Management</h4>
        <form id="admin-user-form" class="module-toolbar" novalidate>
          <input name="email" type="email" placeholder="staff@labsoft.local" required />
          <input name="password" type="password" placeholder="Temp password" required />
          <select name="role">
            <option value="staff">Staff</option>
            <option value="admin">Admin</option>
          </select>
          <button class="secondary-btn light" type="submit">Add User</button>
        </form>
        <div class="table-wrap admin-table-wrap">
          <table>
            <thead>
              <tr>
                <th>Email</th>
                <th>Role</th>
                <th>Status</th>
                <th>Created At</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              ${
                adminUsers.length
                  ? adminUsers
                      .map(
                        (user) => `
                    <tr>
                      <td>${escapeHtml(user.email)}</td>
                      <td>${escapeHtml(user.role)}</td>
                      <td>${user.isActive ? 'Active' : 'Disabled'}</td>
                      <td>${escapeHtml(user.createdAt.slice(0, 10))}</td>
                      <td class="actions-col">
                        <button class="table-action" data-admin-toggle="${escapeHtml(user.id)}" data-next-state="${user.isActive ? 'disable' : 'enable'}" type="button">${user.isActive ? 'Disable' : 'Enable'}</button>
                        <button class="table-action" data-admin-reset="${escapeHtml(user.id)}" type="button">Reset Password</button>
                      </td>
                    </tr>
                  `
                      )
                      .join('')
                  : '<tr><td colspan="5">No users found.</td></tr>'
              }
            </tbody>
          </table>
        </div>
      </section>

      <section class="admin-alerts">
        <h4>Due Alerts</h4>
        ${
          adminAlerts.length
            ? `<ul>${adminAlerts
                .map(
                  (alert) => `<li><strong>${escapeHtml(alert.type.toUpperCase())}</strong> • ${escapeHtml(alert.message)} • Due: ${escapeHtml(alert.dueOn)}</li>`
                )
                .join('')}</ul>`
            : '<p class="admin-note">No alerts right now.</p>'
        }
      </section>

      <section class="admin-backups">
        <h4>Restore Backup</h4>
        <div class="form-actions">
          <select id="admin-backup-select">
            ${
              adminBackups.length
                ? adminBackups.map((fileName) => `<option value="${escapeHtml(fileName)}">${escapeHtml(fileName)}</option>`).join('')
                : '<option value="">No backup available</option>'
            }
          </select>
          <button id="admin-backup-restore" class="secondary-btn light" type="button">Restore Selected</button>
        </div>
      </section>

      <div class="admin-lists">
        <section class="admin-list-card">
          <h4>Recent Issue Entries</h4>
          ${
            recentIssue.length
              ? `<ul>${recentIssue
                  .map(
                    (record) =>
                      `<li><strong>${escapeHtml(record.srNo)}</strong> • ${escapeHtml(record.sampleDescription)} • ${escapeHtml(record.issuedOn)}</li>`
                  )
                  .join('')}</ul>`
              : '<p class="admin-note">No issue entries yet.</p>'
          }
        </section>

        <section class="admin-list-card">
          <h4>Recent Receiving Entries</h4>
          ${
            recentDrawn.length
              ? `<ul>${recentDrawn
                  .map(
                    (record) =>
                      `<li><strong>${escapeHtml(record.srNo)}</strong> • ${escapeHtml(record.sampleDescription)} • ${escapeHtml(record.sampleDrawnOn)}</li>`
                  )
                  .join('')}</ul>`
              : '<p class="admin-note">No receiving entries yet.</p>'
          }
        </section>
        <section class="admin-list-card">
          <h4>Recent Audit Trail</h4>
          ${
            adminAuditEntries.length
              ? `<ul>${adminAuditEntries
                  .slice(0, 6)
                  .map(
                    (entry) =>
                      `<li><strong>${escapeHtml(entry.action)}</strong> • ${escapeHtml(entry.actor)} • ${escapeHtml(entry.createdAt.slice(0, 16).replace('T', ' '))}</li>`
                  )
                  .join('')}</ul>`
              : '<p class="admin-note">No audit entries yet.</p>'
          }
        </section>
        <section class="admin-list-card">
          <h4>Recent Register History</h4>
          ${
            adminRegisterHistoryEntries.length
              ? `<ul>${adminRegisterHistoryEntries
                  .slice(0, 6)
                  .map(
                    (entry) =>
                      `<li><strong>${escapeHtml(entry.action)}</strong> • ${escapeHtml(entry.source.toUpperCase())} • Sr.No. ${escapeHtml(entry.srNo)} • ${escapeHtml(entry.createdAt.slice(0, 16).replace('T', ' '))}</li>`
                  )
                  .join('')}</ul>`
              : '<p class="admin-note">No register history yet.</p>'
          }
        </section>
      </div>
    </section>
  `
}

const renderLogin = (message = '', routeMode: 'push' | 'replace' | null = null): void => {
  activeSession = null
  currentView = 'login'

  if (routeMode) {
    updateUrlView('login', routeMode)
  }

  app.innerHTML = `
    <main class="layout">
      <section class="brand-panel">
        <p class="brand-kicker">LABSOFT</p>
        <h1>Welcome back</h1>
        <p class="brand-copy">Securely access your workspace and continue your workflow.</p>
      </section>

      <section class="form-panel">
        <div class="form-header">
          <h2>Sign in</h2>
          <p>Use your account credentials</p>
        </div>

        <form id="login-form" novalidate>
          <label for="email">Email</label>
          <input id="email" name="email" type="email" autocomplete="email" placeholder="you@example.com" required />

          <label for="password">Password</label>
          <div class="password-row">
            <input id="password" name="password" type="password" autocomplete="current-password" placeholder="••••••••" required />
            <button id="toggle-password" class="ghost-btn" type="button" aria-label="Show password">Show</button>
          </div>

          <div class="form-meta">
            <label class="checkbox-row" for="remember">
              <input id="remember" name="remember" type="checkbox" />
              <span>Remember me</span>
            </label>
            <a href="#" aria-label="Forgot password">Forgot password?</a>
          </div>

          <button class="primary-btn" type="submit">Login</button>
        </form>

        <p id="message" class="message">${message}</p>
        <p class="hint">Enter your account credentials to continue.</p>
      </section>
    </main>
  `

  const form = document.querySelector<HTMLFormElement>('#login-form')
  const messageEl = document.querySelector<HTMLParagraphElement>('#message')
  const passwordInput = document.querySelector<HTMLInputElement>('#password')
  const togglePasswordButton = document.querySelector<HTMLButtonElement>('#toggle-password')
  const loginButton = document.querySelector<HTMLButtonElement>('button[type="submit"]')

  if (!form || !messageEl || !passwordInput || !togglePasswordButton || !loginButton) {
    return
  }

  togglePasswordButton.addEventListener('click', () => {
    const currentType = passwordInput.getAttribute('type')
    const isPasswordHidden = currentType === 'password'
    passwordInput.setAttribute('type', isPasswordHidden ? 'text' : 'password')
    togglePasswordButton.textContent = isPasswordHidden ? 'Hide' : 'Show'
  })

  form.addEventListener('submit', async (event) => {
    event.preventDefault()
    const formData = new FormData(form)
    const email = String(formData.get('email') ?? '').trim()
    const password = String(formData.get('password') ?? '').trim()

    if (!email || !password) {
      messageEl.textContent = 'Email and password are required.'
      messageEl.dataset.state = 'error'
      return
    }

    try {
      loginButton.disabled = true
      loginButton.textContent = 'Signing in...'
      messageEl.textContent = ''
      messageEl.dataset.state = ''

      const result = await authenticate(email, password)
      const nextSession: Session = { token: result.token, email: result.user.email, role: result.user.role }
      saveSession(nextSession)
      await loadRegisters(nextSession.token)
      try {
        await loadTestMaster(nextSession.token)
      } catch {
        setTestMaster([], [])
      }
      renderDashboard(nextSession, 'issue-entry', 'replace')
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unable to sign in.'
      messageEl.textContent = errorMessage
      messageEl.dataset.state = 'error'
    } finally {
      loginButton.disabled = false
      loginButton.textContent = 'Login'
    }
  })
}

const renderDashboard = (session: Session, currentModule: ModuleKey, routeMode: 'push' | 'replace' | null = null): void => {
  const menuItems = getMenuItems(session.role)
  const selectedModule = menuItems.includes(currentModule) ? currentModule : menuItems[0]

  activeSession = session
  currentView = selectedModule
  lastKnownDayKey = getTodayLocalDateKey()
  lastKnownOverdueCount = getActivityStats().overdueEntries

  if (routeMode) {
    updateUrlView(selectedModule, routeMode)
  }

  const content =
    selectedModule === 'issue-entry'
      ? renderIssueEntryModule()
      : selectedModule === 'issue-records'
        ? renderIssueRecordsModule()
        : selectedModule === 'drawn-entry'
          ? renderDrawnEntryModule()
          : selectedModule === 'drawn-records'
            ? renderDrawnRecordsModule()
            : renderAdminModule()

  app.innerHTML = `
    <main class="dashboard-shell">
      <aside class="dashboard-sidebar">
        <p class="brand-kicker">LABSOFT</p>
        <h2>Dashboard</h2>
        <p class="sidebar-meta">${escapeHtml(session.email)}</p>
        <p class="sidebar-role">Role: ${session.role}</p>
        <nav class="menu-list">
          ${menuItems
            .map(
              (module) => `
            <button class="menu-btn ${module === selectedModule ? 'active' : ''}" data-module="${module}" type="button">
              ${getModuleLabel(module)}
            </button>
          `
            )
            .join('')}
        </nav>
        <button id="logout-btn" class="secondary-btn" type="button">Logout</button>
      </aside>

      <section class="dashboard-content">
        <div class="module-header">
          <h2>${getModuleLabel(selectedModule)}</h2>
        </div>
        ${renderActivityCards()}
        ${pendingDelete ? `<div class="undo-banner"><span>Entry deleted. Undo available for 8s.</span><button id="undo-delete-btn" class="secondary-btn light" type="button">Undo Delete</button></div>` : ''}
        ${content}
      </section>
    </main>
  `

  const logoutBtn = document.querySelector<HTMLButtonElement>('#logout-btn')
  const menuButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('.menu-btn'))
  const issueForm = document.querySelector<HTMLFormElement>('#issue-form')
  const drawnForm = document.querySelector<HTMLFormElement>('#drawn-form')
  const issueSearchInput = document.querySelector<HTMLInputElement>('#issue-search')
  const issueFromInput = document.querySelector<HTMLInputElement>('#issue-from')
  const issueToInput = document.querySelector<HTMLInputElement>('#issue-to')
  const issueCategoryInput = document.querySelector<HTMLSelectElement>('#issue-filter-category')
  const issueIssuedByInput = document.querySelector<HTMLInputElement>('#issue-filter-issued-by')
  const issueIssuedToInput = document.querySelector<HTMLInputElement>('#issue-filter-issued-to')
  const issueParameterInput = document.querySelector<HTMLInputElement>('#issue-filter-parameter')
  const issueFilterResetButton = document.querySelector<HTMLButtonElement>('#issue-filter-reset')
  const issueExportButton = document.querySelector<HTMLButtonElement>('#issue-export')
  const issueExportPdfButton = document.querySelector<HTMLButtonElement>('#issue-export-pdf')
  const issueCancelEditButton = document.querySelector<HTMLButtonElement>('#issue-cancel-edit')
  const drawnSearchInput = document.querySelector<HTMLInputElement>('#drawn-search')
  const drawnFromInput = document.querySelector<HTMLInputElement>('#drawn-from')
  const drawnToInput = document.querySelector<HTMLInputElement>('#drawn-to')
  const drawnCategoryInput = document.querySelector<HTMLSelectElement>('#drawn-filter-category')
  const drawnByInput = document.querySelector<HTMLInputElement>('#drawn-filter-by')
  const drawnCustomerInput = document.querySelector<HTMLInputElement>('#drawn-filter-customer')
  const drawnParameterInput = document.querySelector<HTMLInputElement>('#drawn-filter-parameter')
  const drawnFilterResetButton = document.querySelector<HTMLButtonElement>('#drawn-filter-reset')
  const drawnExportButton = document.querySelector<HTMLButtonElement>('#drawn-export')
  const drawnExportPdfButton = document.querySelector<HTMLButtonElement>('#drawn-export-pdf')
  const drawnCancelEditButton = document.querySelector<HTMLButtonElement>('#drawn-cancel-edit')
  const undoDeleteButton = document.querySelector<HTMLButtonElement>('#undo-delete-btn')
  const issueEditButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-issue-edit]'))
  const issueDeleteButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-issue-delete]'))
  const drawnEditButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-drawn-edit]'))
  const drawnDeleteButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-drawn-delete]'))
  const adminRefreshButton = document.querySelector<HTMLButtonElement>('#admin-refresh')
  const adminExportIssueButton = document.querySelector<HTMLButtonElement>('#admin-export-issue')
  const adminExportDrawnButton = document.querySelector<HTMLButtonElement>('#admin-export-drawn')
  const adminUserForm = document.querySelector<HTMLFormElement>('#admin-user-form')
  const adminUserToggleButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-admin-toggle]'))
  const adminUserResetButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-admin-reset]'))
  const adminBackupCreateButton = document.querySelector<HTMLButtonElement>('#admin-backup-create')
  const adminBackupRestoreButton = document.querySelector<HTMLButtonElement>('#admin-backup-restore')
  const adminBackupSelect = document.querySelector<HTMLSelectElement>('#admin-backup-select')

  if (!logoutBtn) {
    return
  }

  menuButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      const module = button.dataset.module as ModuleKey | undefined
      if (!module) {
        return
      }

      if (module === 'admin-panel' && session.role === 'admin') {
        try {
          await loadAdminPanelData(session.token)
          adminMessage = ''
          adminMessageState = ''
        } catch (error) {
          adminMessage = error instanceof Error ? error.message : 'Failed to load admin data.'
          adminMessageState = 'error'
        }
      }

      renderDashboard(session, module, 'push')
    })
  })

  issueSearchInput?.addEventListener('input', () => {
    issueSearch = issueSearchInput.value
    renderDashboard(session, 'issue-records')
  })

  issueFromInput?.addEventListener('change', () => {
    issueFromDate = issueFromInput.value
    renderDashboard(session, 'issue-records')
  })

  issueCategoryInput?.addEventListener('change', () => {
    issueCategoryFilter = issueCategoryInput.value || 'All'
    renderDashboard(session, 'issue-records')
  })

  issueToInput?.addEventListener('change', () => {
    issueToDate = issueToInput.value
    renderDashboard(session, 'issue-records')
  })

  issueIssuedByInput?.addEventListener('input', () => {
    issueIssuedByFilter = issueIssuedByInput.value
    renderDashboard(session, 'issue-records')
  })

  issueIssuedToInput?.addEventListener('input', () => {
    issueIssuedToFilter = issueIssuedToInput.value
    renderDashboard(session, 'issue-records')
  })

  issueParameterInput?.addEventListener('input', () => {
    issueParameterFilter = issueParameterInput.value
    renderDashboard(session, 'issue-records')
  })

  issueFilterResetButton?.addEventListener('click', () => {
    resetIssueFilters()
    renderDashboard(session, 'issue-records')
  })

  issueExportButton?.addEventListener('click', () => {
    const rows = getFilteredIssueRecords().map((record) => [
      record.srNo,
      record.codeNo,
      getIssueStatus(record),
      record.sampleDescription,
      record.parameterToBeTested,
      record.issuedOn,
      record.issuedBy,
      record.issuedTo,
      record.reportDueOn,
      record.receivedBy,
      record.reportedOn,
      record.reportedByRemarks
    ])
    downloadCsv('issue-register.csv', ['Sr.No.', 'Code No.', 'Status', 'Sample Description', 'Parameter to be tested', 'Issued On', 'Issued By', 'Issued To', 'Report Due On', 'Received By', 'Reported On', 'ReportedBy/Remarks'], rows)
  })

  issueExportPdfButton?.addEventListener('click', async () => {
    try {
      await downloadIssueRegisterPdf(getFilteredIssueRecords())
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unable to export issue PDF.'
      window.alert(errorMessage)
    }
  })

  issueCancelEditButton?.addEventListener('click', () => {
    issueEditingId = ''
    renderDashboard(session, 'issue-entry')
  })

  issueEditButtons.forEach((button) => {
    button.addEventListener('click', () => {
      issueEditingId = button.dataset.issueEdit ?? ''
      renderDashboard(session, 'issue-entry')
    })
  })

  issueDeleteButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      if (session.role !== 'admin') {
        window.alert('Only admin can delete entries.')
        return
      }

      const recordId = button.dataset.issueDelete ?? ''
      if (!recordId || !window.confirm('Delete this issue entry?')) {
        return
      }

      try {
        const index = issueRecords.findIndex((record) => record.id === recordId)
        if (index >= 0) {
          const removed = issueRecords[index]

          await deleteIssueRecordApi(session.token, recordId)
          issueRecords.splice(index, 1)

          if (pendingDelete) {
            clearTimeout(pendingDelete.timeoutId)
            pendingDelete = null
          }

          pendingDelete = {
            source: 'issue',
            module: 'issue-records',
            index,
            record: removed,
            timeoutId: setTimeout(() => {
              pendingDelete = null
              renderDashboard(session, 'issue-records')
            }, SOFT_DELETE_TIMEOUT_MS)
          }
        }

        if (issueEditingId === recordId) {
          issueEditingId = ''
        }

        renderDashboard(session, 'issue-records')
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unable to delete issue entry.'
        window.alert(errorMessage)
      }
    })
  })

  drawnSearchInput?.addEventListener('input', () => {
    drawnSearch = drawnSearchInput.value
    renderDashboard(session, 'drawn-records')
  })

  drawnFromInput?.addEventListener('change', () => {
    drawnFromDate = drawnFromInput.value
    renderDashboard(session, 'drawn-records')
  })

  drawnToInput?.addEventListener('change', () => {
    drawnToDate = drawnToInput.value
    renderDashboard(session, 'drawn-records')
  })

  drawnCategoryInput?.addEventListener('change', () => {
    drawnCategoryFilter = drawnCategoryInput.value || 'All'
    renderDashboard(session, 'drawn-records')
  })

  drawnByInput?.addEventListener('input', () => {
    drawnByFilter = drawnByInput.value
    renderDashboard(session, 'drawn-records')
  })

  drawnCustomerInput?.addEventListener('input', () => {
    drawnCustomerFilter = drawnCustomerInput.value
    renderDashboard(session, 'drawn-records')
  })

  drawnParameterInput?.addEventListener('input', () => {
    drawnParameterFilter = drawnParameterInput.value
    renderDashboard(session, 'drawn-records')
  })

  drawnFilterResetButton?.addEventListener('click', () => {
    resetDrawnFilters()
    renderDashboard(session, 'drawn-records')
  })

  drawnExportButton?.addEventListener('click', () => {
    const rows = getFilteredDrawnRecords().map((record) => [
      record.srNo,
      getDrawnStatus(record),
      record.sampleDescription,
      record.sampleDrawnOn,
      record.sampleDrawnBy,
      record.customerNameAddress,
      record.parameterToBeTested,
      record.reportDueOn,
      record.sampleReceivedBy
    ])
    downloadCsv('drawn-sample-register.csv', ['Sr.No.', 'Status', 'Sample Description', 'Sample Drawn on', 'Sample Drawn By', 'Customer Name & Address', 'Parameter to be Tested', 'Report Due On', 'Sample Received By'], rows)
  })

  drawnExportPdfButton?.addEventListener('click', async () => {
    try {
      await downloadDrawnRegisterPdf(getFilteredDrawnRecords())
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unable to export receiving PDF.'
      window.alert(errorMessage)
    }
  })

  drawnCancelEditButton?.addEventListener('click', () => {
    drawnEditingId = ''
    renderDashboard(session, 'drawn-entry')
  })

  drawnEditButtons.forEach((button) => {
    button.addEventListener('click', () => {
      drawnEditingId = button.dataset.drawnEdit ?? ''
      renderDashboard(session, 'drawn-entry')
    })
  })

  drawnDeleteButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      if (session.role !== 'admin') {
        window.alert('Only admin can delete entries.')
        return
      }

      const recordId = button.dataset.drawnDelete ?? ''
      if (!recordId || !window.confirm('Delete this drawn sample entry?')) {
        return
      }

      try {
        const index = drawnRecords.findIndex((record) => record.id === recordId)
        if (index >= 0) {
          const removed = drawnRecords[index]

          await deleteDrawnRecordApi(session.token, recordId)
          drawnRecords.splice(index, 1)

          if (pendingDelete) {
            clearTimeout(pendingDelete.timeoutId)
            pendingDelete = null
          }

          pendingDelete = {
            source: 'drawn',
            module: 'drawn-records',
            index,
            record: removed,
            timeoutId: setTimeout(() => {
              pendingDelete = null
              renderDashboard(session, 'drawn-records')
            }, SOFT_DELETE_TIMEOUT_MS)
          }
        }

        if (drawnEditingId === recordId) {
          drawnEditingId = ''
        }

        renderDashboard(session, 'drawn-records')
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unable to delete drawn entry.'
        window.alert(errorMessage)
      }
    })
  })

  adminRefreshButton?.addEventListener('click', async () => {
    try {
      await loadAdminPanelData(session.token)
      await loadRegisters(session.token)
      adminMessage = 'Admin data refreshed successfully.'
      adminMessageState = ''
      renderDashboard(session, 'admin-panel')
    } catch (error) {
      adminMessage = error instanceof Error ? error.message : 'Unable to refresh admin data.'
      adminMessageState = 'error'
      renderDashboard(session, 'admin-panel')
    }
  })

  adminExportIssueButton?.addEventListener('click', () => {
    const rows = issueRecords.map((record) => [
      record.srNo,
      record.codeNo,
      getIssueStatus(record),
      record.sampleDescription,
      record.parameterToBeTested,
      record.issuedOn,
      record.issuedBy,
      record.issuedTo,
      record.reportDueOn,
      record.receivedBy,
      record.reportedOn,
      record.reportedByRemarks
    ])

    downloadCsv('issue-register.csv', ['Sr.No.', 'Code No.', 'Status', 'Sample Description', 'Parameter to be tested', 'Issued On', 'Issued By', 'Issued To', 'Report Due On', 'Received By', 'Reported On', 'ReportedBy/Remarks'], rows)
  })

  adminExportDrawnButton?.addEventListener('click', () => {
    const rows = drawnRecords.map((record) => [
      record.srNo,
      getDrawnStatus(record),
      record.sampleDescription,
      record.sampleDrawnOn,
      record.sampleDrawnBy,
      record.customerNameAddress,
      record.parameterToBeTested,
      record.reportDueOn,
      record.sampleReceivedBy
    ])

    downloadCsv('drawn-sample-register.csv', ['Sr.No.', 'Status', 'Sample Description', 'Sample Drawn on', 'Sample Drawn By', 'Customer Name & Address', 'Parameter to be Tested', 'Report Due On', 'Sample Received By'], rows)
  })

  adminUserForm?.addEventListener('submit', async (event) => {
    event.preventDefault()
    const formData = new FormData(adminUserForm)
    const email = String(formData.get('email') ?? '').trim()
    const password = String(formData.get('password') ?? '').trim()
    const role = (formData.get('role') === 'admin' ? 'admin' : 'staff') as UserRole

    if (!email || !password) {
      adminMessage = 'Email and password are required.'
      adminMessageState = 'error'
      renderDashboard(session, 'admin-panel')
      return
    }

    try {
      await createAdminUser(session.token, { email, password, role })
      await loadAdminPanelData(session.token)
      adminMessage = `User created: ${email}`
      adminMessageState = ''
      renderDashboard(session, 'admin-panel')
    } catch (error) {
      adminMessage = error instanceof Error ? error.message : 'Unable to create user.'
      adminMessageState = 'error'
      renderDashboard(session, 'admin-panel')
    }
  })

  adminUserToggleButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      const userId = button.dataset.adminToggle ?? ''
      const nextState = button.dataset.nextState === 'enable'

      if (!userId) {
        return
      }

      try {
        await updateAdminUserStatus(session.token, userId, nextState)
        await loadAdminPanelData(session.token)
        adminMessage = nextState ? 'User enabled.' : 'User disabled.'
        adminMessageState = ''
        renderDashboard(session, 'admin-panel')
      } catch (error) {
        adminMessage = error instanceof Error ? error.message : 'Unable to update user status.'
        adminMessageState = 'error'
        renderDashboard(session, 'admin-panel')
      }
    })
  })

  adminUserResetButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      const userId = button.dataset.adminReset ?? ''
      if (!userId) {
        return
      }

      const nextPassword = window.prompt('Enter new password (min 8 chars with letter, number, symbol):', '')
      if (!nextPassword) {
        return
      }

      try {
        await resetAdminUserPassword(session.token, userId, nextPassword)
        adminMessage = 'Password reset successful.'
        adminMessageState = ''
        renderDashboard(session, 'admin-panel')
      } catch (error) {
        adminMessage = error instanceof Error ? error.message : 'Unable to reset password.'
        adminMessageState = 'error'
        renderDashboard(session, 'admin-panel')
      }
    })
  })

  adminBackupCreateButton?.addEventListener('click', async () => {
    try {
      const fileName = await createBackup(session.token)
      await loadAdminPanelData(session.token)
      adminMessage = `Backup created: ${fileName}`
      adminMessageState = ''
      renderDashboard(session, 'admin-panel')
    } catch (error) {
      adminMessage = error instanceof Error ? error.message : 'Unable to create backup.'
      adminMessageState = 'error'
      renderDashboard(session, 'admin-panel')
    }
  })

  adminBackupRestoreButton?.addEventListener('click', async () => {
    const fileName = adminBackupSelect?.value ?? ''
    if (!fileName) {
      adminMessage = 'Select a backup first.'
      adminMessageState = 'error'
      renderDashboard(session, 'admin-panel')
      return
    }

    if (!window.confirm(`Restore backup ${fileName}? This will overwrite current data.`)) {
      return
    }

    try {
      await restoreBackup(session.token, fileName)
      await Promise.all([loadRegisters(session.token), loadAdminPanelData(session.token), loadTestMaster(session.token)])
      adminMessage = `Backup restored: ${fileName}`
      adminMessageState = ''
      renderDashboard(session, 'admin-panel')
    } catch (error) {
      adminMessage = error instanceof Error ? error.message : 'Unable to restore backup.'
      adminMessageState = 'error'
      renderDashboard(session, 'admin-panel')
    }
  })

  if (issueForm) {
    initializeIssueAutoUi(issueForm, Boolean(issueEditingId))

    if (!issueEditingId) {
      const syncIssueDraft = (): void => {
        const payload = syncIssueDraftPayload(issueForm)

        saveDraft(ISSUE_DRAFT_KEY, payload)
      }

      issueForm.addEventListener('input', syncIssueDraft)
      issueForm.addEventListener('change', syncIssueDraft)
    }

    issueForm.addEventListener('submit', async (event) => {
      event.preventDefault()
      const formData = new FormData(issueForm)
      const payload = readIssuePayloadFromForm(formData)

      if (
        !payload.srNo ||
        !payload.codeNo ||
        !payload.sampleDescription ||
        !payload.parameterToBeTested ||
        !payload.issuedOn ||
        !payload.issuedBy ||
        !payload.issuedTo ||
        !payload.reportDueOn ||
        !payload.receivedBy
      ) {
        window.alert('Please fill all required fields.')
        return
      }

      if (payload.status === 'Reported' && !payload.reportedOn) {
        window.alert('Reported On is required when status is Reported.')
        return
      }

      if (payload.status !== 'Reported') {
        payload.reportedOn = ''
        payload.reportedByRemarks = payload.reportedByRemarks || ''
      }

      const issueDuplicateMessage = hasIssueDuplicate(payload, issueEditingId)
      if (issueDuplicateMessage) {
        window.alert(issueDuplicateMessage)
        return
      }

      try {
        if (issueEditingId) {
          const updated = await updateIssueRecord(session.token, issueEditingId, payload)
          const index = issueRecords.findIndex((record) => record.id === issueEditingId)
          if (index >= 0) {
            issueRecords[index] = updated
          }
          issueEditingId = ''
          renderDashboard(session, 'issue-records')
        } else {
          const created = await createIssueRecord(session.token, payload)
          issueRecords.unshift(created)
          clearDraft(ISSUE_DRAFT_KEY)
          renderDashboard(session, 'issue-entry')
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unable to save issue entry.'
        window.alert(errorMessage)
      }
    })
  }

  if (drawnForm) {
    initializeDrawnAutoUi(drawnForm, Boolean(drawnEditingId))

    if (!drawnEditingId) {
      const syncDrawnDraft = (): void => {
        const payload = syncDrawnDraftPayload(drawnForm)

        saveDraft(DRAWN_DRAFT_KEY, payload)
      }

      drawnForm.addEventListener('input', syncDrawnDraft)
      drawnForm.addEventListener('change', syncDrawnDraft)
    }

    drawnForm.addEventListener('submit', async (event) => {
      event.preventDefault()
      const formData = new FormData(drawnForm)
      const payload = readDrawnPayloadFromForm(formData)

      if (
        !payload.srNo ||
        !payload.sampleDescription ||
        !payload.sampleDrawnOn ||
        !payload.sampleDrawnBy ||
        !payload.customerNameAddress ||
        !payload.parameterToBeTested ||
        !payload.reportDueOn ||
        !payload.sampleReceivedBy
      ) {
        window.alert('Please fill all required fields.')
        return
      }

      const drawnDuplicateMessage = hasDrawnDuplicate(payload, drawnEditingId)
      if (drawnDuplicateMessage) {
        window.alert(drawnDuplicateMessage)
        return
      }

      try {
        if (drawnEditingId) {
          const updated = await updateDrawnRecord(session.token, drawnEditingId, payload)
          const index = drawnRecords.findIndex((record) => record.id === drawnEditingId)
          if (index >= 0) {
            drawnRecords[index] = updated
          }
          drawnEditingId = ''
          renderDashboard(session, 'drawn-records')
        } else {
          const created = await createDrawnRecord(session.token, payload)
          drawnRecords.unshift(created)
          clearDraft(DRAWN_DRAFT_KEY)
          renderDashboard(session, 'drawn-entry')
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unable to save drawn sample entry.'
        window.alert(errorMessage)
      }
    })
  }

  undoDeleteButton?.addEventListener('click', async () => {
    if (!pendingDelete) {
      return
    }

    const toRestore = pendingDelete
    clearTimeout(toRestore.timeoutId)
    pendingDelete = null

    try {
      if (toRestore.source === 'issue') {
        const restored = await createIssueRecord(session.token, toIssueCreatePayload(toRestore.record as IssueRecord))
        issueRecords.splice(Math.min(toRestore.index, issueRecords.length), 0, restored)
      } else {
        const restored = await createDrawnRecord(session.token, toDrawnCreatePayload(toRestore.record as DrawnRecord))
        drawnRecords.splice(Math.min(toRestore.index, drawnRecords.length), 0, restored)
      }

      renderDashboard(session, toRestore.module)
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unable to restore deleted entry.'
      window.alert(errorMessage)
      renderDashboard(session, toRestore.module)
    }
  })

  logoutBtn.addEventListener('click', () => {
    if (pendingDelete) {
      clearTimeout(pendingDelete.timeoutId)
      pendingDelete = null
    }

    clearSession()
    renderLogin('You have been logged out.')
  })
}

const initApp = async (): Promise<void> => {
  const existingSession = getSession()

  if (!existingSession) {
    renderLogin('', 'replace')
    return
  }

  try {
    const profile = await fetchCurrentUser(existingSession.token)
    const nextSession: Session = {
      token: existingSession.token,
      email: profile.user.email,
      role: profile.user.role
    }

    saveSession(nextSession)
    await loadRegisters(nextSession.token)
    try {
      await loadTestMaster(nextSession.token)
    } catch {
      setTestMaster([], [])
    }
    const initialView = getHashView()

    if (initialView === 'login') {
      renderDashboard(nextSession, 'issue-entry', 'replace')
      return
    }

    if (initialView && initialView !== 'admin-panel') {
      renderDashboard(nextSession, initialView, 'replace')
      return
    }

    if (initialView === 'admin-panel' && nextSession.role === 'admin') {
      try {
        await loadAdminPanelData(nextSession.token)
        adminMessage = ''
        adminMessageState = ''
      } catch (error) {
        adminMessage = error instanceof Error ? error.message : 'Failed to load admin data.'
        adminMessageState = 'error'
      }

      renderDashboard(nextSession, 'admin-panel', 'replace')
      return
    }

    renderDashboard(nextSession, 'issue-entry', 'replace')
  } catch {
    clearSession()
    renderLogin('Session expired. Please login again.', 'replace')
  }
}

const handleBrowserNavigation = (): void => {
  const routeView = getHashView()

  if (!routeView || routeView === currentView) {
    return
  }

  if (routeView === 'login') {
    if (activeSession) {
      renderDashboard(activeSession, 'issue-entry', 'replace')
      return
    }

    renderLogin()
    return
  }

  if (!activeSession) {
    renderLogin('Please login to continue.', 'replace')
    return
  }

  if (routeView === 'admin-panel') {
    if (activeSession.role !== 'admin') {
      renderDashboard(activeSession, 'issue-entry', 'replace')
      return
    }

    void (async () => {
      try {
        await loadAdminPanelData(activeSession.token)
        adminMessage = ''
        adminMessageState = ''
      } catch (error) {
        adminMessage = error instanceof Error ? error.message : 'Failed to load admin data.'
        adminMessageState = 'error'
      }

      renderDashboard(activeSession, 'admin-panel')
    })()

    return
  }

  renderDashboard(activeSession, routeView)
}

const startDailyActivityRefresh = (): void => {
  window.setInterval(() => {
    const todayKey = getTodayLocalDateKey()
    const currentOverdueCount = getActivityStats().overdueEntries
    const hasDayChanged = todayKey !== lastKnownDayKey
    const hasOverdueChanged = currentOverdueCount !== lastKnownOverdueCount

    if (!hasDayChanged && !hasOverdueChanged) {
      return
    }

    lastKnownDayKey = todayKey
    lastKnownOverdueCount = currentOverdueCount

    if (!activeSession || !currentView || currentView === 'login') {
      return
    }

    renderDashboard(activeSession, currentView)
  }, 60 * 1000)
}

window.addEventListener('popstate', handleBrowserNavigation)
window.addEventListener('hashchange', handleBrowserNavigation)
startDailyActivityRefresh()

void initApp()
