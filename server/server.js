import express from 'express'
import bcrypt from 'bcryptjs'
import jwt from 'jsonwebtoken'
import fs from 'node:fs/promises'
import path from 'node:path'
import crypto from 'node:crypto'
import { fileURLToPath } from 'node:url'

const app = express()
const port = process.env.PORT ?? 3001
const jwtSecret = process.env.JWT_SECRET ?? 'dev-only-secret-change-me'
const jwtExpiresIn = '1h'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const clientDistPath = path.join(__dirname, '..', 'dist')
const usersDbPath = path.join(__dirname, 'data', 'users.json')
const registersDbPath = path.join(__dirname, 'data', 'registers.json')
const auditDbPath = path.join(__dirname, 'data', 'audit.json')
const registerHistoryDbPath = path.join(__dirname, 'data', 'register-history.json')
const testMasterDbPath = path.join(__dirname, 'data', 'test-master.json')
const backupsDirPath = path.join(__dirname, 'data', 'backups')

const DEFAULT_USER = {
  email: 'admin@labsoft.dev',
  password: 'Labsoft123',
  role: 'admin'
}

const RECORD_STATUSES = ['Pending', 'In Progress', 'Reported']

const DEFAULT_TEST_MASTER = {
  tests: [
    {
      id: 't1',
      testName: 'Ambient Air Quality Monitoring & Analysis',
      description: 'Ambient Air Quality Monitoring & Analysis',
      displayOrder: 1
    },
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
    { id: 't9', testName: 'Drinking Water Testing', description: 'Drinking Water Testing', displayOrder: 9 },
    { id: 't10', testName: 'Ground Water Quality', description: 'Ground Water Quality', displayOrder: 10 },
    { id: 't11', testName: 'Surface Water Testing', description: 'Surface Water Testing', displayOrder: 11 },
    { id: 't12', testName: 'Soil Quality Test', description: 'Soil Quality Test', displayOrder: 12 }
  ],
  parameters: [
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
      parameterName: 'pH Value, Colour, Odour, Taste, Turbidity, Total Dissolved Solids (TDS), Calcium (as Ca), Chloride (as Cl), Fluoride (as F), Iron (as Fe), Magnesium (as Mg), Total Hardness (as CaCO3), Sulphate',
      displayOrder: 1
    },
    {
      id: 'p10',
      testId: 't10',
      parameterName: 'pH Value, Colour, Odour, Taste, Turbidity, Total Dissolved Solids, Total Hardness (as CaCO3), Calcium (as Ca), Magnesium (as Mg), Chloride (as Cl), Iron (as Fe), Fluoride (as F), Free Residual Chlorine, Phenolic Compound, Anionic Surface Detergents (as MBAS), Sulphate (as SO4), Nitrate (as NO3), Alkalinity (as CaCO3), Copper (as Cu), Total Ammonia, Sulphide (as H2S), Zinc (as Zn), Manganese (as Mn), Boron (as B), Selenium (as Se), Cadmium (as Cd), Lead (as Pb), Total Chromium (as Cr), Nickel (as Ni), Arsenic (as As)',
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
}

const getUserRole = (user) => {
  if (user?.role === 'admin' || user?.role === 'staff' || user?.role === 'customer') {
    return user.role
  }

  if (user?.role === 'customer-care') {
    return 'customer'
  }

  return user?.email === DEFAULT_USER.email ? 'admin' : 'staff'
}

const getRequestRole = (req) => getUserRole({ role: req.user?.role, email: req.user?.email })

const makeRecordId = (prefix) => `${prefix}_${crypto.randomUUID()}`
const makeUserId = () => `u_${crypto.randomUUID()}`
const makeAuditId = () => `aud_${crypto.randomUUID()}`

const isValidEmail = (value) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)

const isStrongPassword = (value) => {
  if (typeof value !== 'string') {
    return false
  }

  const hasMinimumLength = value.length >= 8
  const hasLetter = /[A-Za-z]/.test(value)
  const hasDigit = /\d/.test(value)
  const hasSymbol = /[^A-Za-z0-9]/.test(value)
  return hasMinimumLength && hasLetter && hasDigit && hasSymbol
}

const normalizeRecordStatus = (value, fallback = 'Pending') => {
  if (RECORD_STATUSES.includes(value)) {
    return value
  }

  return RECORD_STATUSES.includes(fallback) ? fallback : 'Pending'
}

const normalizeText = (value) => String(value ?? '').trim().toLowerCase()

const normalizeUserCode = (value) => String(value ?? '').trim()
const normalizeUserName = (value) => String(value ?? '').trim()

const getEmailLabel = (email) => {
  const normalizedEmail = String(email ?? '').trim()
  if (!normalizedEmail.includes('@')) {
    return normalizedEmail
  }

  return normalizedEmail.split('@')[0]
}

const getUserDisplayName = (user) => {
  const explicitName = normalizeUserName(user?.name)
  if (explicitName) {
    return explicitName
  }

  return getEmailLabel(user?.email)
}

const resolveUserDisplayByCode = (users, value) => {
  const normalizedValue = normalizeUserCode(value)
  if (!normalizedValue) {
    return ''
  }

  const user = users.find((entry) => normalizeUserCode(entry.userCode).toLowerCase() === normalizedValue.toLowerCase())
  return user ? getUserDisplayName(user) : normalizedValue
}

const requiresUlrNo = (sampleDescription) => {
  const value = String(sampleDescription ?? '').toLowerCase()
  return value.includes('drinking water') || value.includes('ground water')
}

const isUserCodeTaken = (users, userCode, excludeUserId = '') => {
  const normalizedCode = normalizeUserCode(userCode).toLowerCase()
  if (!normalizedCode) {
    return false
  }

  return users.some(
    (entry) => String(entry.id) !== String(excludeUserId) && normalizeUserCode(entry.userCode).toLowerCase() === normalizedCode
  )
}

const getAuthenticatedUser = async (req) => {
  const users = await readUsers()
  const user = users.find((entry) => String(entry.id) === String(req.user?.sub))

  if (user) {
    return { user, users }
  }

  const fallback = users.find((entry) => entry.email === req.user?.email)
  return { user: fallback ?? null, users }
}

const sanitizeUser = (user) => ({
  id: user.id,
  email: user.email,
  name: getUserDisplayName(user),
  role: getUserRole(user),
  userCode: normalizeUserCode(user.userCode),
  isActive: user.isActive !== false,
  createdAt: user.createdAt
})

const readUsers = async () => {
  try {
    const content = await fs.readFile(usersDbPath, 'utf8')
    const users = JSON.parse(content)
    return Array.isArray(users) ? users : []
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
      return []
    }
    throw error
  }
}

const writeUsers = async (users) => {
  await fs.writeFile(usersDbPath, JSON.stringify(users, null, 2))
}

const readRegisters = async () => {
  try {
    const content = await fs.readFile(registersDbPath, 'utf8')
    const parsed = JSON.parse(content)

    return {
      issueRecords: Array.isArray(parsed.issueRecords) ? parsed.issueRecords : [],
      drawnRecords: Array.isArray(parsed.drawnRecords) ? parsed.drawnRecords : []
    }
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
      return { issueRecords: [], drawnRecords: [] }
    }
    throw error
  }
}

const writeRegisters = async (registers) => {
  await fs.writeFile(registersDbPath, JSON.stringify(registers, null, 2))
}

const readTestMaster = async () => {
  try {
    const content = await fs.readFile(testMasterDbPath, 'utf8')
    const parsed = JSON.parse(content)

    return {
      tests: Array.isArray(parsed.tests) ? parsed.tests : [],
      parameters: Array.isArray(parsed.parameters) ? parsed.parameters : []
    }
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
      return { tests: [], parameters: [] }
    }
    throw error
  }
}

const writeTestMaster = async (payload) => {
  await fs.writeFile(testMasterDbPath, JSON.stringify(payload, null, 2))
}

const readAudit = async () => {
  try {
    const content = await fs.readFile(auditDbPath, 'utf8')
    const entries = JSON.parse(content)
    return Array.isArray(entries) ? entries : []
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
      return []
    }
    throw error
  }
}

const writeAudit = async (entries) => {
  await fs.writeFile(auditDbPath, JSON.stringify(entries, null, 2))
}

const readRegisterHistory = async () => {
  try {
    const content = await fs.readFile(registerHistoryDbPath, 'utf8')
    const entries = JSON.parse(content)
    return Array.isArray(entries) ? entries : []
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
      return []
    }
    throw error
  }
}

const writeRegisterHistory = async (entries) => {
  await fs.writeFile(registerHistoryDbPath, JSON.stringify(entries, null, 2))
}

const readBackupFile = async (fileName) => {
  const normalizedFileName = String(fileName ?? '').trim()
  if (!normalizedFileName || normalizedFileName.includes('/') || normalizedFileName.includes('..')) {
    const error = new Error('Valid backup file name is required.')
    error.statusCode = 400
    throw error
  }

  const filePath = path.join(backupsDirPath, normalizedFileName)

  try {
    const content = await fs.readFile(filePath, 'utf8')
    return JSON.parse(content)
  } catch {
    const error = new Error('Backup file not found.')
    error.statusCode = 404
    throw error
  }
}

const summarizeBackupPayload = (parsed) => ({
  createdAt: String(parsed?.createdAt ?? '').trim(),
  createdBy: String(parsed?.createdBy ?? '').trim(),
  usersCount: Array.isArray(parsed?.users) ? parsed.users.length : 0,
  issueRecordsCount: Array.isArray(parsed?.registers?.issueRecords) ? parsed.registers.issueRecords.length : 0,
  drawnRecordsCount: Array.isArray(parsed?.registers?.drawnRecords) ? parsed.registers.drawnRecords.length : 0,
  auditCount: Array.isArray(parsed?.audit) ? parsed.audit.length : 0,
  registerHistoryCount: Array.isArray(parsed?.registerHistory) ? parsed.registerHistory.length : 0,
  testMasterTestsCount: Array.isArray(parsed?.testMaster?.tests) ? parsed.testMaster.tests.length : 0,
  testMasterParametersCount: Array.isArray(parsed?.testMaster?.parameters) ? parsed.testMaster.parameters.length : 0
})

const createSystemBackup = async (createdBy = 'system') => {
  await ensureBackupsDir()
  const users = await readUsers()
  const registers = await readRegisters()
  const audit = await readAudit()
  const registerHistory = await readRegisterHistory()
  const testMaster = await readTestMaster()

  const fileName = `backup-${new Date().toISOString().replace(/[:.]/g, '-')}.json`
  const filePath = path.join(backupsDirPath, fileName)
  const payload = {
    createdAt: new Date().toISOString(),
    createdBy,
    users,
    registers,
    audit,
    registerHistory,
    testMaster
  }

  await fs.writeFile(filePath, JSON.stringify(payload, null, 2))
  return fileName
}

const appendRegisterHistory = async ({
  action,
  source,
  actor = 'system',
  recordId = '-',
  srNo = '-',
  data = null,
  beforeData = null,
  afterData = null
}) => {
  const entries = await readRegisterHistory()
  entries.push({
    id: `hist_${crypto.randomUUID()}`,
    action,
    source,
    actor,
    recordId,
    srNo,
    data,
    beforeData,
    afterData,
    createdAt: new Date().toISOString()
  })

  await writeRegisterHistory(entries)
}

const appendAudit = async ({ actor = 'system', action, target = '-', details = '', before = null, after = null }) => {
  const entries = await readAudit()
  entries.unshift({
    id: makeAuditId(),
    actor,
    action,
    target,
    details,
    before,
    after,
    createdAt: new Date().toISOString()
  })
  await writeAudit(entries.slice(0, 1000))
}

const normalizeRegisterRecords = (records, prefix) => {
  let changed = false
  const normalized = records.map((record) => {
    const next = { ...record }

    if (!next.id) {
      next.id = makeRecordId(prefix)
      changed = true
    }

    if (!next.createdAt) {
      next.createdAt = new Date().toISOString()
      changed = true
    }

    return next
  })

  return { normalized, changed }
}

const ensureRegisterIds = async () => {
  const registers = await readRegisters()
  const issueResult = normalizeRegisterRecords(registers.issueRecords, 'iss')
  const drawnResult = normalizeRegisterRecords(registers.drawnRecords, 'drw')

  const nextIssueRecords = issueResult.normalized.map((record) => ({
    ...record,
    status: normalizeRecordStatus(record.status, String(record.reportedOn ?? '').trim() ? 'Reported' : 'Pending')
  }))

  const nextDrawnRecords = drawnResult.normalized.map((record) => ({
    ...record,
    status: normalizeRecordStatus(record.status, 'Pending')
  }))

  const statusChanged =
    nextIssueRecords.some((record, index) => record.status !== issueResult.normalized[index].status) ||
    nextDrawnRecords.some((record, index) => record.status !== drawnResult.normalized[index].status)

  if (!issueResult.changed && !drawnResult.changed && !statusChanged) {
    return
  }

  await writeRegisters({
    issueRecords: nextIssueRecords,
    drawnRecords: nextDrawnRecords
  })
}

const resequenceBySrNo = (records) => {
  const sorted = [...records].sort((first, second) => {
    const firstNumeric = Number.parseInt(String(first?.srNo ?? '').match(/\d+/)?.[0] ?? '', 10)
    const secondNumeric = Number.parseInt(String(second?.srNo ?? '').match(/\d+/)?.[0] ?? '', 10)

    if (Number.isFinite(firstNumeric) && Number.isFinite(secondNumeric) && firstNumeric !== secondNumeric) {
      return firstNumeric - secondNumeric
    }

    return String(first?.srNo ?? '').localeCompare(String(second?.srNo ?? ''), undefined, { numeric: true, sensitivity: 'base' })
  })

  const nextSerialById = new Map(sorted.map((record, index) => [String(record.id), String(index + 1)]))
  return records.map((record) => ({
    ...record,
    srNo: nextSerialById.get(String(record.id)) ?? String(record.srNo ?? '')
  }))
}

const ensureAuditFile = async () => {
  const entries = await readAudit()
  await writeAudit(entries)
}

const ensureRegisterHistoryFile = async () => {
  const entries = await readRegisterHistory()
  await writeRegisterHistory(entries)
}

const ensureTestMasterFile = async () => {
  const current = await readTestMaster()
  const hasTests = Array.isArray(current.tests) && current.tests.length > 0
  const hasParameters = Array.isArray(current.parameters) && current.parameters.length > 0

  if (hasTests && hasParameters) {
    return
  }

  await writeTestMaster(DEFAULT_TEST_MASTER)
}

const ensureBackupsDir = async () => {
  await fs.mkdir(backupsDirPath, { recursive: true })
}

const requireFields = (obj, fields) => {
  return fields.every((field) => {
    const value = obj?.[field]
    return typeof value === 'string' && value.trim().length > 0
  })
}

const ensureSeedUser = async () => {
  const users = await readUsers()
  const existing = users.find((user) => user.email === DEFAULT_USER.email)
  if (existing) {
    return
  }

  const passwordHash = await bcrypt.hash(DEFAULT_USER.password, 10)
  users.push({
    id: 'u_admin',
    email: DEFAULT_USER.email,
    role: DEFAULT_USER.role,
    isActive: true,
    passwordHash,
    createdAt: new Date().toISOString()
  })

  await writeUsers(users)
}

const ensureUserDefaults = async () => {
  const users = await readUsers()
  let changed = false

  const normalized = users.map((user) => {
    const next = { ...user }

    if (!next.id) {
      next.id = makeUserId()
      changed = true
    }

    const normalizedRole = getUserRole(next)
    if (!next.role || next.role !== normalizedRole) {
      next.role = normalizedRole
      changed = true
    }

    const normalizedName = normalizeUserName(next.name) || getEmailLabel(next.email)
    if (String(next.name ?? '') !== normalizedName) {
      next.name = normalizedName
      changed = true
    }

    if (typeof next.isActive !== 'boolean') {
      next.isActive = true
      changed = true
    }

    const normalizedUserCode = normalizeUserCode(next.userCode)
    if (String(next.userCode ?? '') !== normalizedUserCode) {
      next.userCode = normalizedUserCode
      changed = true
    }

    if (!next.createdAt) {
      next.createdAt = new Date().toISOString()
      changed = true
    }

    return next
  })

  if (changed) {
    await writeUsers(normalized)
  }
}

const createToken = (payload) => {
  return jwt.sign(payload, jwtSecret, { expiresIn: jwtExpiresIn })
}

const authenticateToken = (req, res, next) => {
  const authorization = req.header('authorization')
  if (!authorization?.startsWith('Bearer ')) {
    return res.status(401).json({ message: 'Missing authorization token.' })
  }

  const token = authorization.replace('Bearer ', '').trim()

  try {
    const decoded = jwt.verify(token, jwtSecret)
    req.user = decoded
    return next()
  } catch {
    return res.status(401).json({ message: 'Invalid or expired token.' })
  }
}

const requireAdmin = (req, res, next) => {
  if (getRequestRole(req) !== 'admin') {
    return res.status(403).json({ message: 'Admin access required.' })
  }

  return next()
}

const requireDrawnAccess = (req, res, next) => {
  const role = getRequestRole(req)
  if (role === 'staff') {
    return res.status(403).json({ message: 'Staff can access only Issue Register modules.' })
  }

  return next()
}

const requireDrawnManageAccess = (req, res, next) => {
  const role = getRequestRole(req)
  if (role !== 'admin' && role !== 'customer') {
    return res.status(403).json({ message: 'Only admin and customer care can modify or delete receiving records.' })
  }

  return next()
}

const requireIssueEntryAccess = (req, res, next) => {
  const role = getRequestRole(req)
  if (role !== 'admin' && role !== 'staff' && role !== 'customer') {
    return res.status(403).json({ message: 'Only admin, staff, and customer care can create issue records.' })
  }

  return next()
}

const requireIssueManageAccess = (req, res, next) => {
  const role = getRequestRole(req)
  if (role !== 'admin' && role !== 'staff' && role !== 'customer') {
    return res.status(403).json({ message: 'Only admin, staff, and customer care can modify or delete issue records.' })
  }

  return next()
}

const allowedSampleCategories = ['air', 'water', 'soil', 'noise']

const getSampleCategory = (sampleDescription) => {
  const normalized = String(sampleDescription ?? '').trim().toLowerCase()
  return allowedSampleCategories.find((category) => normalized.includes(category)) ?? ''
}

const isValidDateInput = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value ?? '').trim())

const buildValidationError = (field, message) => ({ field, message })

const validateIssueRecordPayload = (record, existingRecords = [], excludeId = '') => {
  const errors = []
  const requiredFields = [
    ['srNo', 'Sr.No. is required.'],
    ['codeNo', 'Code No. is required.'],
    ['sampleDescription', 'Sample description is required.'],
    ['parameterToBeTested', 'Parameter to be tested is required.'],
    ['issuedOn', 'Issued On date is required.'],
    ['issuedBy', 'Issued By is required.'],
    ['issuedTo', 'Issued To is required.'],
    ['reportDueOn', 'Report Due On date is required.'],
    ['receivedBy', 'Received By is required.']
  ]

  requiredFields.forEach(([field, message]) => {
    if (typeof record?.[field] !== 'string' || !String(record[field]).trim()) {
      errors.push(buildValidationError(field, message))
    }
  })

  const sampleCategory = getSampleCategory(record?.sampleDescription)
  if (!sampleCategory) {
    errors.push(buildValidationError('sampleDescription', 'Sample category must include Air, Water, Soil, or Noise.'))
  }

  if (record?.issuedOn && !isValidDateInput(record.issuedOn)) {
    errors.push(buildValidationError('issuedOn', 'Issued On must be a valid date.'))
  }

  if (record?.reportDueOn && !isValidDateInput(record.reportDueOn)) {
    errors.push(buildValidationError('reportDueOn', 'Report Due On must be a valid date.'))
  }

  if (record?.reportedOn && String(record.reportedOn).trim() && !isValidDateInput(record.reportedOn)) {
    errors.push(buildValidationError('reportedOn', 'Reported On must be a valid date.'))
  }

  const issuedOnTime = Date.parse(String(record?.issuedOn ?? ''))
  const dueOnTime = Date.parse(String(record?.reportDueOn ?? ''))
  const reportedOnTime = Date.parse(String(record?.reportedOn ?? ''))
  if (Number.isFinite(issuedOnTime) && Number.isFinite(dueOnTime) && dueOnTime < issuedOnTime) {
    errors.push(buildValidationError('reportDueOn', 'Report Due On cannot be earlier than Issued On.'))
  }

  if (String(record?.reportedOn ?? '').trim() && Number.isFinite(issuedOnTime) && Number.isFinite(reportedOnTime) && reportedOnTime < issuedOnTime) {
    errors.push(buildValidationError('reportedOn', 'Reported On cannot be earlier than Issued On.'))
  }

  const srNo = normalizeText(record?.srNo)
  const codeNo = normalizeText(record?.codeNo)
  const ulrNo = normalizeText(record?.ulrNo)
  const duplicate = existingRecords.find(
    (entry) =>
      String(entry.id) !== String(excludeId) &&
      (normalizeText(entry.srNo) === srNo ||
        normalizeText(entry.codeNo) === codeNo ||
        (ulrNo && normalizeText(entry.ulrNo) === ulrNo))
  )

  if (duplicate) {
    if (normalizeText(duplicate.srNo) === srNo) {
      errors.push(buildValidationError('srNo', `Duplicate Sr.No. found: ${record.srNo}`))
    }
    if (normalizeText(duplicate.codeNo) === codeNo) {
      errors.push(buildValidationError('codeNo', `Duplicate Code No. found: ${record.codeNo}`))
    }
    if (ulrNo && normalizeText(duplicate.ulrNo) === ulrNo) {
      errors.push(buildValidationError('ulrNo', `Duplicate ULR No. found: ${record.ulrNo}`))
    }
  }

  return errors
}

const validateDrawnRecordPayload = (record, existingRecords = [], excludeId = '') => {
  const errors = []
  const requiredFields = [
    ['srNo', 'Sr.No. is required.'],
    ['reportCode', 'Report Code is required.'],
    ['sampleDescription', 'Sample description is required.'],
    ['sampleDrawnOn', 'Sample Drawn On date is required.'],
    ['sampleDrawnBy', 'Sample Drawn By is required.'],
    ['customerNameAddress', 'Customer Name & Address is required.'],
    ['parameterToBeTested', 'Parameter to be tested is required.'],
    ['reportDueOn', 'Report Due On date is required.'],
    ['sampleReceivedBy', 'Sample Received By is required.']
  ]

  requiredFields.forEach(([field, message]) => {
    if (typeof record?.[field] !== 'string' || !String(record[field]).trim()) {
      errors.push(buildValidationError(field, message))
    }
  })

  const sampleCategory = getSampleCategory(record?.sampleDescription)
  if (!sampleCategory) {
    errors.push(buildValidationError('sampleDescription', 'Sample category must include Air, Water, Soil, or Noise.'))
  }

  if (record?.sampleDrawnOn && !isValidDateInput(record.sampleDrawnOn)) {
    errors.push(buildValidationError('sampleDrawnOn', 'Sample Drawn On must be a valid date.'))
  }

  if (record?.reportDueOn && !isValidDateInput(record.reportDueOn)) {
    errors.push(buildValidationError('reportDueOn', 'Report Due On must be a valid date.'))
  }

  const drawnOnTime = Date.parse(String(record?.sampleDrawnOn ?? ''))
  const dueOnTime = Date.parse(String(record?.reportDueOn ?? ''))
  if (Number.isFinite(drawnOnTime) && Number.isFinite(dueOnTime) && dueOnTime < drawnOnTime) {
    errors.push(buildValidationError('reportDueOn', 'Report Due On cannot be earlier than Sample Drawn On.'))
  }

  const srNo = normalizeText(record?.srNo)
  const reportCode = normalizeText(record?.reportCode)
  const ulrNo = normalizeText(record?.ulrNo)
  const duplicate = existingRecords.find(
    (entry) =>
      String(entry.id) !== String(excludeId) &&
      (normalizeText(entry.srNo) === srNo ||
        normalizeText(entry.reportCode) === reportCode ||
        (ulrNo && normalizeText(entry.ulrNo) === ulrNo))
  )

  if (duplicate) {
    if (normalizeText(duplicate.srNo) === srNo) {
      errors.push(buildValidationError('srNo', `Duplicate Sr.No. found: ${record.srNo}`))
    }
    if (normalizeText(duplicate.reportCode) === reportCode) {
      errors.push(buildValidationError('reportCode', `Duplicate Report Code found: ${record.reportCode}`))
    }
    if (ulrNo && normalizeText(duplicate.ulrNo) === ulrNo) {
      errors.push(buildValidationError('ulrNo', `Duplicate ULR No. found: ${record.ulrNo}`))
    }
  }

  return errors
}

app.use(express.json())

app.use('/api', (_req, res, next) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate')
  res.setHeader('Pragma', 'no-cache')
  res.setHeader('Expires', '0')
  next()
})

app.get('/api/health', (_req, res) => {
  res.json({ ok: true })
})

app.post('/api/login', async (req, res) => {
  const email = String(req.body?.email ?? '').trim().toLowerCase()
  const password = String(req.body?.password ?? '').trim()

  if (!email || !password) {
    return res.status(400).json({ message: 'Email and password are required.' })
  }

  const users = await readUsers()
  const user = users.find((entry) => entry.email.toLowerCase() === email)

  if (!user || !user.passwordHash) {
    return res.status(401).json({ message: 'Invalid credentials.' })
  }

  if (user.isActive === false) {
    return res.status(403).json({ message: 'Account is disabled. Contact admin.' })
  }

  const isPasswordValid = await bcrypt.compare(password, user.passwordHash)
  if (!isPasswordValid) {
    return res.status(401).json({ message: 'Invalid credentials.' })
  }

  const role = getUserRole(user)
  const token = createToken({ sub: user.id, email: user.email, role })

  await appendAudit({ actor: user.email, action: 'LOGIN_SUCCESS', target: user.email, details: `Role: ${role}` })

  return res.json({
    token,
    user: { email: user.email, name: getUserDisplayName(user), role, userCode: normalizeUserCode(user.userCode) }
  })
})

app.get('/api/me', authenticateToken, async (req, res) => {
  const { user } = await getAuthenticatedUser(req)
  const role = user ? getUserRole(user) : getUserRole({ role: req.user.role, email: req.user.email })
  return res.json({
    user: {
      email: req.user.email,
      name: getUserDisplayName(user ?? req.user),
      role,
      userCode: normalizeUserCode(user?.userCode)
    }
  })
})

app.get('/api/user-directory', authenticateToken, async (_req, res) => {
  const users = await readUsers()
  return res.json({
    users: users
      .filter((user) => user.isActive !== false)
      .map((user) => ({
        id: user.id,
        name: getUserDisplayName(user),
        userCode: normalizeUserCode(user.userCode)
      }))
  })
})

app.get('/api/registers', authenticateToken, async (_req, res) => {
  const req = _req
  const registers = await readRegisters()
  const role = getRequestRole(req)
  const users = await readUsers()

  const issueRecords = registers.issueRecords.map((record) => ({
    ...record,
    receivedByName: resolveUserDisplayByCode(users, record.receivedBy)
  }))

  const drawnRecords = registers.drawnRecords.map((record) => ({
    ...record,
    sampleReceivedByName: resolveUserDisplayByCode(users, record.sampleReceivedBy)
  }))

  if (role === 'staff') {
    return res.json({
      issueRecords,
      drawnRecords: []
    })
  }

  return res.json({
    issueRecords,
    drawnRecords
  })
})

app.get('/api/test-master', authenticateToken, async (_req, res) => {
  const testMaster = await readTestMaster()

  const tests = [...testMaster.tests].sort((first, second) => Number(first.displayOrder ?? 0) - Number(second.displayOrder ?? 0))
  const parameters = [...testMaster.parameters].sort(
    (first, second) => Number(first.displayOrder ?? 0) - Number(second.displayOrder ?? 0)
  )

  return res.json({ tests, parameters })
})

app.get('/api/staff/receiving-by-report-code/:reportCode', authenticateToken, async (req, res) => {

  const reportCode = String(req.params?.reportCode ?? '').trim()
  if (!reportCode) {
    return res.status(400).json({ message: 'Report Code is required.' })
  }

  const registers = await readRegisters()
  const normalize = (v) => String(v ?? '').replace(/\s+/g, '').toLowerCase()
  const normalizedInput = normalize(reportCode)
  const debugList = registers.drawnRecords.map(entry => ({ raw: entry.reportCode, normalized: normalize(entry.reportCode) }))
  console.log('Receiving lookup:', { input: reportCode, normalizedInput, debugList })
  const match = registers.drawnRecords.find((entry) => normalize(entry.reportCode) === normalizedInput)

  if (!match) {
    return res.status(404).json({ message: 'No receiving entry found for this Report Code.' })
  }

  return res.json({
    record: {
      srNo: String(match.srNo ?? '').trim(),
      reportCode: String(match.reportCode ?? '').trim(),
      ulrNo: String(match.ulrNo ?? '').trim(),
      sampleDescription: String(match.sampleDescription ?? '').trim(),
      parameterToBeTested: String(match.parameterToBeTested ?? '').trim(),
      issuedOn: String(match.sampleDrawnOn ?? '').trim(),
      issuedBy: String(match.sampleDrawnBy ?? '').trim(),
      issuedTo: String(match.sampleReceivedBy ?? '').trim(),
      reportDueOn: String(match.reportDueOn ?? '').trim()
    }
  })
})

const findReceivingByReportCode = (registers, reportCode) => {
  const normalizedInput = String(reportCode ?? '').replace(/\s+/g, '').toLowerCase()
  return registers.drawnRecords.find((entry) => String(entry.reportCode ?? '').replace(/\s+/g, '').toLowerCase() === normalizedInput) ?? null
}

app.post('/api/registers/issue', authenticateToken, requireIssueEntryAccess, async (req, res) => {
  const record = req.body
  const requiredFields = [
    'srNo',
    'codeNo',
    'sampleDescription',
    'parameterToBeTested',
    'issuedOn',
    'issuedBy',
    'issuedTo',
    'reportDueOn',
    'receivedBy'
  ]

  const { user: actor } = await getAuthenticatedUser(req)
  const actorRole = actor ? getUserRole(actor) : getRequestRole(req)
  const assignedUserCode = normalizeUserCode(actor?.userCode)
  if (actorRole === 'staff' && !assignedUserCode) {
    return res.status(400).json({ message: 'Unique number is not assigned to your account. Contact admin.' })
  }

  const registers = await readRegisters()
  const receivingSource = actorRole === 'staff' ? findReceivingByReportCode(registers, record.codeNo) : null
  if (actorRole === 'staff' && !receivingSource) {
    return res.status(404).json({ message: 'No receiving entry found for this Report Code.' })
  }

  const effectiveRecord = actorRole !== 'staff'
    ? record
    : {
        ...record,
        srNo: String(receivingSource?.srNo ?? '').trim(),
        codeNo: String(receivingSource?.reportCode ?? '').trim(),
        sampleDescription: String(receivingSource?.sampleDescription ?? '').trim(),
        parameterToBeTested: String(receivingSource?.parameterToBeTested ?? '').trim(),
        ulrNo: String(receivingSource?.ulrNo ?? '').trim()
      }

  const ulrNo = String(effectiveRecord.ulrNo ?? '').trim()
  const needsUlrNo = requiresUlrNo(effectiveRecord.sampleDescription)
  if (needsUlrNo && !ulrNo) {
    return res.status(400).json({ message: 'ULR No. is required for Drinking Water or Ground Water samples.', field: 'ulrNo' })
  }

  const validationErrors = validateIssueRecordPayload(effectiveRecord, registers.issueRecords)
  if (validationErrors.length > 0) {
    return res.status(400).json({ message: validationErrors[0].message, field: validationErrors[0].field, errors: validationErrors })
  }

  const normalizedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(effectiveRecord[field]).trim()])),
    ulrNo: needsUlrNo ? ulrNo : '',
    receivedBy: actorRole === 'staff' ? assignedUserCode : String(effectiveRecord.receivedBy).trim(),
    status: normalizeRecordStatus(effectiveRecord.status, String(effectiveRecord.reportedOn ?? '').trim() ? 'Reported' : 'Pending'),
    reportedOn: String(effectiveRecord.reportedOn ?? '').trim(),
    reportedByRemarks: String(effectiveRecord.reportedByRemarks ?? '').trim()
  }

  if (normalizedRecord.status !== 'Reported') {
    normalizedRecord.reportedOn = ''
  }

  normalizedRecord.id = makeRecordId('iss')
  normalizedRecord.createdAt = new Date().toISOString()

  registers.issueRecords.unshift(normalizedRecord)
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'CREATE',
    source: 'issue',
    actor: req.user?.email,
    recordId: normalizedRecord.id,
    srNo: normalizedRecord.srNo,
    data: normalizedRecord,
    afterData: normalizedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'ISSUE_CREATE',
    target: normalizedRecord.id,
    details: `SrNo: ${normalizedRecord.srNo}`,
    after: normalizedRecord
  })

  return res.status(201).json({ record: normalizedRecord })
})

app.put('/api/registers/issue/:id', authenticateToken, requireIssueManageAccess, async (req, res) => {
  const { id } = req.params
  const record = req.body
  const requiredFields = [
    'srNo',
    'codeNo',
    'sampleDescription',
    'parameterToBeTested',
    'issuedOn',
    'issuedBy',
    'issuedTo',
    'reportDueOn',
    'receivedBy'
  ]

  const { user: actor } = await getAuthenticatedUser(req)
  const actorRole = actor ? getUserRole(actor) : getRequestRole(req)
  const assignedUserCode = normalizeUserCode(actor?.userCode)
  if (actorRole === 'staff' && !assignedUserCode) {
    return res.status(400).json({ message: 'Unique number is not assigned to your account. Contact admin.' })
  }

  const registers = await readRegisters()
  const index = registers.issueRecords.findIndex((entry) => String(entry.id) === String(id))
  if (index === -1) {
    return res.status(404).json({ message: 'Issue record not found.' })
  }

  const previousRecord = registers.issueRecords[index]
  const effectiveRecord = actorRole !== 'staff'
    ? record
    : {
        ...record,
        srNo: String(previousRecord.srNo ?? '').trim(),
        codeNo: String(previousRecord.codeNo ?? '').trim(),
        sampleDescription: String(previousRecord.sampleDescription ?? '').trim(),
        parameterToBeTested: String(previousRecord.parameterToBeTested ?? '').trim(),
        ulrNo: String(previousRecord.ulrNo ?? '').trim()
      }

  const ulrNo = String(effectiveRecord.ulrNo ?? '').trim()
  const needsUlrNo = requiresUlrNo(effectiveRecord.sampleDescription)
  if (needsUlrNo && !ulrNo) {
    return res.status(400).json({ message: 'ULR No. is required for Drinking Water or Ground Water samples.', field: 'ulrNo' })
  }

  const validationErrors = validateIssueRecordPayload(effectiveRecord, registers.issueRecords, id)
  if (validationErrors.length > 0) {
    return res.status(400).json({ message: validationErrors[0].message, field: validationErrors[0].field, errors: validationErrors })
  }

  const updatedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(effectiveRecord[field]).trim()])),
    ulrNo: needsUlrNo ? ulrNo : '',
    receivedBy: actorRole === 'staff' ? assignedUserCode : String(effectiveRecord.receivedBy).trim(),
    status: normalizeRecordStatus(effectiveRecord.status, String(effectiveRecord.reportedOn ?? '').trim() ? 'Reported' : 'Pending'),
    reportedOn: String(effectiveRecord.reportedOn ?? '').trim(),
    reportedByRemarks: String(effectiveRecord.reportedByRemarks ?? '').trim(),
    id: registers.issueRecords[index].id,
    createdAt: registers.issueRecords[index].createdAt ?? new Date().toISOString()
  }

  if (updatedRecord.status !== 'Reported') {
    updatedRecord.reportedOn = ''
  }

  registers.issueRecords[index] = updatedRecord
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'UPDATE',
    source: 'issue',
    actor: req.user?.email,
    recordId: updatedRecord.id,
    srNo: updatedRecord.srNo,
    data: updatedRecord,
    beforeData: previousRecord,
    afterData: updatedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'ISSUE_UPDATE',
    target: updatedRecord.id,
    details: `SrNo: ${updatedRecord.srNo}`,
    before: previousRecord,
    after: updatedRecord
  })

  return res.json({ record: updatedRecord })
})

app.delete('/api/registers/issue/:id', authenticateToken, requireIssueManageAccess, async (req, res) => {
  const { id } = req.params
  const registers = await readRegisters()
  const deletedRecord = registers.issueRecords.find((entry) => String(entry.id) === String(id))
  const nextRecords = registers.issueRecords.filter((entry) => String(entry.id) !== String(id))

  if (nextRecords.length === registers.issueRecords.length) {
    return res.status(404).json({ message: 'Issue record not found.' })
  }

  registers.issueRecords = resequenceBySrNo(nextRecords)
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'DELETE',
    source: 'issue',
    actor: req.user?.email,
    recordId: id,
    srNo: deletedRecord?.srNo ?? '-',
    data: deletedRecord ?? null,
    beforeData: deletedRecord ?? null
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'ISSUE_DELETE',
    target: id,
    details: 'Issue record deleted',
    before: deletedRecord ?? null
  })

  return res.status(204).send()
})

app.post('/api/registers/drawn', authenticateToken, requireDrawnAccess, async (req, res) => {
  const record = req.body
  const requiredFields = [
    'srNo',
    'reportCode',
    'sampleDescription',
    'sampleDrawnOn',
    'sampleDrawnBy',
    'customerNameAddress',
    'parameterToBeTested',
    'reportDueOn',
    'sampleReceivedBy'
  ]

  const ulrNo = String(record.ulrNo ?? '').trim()
  const needsUlrNo = requiresUlrNo(record.sampleDescription)
  if (needsUlrNo && !ulrNo) {
    return res.status(400).json({ message: 'ULR No. is required for Drinking Water or Ground Water samples.', field: 'ulrNo' })
  }

  const { user: actor } = await getAuthenticatedUser(req)
  const actorRole = actor ? getUserRole(actor) : getRequestRole(req)

  const registers = await readRegisters()
  const validationErrors = validateDrawnRecordPayload(record, registers.drawnRecords)
  if (validationErrors.length > 0) {
    return res.status(400).json({ message: validationErrors[0].message, field: validationErrors[0].field, errors: validationErrors })
  }

  const normalizedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(record[field]).trim()])),
    ulrNo: needsUlrNo ? ulrNo : '',
    sampleReceivedBy: String(record.sampleReceivedBy).trim(),
    status: normalizeRecordStatus(record.status, 'Pending')
  }
  normalizedRecord.id = makeRecordId('drw')
  normalizedRecord.createdAt = new Date().toISOString()

  registers.drawnRecords.unshift(normalizedRecord)
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'CREATE',
    source: 'drawn',
    actor: req.user?.email,
    recordId: normalizedRecord.id,
    srNo: normalizedRecord.srNo,
    data: normalizedRecord,
    afterData: normalizedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'DRAWN_CREATE',
    target: normalizedRecord.id,
    details: `SrNo: ${normalizedRecord.srNo}`,
    after: normalizedRecord
  })

  return res.status(201).json({ record: normalizedRecord })
})

app.put('/api/registers/drawn/:id', authenticateToken, requireDrawnManageAccess, async (req, res) => {
  const { id } = req.params
  const record = req.body
  const requiredFields = [
    'srNo',
    'reportCode',
    'sampleDescription',
    'sampleDrawnOn',
    'sampleDrawnBy',
    'customerNameAddress',
    'parameterToBeTested',
    'reportDueOn',
    'sampleReceivedBy'
  ]

  const ulrNo = String(record.ulrNo ?? '').trim()
  const needsUlrNo = requiresUlrNo(record.sampleDescription)
  if (needsUlrNo && !ulrNo) {
    return res.status(400).json({ message: 'ULR No. is required for Drinking Water or Ground Water samples.', field: 'ulrNo' })
  }

  const { user: actor } = await getAuthenticatedUser(req)
  const actorRole = actor ? getUserRole(actor) : getRequestRole(req)

  const registers = await readRegisters()
  const index = registers.drawnRecords.findIndex((entry) => String(entry.id) === String(id))
  if (index === -1) {
    return res.status(404).json({ message: 'Drawn record not found.' })
  }

  const validationErrors = validateDrawnRecordPayload(record, registers.drawnRecords, id)
  if (validationErrors.length > 0) {
    return res.status(400).json({ message: validationErrors[0].message, field: validationErrors[0].field, errors: validationErrors })
  }

  const previousRecord = registers.drawnRecords[index]
  const updatedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(record[field]).trim()])),
    ulrNo: needsUlrNo ? ulrNo : '',
    sampleReceivedBy: String(record.sampleReceivedBy).trim(),
    status: normalizeRecordStatus(record.status, 'Pending'),
    id: registers.drawnRecords[index].id,
    createdAt: registers.drawnRecords[index].createdAt ?? new Date().toISOString()
  }

  registers.drawnRecords[index] = updatedRecord
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'UPDATE',
    source: 'drawn',
    actor: req.user?.email,
    recordId: updatedRecord.id,
    srNo: updatedRecord.srNo,
    data: updatedRecord,
    beforeData: previousRecord,
    afterData: updatedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'DRAWN_UPDATE',
    target: updatedRecord.id,
    details: `SrNo: ${updatedRecord.srNo}`,
    before: previousRecord,
    after: updatedRecord
  })

  return res.json({ record: updatedRecord })
})

app.delete('/api/registers/drawn/:id', authenticateToken, requireDrawnManageAccess, async (req, res) => {
  const { id } = req.params
  const registers = await readRegisters()
  const deletedRecord = registers.drawnRecords.find((entry) => String(entry.id) === String(id))
  const nextRecords = registers.drawnRecords.filter((entry) => String(entry.id) !== String(id))

  if (nextRecords.length === registers.drawnRecords.length) {
    return res.status(404).json({ message: 'Drawn record not found.' })
  }

  registers.drawnRecords = resequenceBySrNo(nextRecords)
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'DELETE',
    source: 'drawn',
    actor: req.user?.email,
    recordId: id,
    srNo: deletedRecord?.srNo ?? '-',
    data: deletedRecord ?? null,
    beforeData: deletedRecord ?? null
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'DRAWN_DELETE',
    target: id,
    details: 'Drawn record deleted',
    before: deletedRecord ?? null
  })

  return res.status(204).send()
})

app.use('/api/admin', authenticateToken, requireAdmin)

app.get('/api/admin/users', async (_req, res) => {
  const users = await readUsers()
  return res.json({ users: users.map(sanitizeUser) })
})

app.post('/api/admin/users', async (req, res) => {
  const email = String(req.body?.email ?? '').trim().toLowerCase()
  const name = normalizeUserName(req.body?.name) || getEmailLabel(email)
  const password = String(req.body?.password ?? '').trim()
  const roleInput = String(req.body?.role ?? '').trim().toLowerCase()
  const role = roleInput === 'admin' ? 'admin' : roleInput === 'customer' ? 'customer' : 'staff'
  const userCode = normalizeUserCode(req.body?.userCode)

  if (!isValidEmail(email)) {
    return res.status(400).json({ message: 'Valid email is required.' })
  }

  if (!isStrongPassword(password)) {
    return res.status(400).json({ message: 'Password must be at least 8 chars and include letter, number, and symbol.' })
  }

  if (role !== 'admin' && !userCode) {
    return res.status(400).json({ message: 'Unique number is required for Staff and Customer Care users.' })
  }

  const users = await readUsers()
  if (users.some((entry) => entry.email.toLowerCase() === email)) {
    return res.status(409).json({ message: 'User already exists.' })
  }

  if (isUserCodeTaken(users, userCode)) {
    return res.status(409).json({ message: 'Unique number is already assigned to another user.' })
  }

  const passwordHash = await bcrypt.hash(password, 10)
  const user = {
    id: makeUserId(),
    email,
    name,
    role,
    userCode,
    isActive: true,
    passwordHash,
    createdAt: new Date().toISOString()
  }

  users.unshift(user)
  await writeUsers(users)
  await appendAudit({ actor: req.user?.email, action: 'USER_CREATE', target: user.email, details: `Role: ${role}` })

  return res.status(201).json({ user: sanitizeUser(user) })
})

app.patch('/api/admin/users/:id/status', async (req, res) => {
  const { id } = req.params
  const isActive = req.body?.isActive === true
  const users = await readUsers()
  const userIndex = users.findIndex((entry) => String(entry.id) === String(id))

  if (userIndex === -1) {
    return res.status(404).json({ message: 'User not found.' })
  }

  if (users[userIndex].email === DEFAULT_USER.email && !isActive) {
    return res.status(400).json({ message: 'Default admin account cannot be disabled.' })
  }

  users[userIndex] = {
    ...users[userIndex],
    isActive
  }

  await writeUsers(users)
  await appendAudit({
    actor: req.user?.email,
    action: 'USER_STATUS_UPDATE',
    target: users[userIndex].email,
    details: isActive ? 'Enabled' : 'Disabled'
  })

  return res.json({ user: sanitizeUser(users[userIndex]) })
})

app.post('/api/admin/users/:id/reset-password', async (req, res) => {
  const { id } = req.params
  const password = String(req.body?.password ?? '').trim()
  const users = await readUsers()
  const userIndex = users.findIndex((entry) => String(entry.id) === String(id))

  if (userIndex === -1) {
    return res.status(404).json({ message: 'User not found.' })
  }

  if (!isStrongPassword(password)) {
    return res.status(400).json({ message: 'Password must be at least 8 chars and include letter, number, and symbol.' })
  }

  users[userIndex] = {
    ...users[userIndex],
    passwordHash: await bcrypt.hash(password, 10)
  }

  await writeUsers(users)
  await appendAudit({
    actor: req.user?.email,
    action: 'USER_PASSWORD_RESET',
    target: users[userIndex].email,
    details: 'Password reset by admin'
  })

  return res.status(204).send()
})

app.delete('/api/admin/users/:id', async (req, res) => {
  const { id } = req.params
  const users = await readUsers()
  const userIndex = users.findIndex((entry) => String(entry.id) === String(id))

  if (userIndex === -1) {
    return res.status(404).json({ message: 'User not found.' })
  }

  const targetUser = users[userIndex]
  if (targetUser.email === DEFAULT_USER.email) {
    return res.status(400).json({ message: 'Default admin account cannot be deleted.' })
  }

  if (String(targetUser.id) === String(req.user?.sub)) {
    return res.status(400).json({ message: 'You cannot delete your own active account.' })
  }

  users.splice(userIndex, 1)
  await writeUsers(users)
  await appendAudit({
    actor: req.user?.email,
    action: 'USER_DELETE',
    target: targetUser.email,
    details: `Role: ${getUserRole(targetUser)}`,
    before: targetUser
  })

  return res.status(204).send()
})

app.get('/api/admin/audit', async (req, res) => {
  const limitRaw = Number(req.query?.limit ?? 50)
  const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(200, Math.floor(limitRaw))) : 50
  const entries = await readAudit()
  return res.json({ entries: entries.slice(0, limit) })
})

app.get('/api/admin/register-history', async (req, res) => {
  const limitRaw = Number(req.query?.limit ?? 50)
  const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(500, Math.floor(limitRaw))) : 50
  const entries = await readRegisterHistory()
  const sorted = [...entries].sort((first, second) => Date.parse(second.createdAt ?? '') - Date.parse(first.createdAt ?? ''))
  return res.json({ entries: sorted.slice(0, limit) })
})

app.get('/api/admin/alerts', async (_req, res) => {
  const registers = await readRegisters()
  const now = new Date()
  const maxDueDate = new Date(now)
  maxDueDate.setDate(now.getDate() + 2)

  const alerts = []

  registers.issueRecords.forEach((record) => {
    const status = normalizeRecordStatus(record.status, String(record.reportedOn ?? '').trim().length > 0 ? 'Reported' : 'Pending')
    if (status === 'Reported') {
      return
    }

    const due = new Date(record.reportDueOn)
    if (Number.isNaN(due.getTime())) {
      return
    }

    if (due < now) {
      alerts.push({
        type: 'overdue',
        source: 'issue',
        recordId: record.id,
        srNo: record.srNo,
        dueOn: record.reportDueOn,
        message: `Issue report overdue for Sr.No. ${record.srNo}`
      })
      return
    }

    if (due <= maxDueDate) {
      alerts.push({
        type: 'due-soon',
        source: 'issue',
        recordId: record.id,
        srNo: record.srNo,
        dueOn: record.reportDueOn,
        message: `Issue report due soon for Sr.No. ${record.srNo}`
      })
    }
  })

  registers.drawnRecords.forEach((record) => {
    const status = normalizeRecordStatus(record.status, 'Pending')
    if (status === 'Reported') {
      return
    }

    const due = new Date(record.reportDueOn)
    if (Number.isNaN(due.getTime())) {
      return
    }

    if (due < now) {
      alerts.push({
        type: 'overdue',
        source: 'drawn',
        recordId: record.id,
        srNo: record.srNo,
        dueOn: record.reportDueOn,
        message: `Receiving report overdue for Sr.No. ${record.srNo}`
      })
      return
    }

    if (due <= maxDueDate) {
      alerts.push({
        type: 'due-soon',
        source: 'drawn',
        recordId: record.id,
        srNo: record.srNo,
        dueOn: record.reportDueOn,
        message: `Receiving report due soon for Sr.No. ${record.srNo}`
      })
    }
  })

  return res.json({ alerts })
})

app.post('/api/admin/backup', async (req, res) => {
  const fileName = await createSystemBackup(req.user?.email ?? 'admin')
  await appendAudit({ actor: req.user?.email, action: 'BACKUP_CREATE', target: fileName, details: 'Backup created' })

  return res.status(201).json({ fileName })
})

app.get('/api/admin/backups', async (_req, res) => {
  await ensureBackupsDir()
  const entries = await fs.readdir(backupsDirPath, { withFileTypes: true })
  const backups = entries
    .filter((entry) => entry.isFile() && entry.name.endsWith('.json'))
    .map((entry) => entry.name)
    .sort((first, second) => second.localeCompare(first))

  return res.json({ backups })
})

app.get('/api/admin/backups/:fileName/preview', async (req, res) => {
  try {
    const parsed = await readBackupFile(req.params.fileName)
    return res.json({ preview: summarizeBackupPayload(parsed) })
  } catch (error) {
    return res.status(error.statusCode ?? 500).json({ message: error.message ?? 'Failed to read backup preview.' })
  }
})

app.post('/api/admin/restore', async (req, res) => {
  const fileName = String(req.body?.fileName ?? '').trim()
  const requestedSections = Array.isArray(req.body?.sections) ? req.body.sections : []
  const normalizedSections = [...new Set(requestedSections.map((item) => String(item ?? '').trim()))]
  const validSections = ['users', 'issueRecords', 'drawnRecords', 'audit', 'registerHistory', 'testMaster']

  if (normalizedSections.length === 0 || normalizedSections.some((item) => !validSections.includes(item))) {
    return res.status(400).json({ message: 'Select at least one valid restore section.' })
  }

  let parsed

  try {
    parsed = await readBackupFile(fileName)
  } catch (error) {
    return res.status(error.statusCode ?? 500).json({ message: error.message ?? 'Unable to read backup file.' })
  }

  const users = Array.isArray(parsed?.users) ? parsed.users : null
  const issueRecords = Array.isArray(parsed?.registers?.issueRecords) ? parsed.registers.issueRecords : null
  const drawnRecords = Array.isArray(parsed?.registers?.drawnRecords) ? parsed.registers.drawnRecords : null
  const audit = Array.isArray(parsed?.audit) ? parsed.audit : []
  const registerHistory = Array.isArray(parsed?.registerHistory) ? parsed.registerHistory : []
  const testMaster = {
    tests: Array.isArray(parsed?.testMaster?.tests) ? parsed.testMaster.tests : [],
    parameters: Array.isArray(parsed?.testMaster?.parameters) ? parsed.testMaster.parameters : []
  }

  if (!users || !issueRecords || !drawnRecords) {
    return res.status(400).json({ message: 'Backup file format is invalid.' })
  }

  const safetyBackupFileName = await createSystemBackup(`${req.user?.email ?? 'admin'} (pre-restore)`)

  if (normalizedSections.includes('users')) {
    await writeUsers(users)
  }

  if (normalizedSections.includes('issueRecords') || normalizedSections.includes('drawnRecords')) {
    const currentRegisters = await readRegisters()
    await writeRegisters({
      issueRecords: normalizedSections.includes('issueRecords') ? issueRecords : currentRegisters.issueRecords,
      drawnRecords: normalizedSections.includes('drawnRecords') ? drawnRecords : currentRegisters.drawnRecords
    })
  }

  if (normalizedSections.includes('audit')) {
    await writeAudit(audit)
  }

  if (normalizedSections.includes('registerHistory')) {
    await writeRegisterHistory(registerHistory)
  }

  if (normalizedSections.includes('testMaster')) {
    await writeTestMaster(testMaster)
  }

  await ensureUserDefaults()
  await ensureRegisterIds()
  await ensureTestMasterFile()
  await appendAudit({
    actor: req.user?.email,
    action: 'BACKUP_RESTORE',
    target: fileName,
    details: `Backup restored (${normalizedSections.join(', ')})`,
    after: { sections: normalizedSections, safetyBackupFileName }
  })

  return res.status(200).json({ restoredSections: normalizedSections, safetyBackupFileName })
})

if (process.env.NODE_ENV === 'production') {
  app.use(express.static(clientDistPath))

  app.get(/^(?!\/api).*/, (_req, res) => {
    return res.sendFile(path.join(clientDistPath, 'index.html'))
  })
}

ensureSeedUser()
  .then(() => ensureUserDefaults())
  .then(() => ensureRegisterIds())
  .then(() => ensureAuditFile())
  .then(() => ensureRegisterHistoryFile())
  .then(() => ensureTestMasterFile())
  .then(() => ensureBackupsDir())
  .then(() => {
    app.listen(port, () => {
      console.log(`Auth API running on http://localhost:${port}`)
    })
  })
  .catch((error) => {
    console.error('Failed to start auth API:', error)
    process.exit(1)
  })
