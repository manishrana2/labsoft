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
const backupsDirPath = path.join(__dirname, 'data', 'backups')

const DEFAULT_USER = {
  email: 'admin@labsoft.dev',
  password: 'Labsoft123',
  role: 'admin'
}

const RECORD_STATUSES = ['Pending', 'In Progress', 'Reported']

const getUserRole = (user) => {
  if (user?.role === 'admin' || user?.role === 'staff') {
    return user.role
  }

  return user?.email === DEFAULT_USER.email ? 'admin' : 'staff'
}

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

const sanitizeUser = (user) => ({
  id: user.id,
  email: user.email,
  role: getUserRole(user),
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

const appendRegisterHistory = async ({ action, source, actor = 'system', recordId = '-', srNo = '-', data = null }) => {
  const entries = await readRegisterHistory()
  entries.push({
    id: `hist_${crypto.randomUUID()}`,
    action,
    source,
    actor,
    recordId,
    srNo,
    data,
    createdAt: new Date().toISOString()
  })

  await writeRegisterHistory(entries)
}

const appendAudit = async ({ actor = 'system', action, target = '-', details = '' }) => {
  const entries = await readAudit()
  entries.unshift({
    id: makeAuditId(),
    actor,
    action,
    target,
    details,
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

const ensureAuditFile = async () => {
  const entries = await readAudit()
  await writeAudit(entries)
}

const ensureRegisterHistoryFile = async () => {
  const entries = await readRegisterHistory()
  await writeRegisterHistory(entries)
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

    if (!next.role || (next.role !== 'admin' && next.role !== 'staff')) {
      next.role = getUserRole(next)
      changed = true
    }

    if (typeof next.isActive !== 'boolean') {
      next.isActive = true
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
  if (req.user?.role !== 'admin') {
    return res.status(403).json({ message: 'Admin access required.' })
  }

  return next()
}

app.use(express.json())

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
    user: { email: user.email, role }
  })
})

app.get('/api/me', authenticateToken, (req, res) => {
  return res.json({
    user: {
      email: req.user.email,
      role: req.user.role === 'admin' ? 'admin' : 'staff'
    }
  })
})

app.get('/api/registers', authenticateToken, async (_req, res) => {
  const registers = await readRegisters()
  return res.json(registers)
})

app.post('/api/registers/issue', authenticateToken, async (req, res) => {
  const record = req.body
  const requiredFields = [
    'srNo',
    'codeNo',
    'status',
    'sampleDescription',
    'parameterToBeTested',
    'issuedOn',
    'issuedBy',
    'issuedTo',
    'reportDueOn',
    'receivedBy'
  ]

  if (!requireFields(record, requiredFields)) {
    return res.status(400).json({ message: 'All issue register fields are required.' })
  }

  const registers = await readRegisters()
  const srNo = normalizeText(record.srNo)
  const codeNo = normalizeText(record.codeNo)

  const hasDuplicate = registers.issueRecords.some(
    (entry) => normalizeText(entry.srNo) === srNo || normalizeText(entry.codeNo) === codeNo
  )

  if (hasDuplicate) {
    return res.status(409).json({ message: 'Duplicate issue record: Sr.No. or Code No. already exists.' })
  }

  const normalizedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(record[field]).trim()])),
    status: normalizeRecordStatus(record.status),
    reportedOn: String(record.reportedOn ?? '').trim(),
    reportedByRemarks: String(record.reportedByRemarks ?? '').trim()
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
    data: normalizedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'ISSUE_CREATE',
    target: normalizedRecord.id,
    details: `SrNo: ${normalizedRecord.srNo}, Status: ${normalizedRecord.status}`
  })

  return res.status(201).json({ record: normalizedRecord })
})

app.put('/api/registers/issue/:id', authenticateToken, async (req, res) => {
  const { id } = req.params
  const record = req.body
  const requiredFields = [
    'srNo',
    'codeNo',
    'status',
    'sampleDescription',
    'parameterToBeTested',
    'issuedOn',
    'issuedBy',
    'issuedTo',
    'reportDueOn',
    'receivedBy'
  ]

  if (!requireFields(record, requiredFields)) {
    return res.status(400).json({ message: 'All issue register fields are required.' })
  }

  const registers = await readRegisters()
  const index = registers.issueRecords.findIndex((entry) => String(entry.id) === String(id))
  if (index === -1) {
    return res.status(404).json({ message: 'Issue record not found.' })
  }

  const srNo = normalizeText(record.srNo)
  const codeNo = normalizeText(record.codeNo)
  const hasDuplicate = registers.issueRecords.some(
    (entry) =>
      String(entry.id) !== String(id) &&
      (normalizeText(entry.srNo) === srNo || normalizeText(entry.codeNo) === codeNo)
  )

  if (hasDuplicate) {
    return res.status(409).json({ message: 'Duplicate issue record: Sr.No. or Code No. already exists.' })
  }

  const updatedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(record[field]).trim()])),
    status: normalizeRecordStatus(record.status),
    reportedOn: String(record.reportedOn ?? '').trim(),
    reportedByRemarks: String(record.reportedByRemarks ?? '').trim(),
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
    data: updatedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'ISSUE_UPDATE',
    target: updatedRecord.id,
    details: `SrNo: ${updatedRecord.srNo}, Status: ${updatedRecord.status}`
  })

  return res.json({ record: updatedRecord })
})

app.delete('/api/registers/issue/:id', authenticateToken, async (req, res) => {
  const { id } = req.params
  const registers = await readRegisters()
  const deletedRecord = registers.issueRecords.find((entry) => String(entry.id) === String(id))
  const nextRecords = registers.issueRecords.filter((entry) => String(entry.id) !== String(id))

  if (nextRecords.length === registers.issueRecords.length) {
    return res.status(404).json({ message: 'Issue record not found.' })
  }

  registers.issueRecords = nextRecords
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'DELETE',
    source: 'issue',
    actor: req.user?.email,
    recordId: id,
    srNo: deletedRecord?.srNo ?? '-',
    data: deletedRecord ?? null
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'ISSUE_DELETE',
    target: id,
    details: 'Issue record deleted'
  })

  return res.status(204).send()
})

app.post('/api/registers/drawn', authenticateToken, async (req, res) => {
  const record = req.body
  const requiredFields = [
    'srNo',
    'status',
    'sampleDescription',
    'sampleDrawnOn',
    'sampleDrawnBy',
    'customerNameAddress',
    'parameterToBeTested',
    'reportDueOn',
    'sampleReceivedBy'
  ]

  if (!requireFields(record, requiredFields)) {
    return res.status(400).json({ message: 'All drawn register fields are required.' })
  }

  const registers = await readRegisters()
  const srNo = normalizeText(record.srNo)
  const hasDuplicate = registers.drawnRecords.some((entry) => normalizeText(entry.srNo) === srNo)

  if (hasDuplicate) {
    return res.status(409).json({ message: 'Duplicate drawn record: Sr.No. already exists.' })
  }

  const normalizedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(record[field]).trim()])),
    status: normalizeRecordStatus(record.status)
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
    data: normalizedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'DRAWN_CREATE',
    target: normalizedRecord.id,
    details: `SrNo: ${normalizedRecord.srNo}, Status: ${normalizedRecord.status}`
  })

  return res.status(201).json({ record: normalizedRecord })
})

app.put('/api/registers/drawn/:id', authenticateToken, async (req, res) => {
  const { id } = req.params
  const record = req.body
  const requiredFields = [
    'srNo',
    'status',
    'sampleDescription',
    'sampleDrawnOn',
    'sampleDrawnBy',
    'customerNameAddress',
    'parameterToBeTested',
    'reportDueOn',
    'sampleReceivedBy'
  ]

  if (!requireFields(record, requiredFields)) {
    return res.status(400).json({ message: 'All drawn register fields are required.' })
  }

  const registers = await readRegisters()
  const index = registers.drawnRecords.findIndex((entry) => String(entry.id) === String(id))
  if (index === -1) {
    return res.status(404).json({ message: 'Drawn record not found.' })
  }

  const srNo = normalizeText(record.srNo)
  const hasDuplicate = registers.drawnRecords.some(
    (entry) => String(entry.id) !== String(id) && normalizeText(entry.srNo) === srNo
  )

  if (hasDuplicate) {
    return res.status(409).json({ message: 'Duplicate drawn record: Sr.No. already exists.' })
  }

  const updatedRecord = {
    ...Object.fromEntries(requiredFields.map((field) => [field, String(record[field]).trim()])),
    status: normalizeRecordStatus(record.status),
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
    data: updatedRecord
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'DRAWN_UPDATE',
    target: updatedRecord.id,
    details: `SrNo: ${updatedRecord.srNo}, Status: ${updatedRecord.status}`
  })

  return res.json({ record: updatedRecord })
})

app.delete('/api/registers/drawn/:id', authenticateToken, async (req, res) => {
  const { id } = req.params
  const registers = await readRegisters()
  const deletedRecord = registers.drawnRecords.find((entry) => String(entry.id) === String(id))
  const nextRecords = registers.drawnRecords.filter((entry) => String(entry.id) !== String(id))

  if (nextRecords.length === registers.drawnRecords.length) {
    return res.status(404).json({ message: 'Drawn record not found.' })
  }

  registers.drawnRecords = nextRecords
  await writeRegisters(registers)
  await appendRegisterHistory({
    action: 'DELETE',
    source: 'drawn',
    actor: req.user?.email,
    recordId: id,
    srNo: deletedRecord?.srNo ?? '-',
    data: deletedRecord ?? null
  })
  await appendAudit({
    actor: req.user?.email,
    action: 'DRAWN_DELETE',
    target: id,
    details: 'Drawn record deleted'
  })

  return res.status(204).send()
})

app.get('/api/admin/users', authenticateToken, requireAdmin, async (_req, res) => {
  const users = await readUsers()
  return res.json({ users: users.map(sanitizeUser) })
})

app.post('/api/admin/users', authenticateToken, requireAdmin, async (req, res) => {
  const email = String(req.body?.email ?? '').trim().toLowerCase()
  const password = String(req.body?.password ?? '').trim()
  const role = req.body?.role === 'admin' ? 'admin' : 'staff'

  if (!isValidEmail(email)) {
    return res.status(400).json({ message: 'Valid email is required.' })
  }

  if (!isStrongPassword(password)) {
    return res.status(400).json({ message: 'Password must be at least 8 chars and include letter, number, and symbol.' })
  }

  const users = await readUsers()
  if (users.some((entry) => entry.email.toLowerCase() === email)) {
    return res.status(409).json({ message: 'User already exists.' })
  }

  const passwordHash = await bcrypt.hash(password, 10)
  const user = {
    id: makeUserId(),
    email,
    role,
    isActive: true,
    passwordHash,
    createdAt: new Date().toISOString()
  }

  users.unshift(user)
  await writeUsers(users)
  await appendAudit({ actor: req.user?.email, action: 'USER_CREATE', target: user.email, details: `Role: ${role}` })

  return res.status(201).json({ user: sanitizeUser(user) })
})

app.patch('/api/admin/users/:id/status', authenticateToken, requireAdmin, async (req, res) => {
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

app.post('/api/admin/users/:id/reset-password', authenticateToken, requireAdmin, async (req, res) => {
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

app.get('/api/admin/audit', authenticateToken, requireAdmin, async (req, res) => {
  const limitRaw = Number(req.query?.limit ?? 50)
  const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(200, Math.floor(limitRaw))) : 50
  const entries = await readAudit()
  return res.json({ entries: entries.slice(0, limit) })
})

app.get('/api/admin/register-history', authenticateToken, requireAdmin, async (req, res) => {
  const limitRaw = Number(req.query?.limit ?? 50)
  const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(500, Math.floor(limitRaw))) : 50
  const entries = await readRegisterHistory()
  const sorted = [...entries].sort((first, second) => Date.parse(second.createdAt ?? '') - Date.parse(first.createdAt ?? ''))
  return res.json({ entries: sorted.slice(0, limit) })
})

app.get('/api/admin/alerts', authenticateToken, requireAdmin, async (_req, res) => {
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

app.post('/api/admin/backup', authenticateToken, requireAdmin, async (req, res) => {
  await ensureBackupsDir()
  const users = await readUsers()
  const registers = await readRegisters()
  const audit = await readAudit()
  const registerHistory = await readRegisterHistory()

  const fileName = `backup-${new Date().toISOString().replace(/[:.]/g, '-')}.json`
  const filePath = path.join(backupsDirPath, fileName)
  const payload = {
    createdAt: new Date().toISOString(),
    createdBy: req.user?.email ?? 'admin',
    users,
    registers,
    audit,
    registerHistory
  }

  await fs.writeFile(filePath, JSON.stringify(payload, null, 2))
  await appendAudit({ actor: req.user?.email, action: 'BACKUP_CREATE', target: fileName, details: 'Backup created' })

  return res.status(201).json({ fileName })
})

app.get('/api/admin/backups', authenticateToken, requireAdmin, async (_req, res) => {
  await ensureBackupsDir()
  const entries = await fs.readdir(backupsDirPath, { withFileTypes: true })
  const backups = entries
    .filter((entry) => entry.isFile() && entry.name.endsWith('.json'))
    .map((entry) => entry.name)
    .sort((first, second) => second.localeCompare(first))

  return res.json({ backups })
})

app.post('/api/admin/restore', authenticateToken, requireAdmin, async (req, res) => {
  const fileName = String(req.body?.fileName ?? '').trim()
  if (!fileName || fileName.includes('/') || fileName.includes('..')) {
    return res.status(400).json({ message: 'Valid backup file name is required.' })
  }

  const filePath = path.join(backupsDirPath, fileName)
  let parsed

  try {
    const content = await fs.readFile(filePath, 'utf8')
    parsed = JSON.parse(content)
  } catch {
    return res.status(404).json({ message: 'Backup file not found.' })
  }

  const users = Array.isArray(parsed?.users) ? parsed.users : null
  const issueRecords = Array.isArray(parsed?.registers?.issueRecords) ? parsed.registers.issueRecords : null
  const drawnRecords = Array.isArray(parsed?.registers?.drawnRecords) ? parsed.registers.drawnRecords : null
  const audit = Array.isArray(parsed?.audit) ? parsed.audit : []
  const registerHistory = Array.isArray(parsed?.registerHistory) ? parsed.registerHistory : []

  if (!users || !issueRecords || !drawnRecords) {
    return res.status(400).json({ message: 'Backup file format is invalid.' })
  }

  await writeUsers(users)
  await writeRegisters({ issueRecords, drawnRecords })
  await writeAudit(audit)
  await writeRegisterHistory(registerHistory)
  await ensureUserDefaults()
  await ensureRegisterIds()
  await appendAudit({ actor: req.user?.email, action: 'BACKUP_RESTORE', target: fileName, details: 'Backup restored' })

  return res.status(204).send()
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
