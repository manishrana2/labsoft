import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { Pool } from 'pg'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const dataDirPath = path.join(__dirname, 'data')
const usersDbPath = path.join(dataDirPath, 'users.json')
const registersDbPath = path.join(dataDirPath, 'registers.json')
const auditDbPath = path.join(dataDirPath, 'audit.json')
const registerHistoryDbPath = path.join(dataDirPath, 'register-history.json')
const testMasterDbPath = path.join(dataDirPath, 'test-master.json')

const databaseUrl = String(process.env.DATABASE_URL ?? '').trim()
const useDatabase = databaseUrl.length > 0

const isLocalDatabaseUrl = (value) => /localhost|127\.0\.0\.1/i.test(value)

const pool = useDatabase
  ? new Pool({
      connectionString: databaseUrl,
      ssl: isLocalDatabaseUrl(databaseUrl) ? false : { rejectUnauthorized: false }
    })
  : null

const readJsonFile = async (filePath, fallback) => {
  try {
    const content = await fs.readFile(filePath, 'utf8')
    const parsed = JSON.parse(content)
    return typeof fallback === 'function' ? fallback(parsed) : parsed
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
      return typeof fallback === 'function' ? fallback(null) : fallback
    }

    throw error
  }
}

const normalizeUsers = (users) => (Array.isArray(users) ? users : [])
const normalizeRegisters = (parsed) => ({
  issueRecords: Array.isArray(parsed?.issueRecords) ? parsed.issueRecords : [],
  drawnRecords: Array.isArray(parsed?.drawnRecords) ? parsed.drawnRecords : []
})
const normalizeEntries = (entries) => (Array.isArray(entries) ? entries : [])
const normalizeTestMaster = (parsed) => ({
  tests: Array.isArray(parsed?.tests) ? parsed.tests : [],
  parameters: Array.isArray(parsed?.parameters) ? parsed.parameters : []
})

const readUsersFromFiles = async () => readJsonFile(usersDbPath, normalizeUsers)
const readRegistersFromFiles = async () => readJsonFile(registersDbPath, normalizeRegisters)
const readAuditFromFiles = async () => readJsonFile(auditDbPath, normalizeEntries)
const readRegisterHistoryFromFiles = async () => readJsonFile(registerHistoryDbPath, normalizeEntries)
const readTestMasterFromFiles = async () => readJsonFile(testMasterDbPath, normalizeTestMaster)

const writeUsersToFiles = async (users) => {
  await fs.writeFile(usersDbPath, JSON.stringify(users, null, 2))
}

const writeRegistersToFiles = async (registers) => {
  await fs.writeFile(registersDbPath, JSON.stringify(registers, null, 2))
}

const writeAuditToFiles = async (entries) => {
  await fs.writeFile(auditDbPath, JSON.stringify(entries, null, 2))
}

const writeRegisterHistoryToFiles = async (entries) => {
  await fs.writeFile(registerHistoryDbPath, JSON.stringify(entries, null, 2))
}

const writeTestMasterToFiles = async (payload) => {
  await fs.writeFile(testMasterDbPath, JSON.stringify(payload, null, 2))
}

const ensureSchema = async () => {
  if (!pool) {
    return
  }

  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      data JSONB NOT NULL
    );

    CREATE TABLE IF NOT EXISTS issue_records (
      id TEXT PRIMARY KEY,
      sr_no TEXT NOT NULL DEFAULT '',
      code_no TEXT NOT NULL DEFAULT '',
      ulr_no TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      data JSONB NOT NULL
    );

    CREATE TABLE IF NOT EXISTS drawn_records (
      id TEXT PRIMARY KEY,
      sr_no TEXT NOT NULL DEFAULT '',
      report_code TEXT NOT NULL DEFAULT '',
      ulr_no TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      data JSONB NOT NULL
    );

    CREATE TABLE IF NOT EXISTS audit_entries (
      id TEXT PRIMARY KEY,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      data JSONB NOT NULL
    );

    CREATE TABLE IF NOT EXISTS register_history_entries (
      id TEXT PRIMARY KEY,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      data JSONB NOT NULL
    );

    CREATE TABLE IF NOT EXISTS test_master_sections (
      section TEXT PRIMARY KEY,
      data JSONB NOT NULL
    );
  `)
}

const countRows = async (tableName) => {
  const result = await pool.query(`SELECT COUNT(*)::int AS count FROM ${tableName}`)
  return Number(result.rows[0]?.count ?? 0)
}

const replaceTable = async (client, tableName, rows, mapRow) => {
  await client.query(`DELETE FROM ${tableName}`)

  for (const row of rows) {
    const mapped = mapRow(row)
    await client.query(mapped.text, mapped.values)
  }
}

const migrateFileDataToDatabase = async () => {
  if (!pool) {
    return
  }

  const [usersCount, issueCount, drawnCount, auditCount, historyCount, testMasterCount] = await Promise.all([
    countRows('users'),
    countRows('issue_records'),
    countRows('drawn_records'),
    countRows('audit_entries'),
    countRows('register_history_entries'),
    countRows('test_master_sections')
  ])

  if (usersCount || issueCount || drawnCount || auditCount || historyCount || testMasterCount) {
    return
  }

  const [users, registers, audit, registerHistory, testMaster] = await Promise.all([
    readUsersFromFiles(),
    readRegistersFromFiles(),
    readAuditFromFiles(),
    readRegisterHistoryFromFiles(),
    readTestMasterFromFiles()
  ])

  const client = await pool.connect()
  try {
    await client.query('BEGIN')

    await replaceTable(client, 'users', users, (user) => ({
      text: 'INSERT INTO users (id, created_at, data) VALUES ($1, COALESCE($2::timestamptz, NOW()), $3::jsonb)',
      values: [String(user.id), String(user.createdAt ?? '') || null, JSON.stringify(user)]
    }))

    await replaceTable(client, 'issue_records', registers.issueRecords, (record) => ({
      text: 'INSERT INTO issue_records (id, sr_no, code_no, ulr_no, created_at, data) VALUES ($1, $2, $3, $4, COALESCE($5::timestamptz, NOW()), $6::jsonb)',
      values: [
        String(record.id),
        String(record.srNo ?? ''),
        String(record.codeNo ?? ''),
        String(record.ulrNo ?? ''),
        String(record.createdAt ?? '') || null,
        JSON.stringify(record)
      ]
    }))

    await replaceTable(client, 'drawn_records', registers.drawnRecords, (record) => ({
      text: 'INSERT INTO drawn_records (id, sr_no, report_code, ulr_no, created_at, data) VALUES ($1, $2, $3, $4, COALESCE($5::timestamptz, NOW()), $6::jsonb)',
      values: [
        String(record.id),
        String(record.srNo ?? ''),
        String(record.reportCode ?? ''),
        String(record.ulrNo ?? ''),
        String(record.createdAt ?? '') || null,
        JSON.stringify(record)
      ]
    }))

    await replaceTable(client, 'audit_entries', audit, (entry) => ({
      text: 'INSERT INTO audit_entries (id, created_at, data) VALUES ($1, COALESCE($2::timestamptz, NOW()), $3::jsonb)',
      values: [String(entry.id), String(entry.createdAt ?? '') || null, JSON.stringify(entry)]
    }))

    await replaceTable(client, 'register_history_entries', registerHistory, (entry) => ({
      text: 'INSERT INTO register_history_entries (id, created_at, data) VALUES ($1, COALESCE($2::timestamptz, NOW()), $3::jsonb)',
      values: [String(entry.id), String(entry.createdAt ?? '') || null, JSON.stringify(entry)]
    }))

    await replaceTable(
      client,
      'test_master_sections',
      [
        { section: 'tests', data: testMaster.tests },
        { section: 'parameters', data: testMaster.parameters }
      ],
      (entry) => ({
        text: 'INSERT INTO test_master_sections (section, data) VALUES ($1, $2::jsonb)',
        values: [entry.section, JSON.stringify(entry.data)]
      })
    )

    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

const mapRowData = (rows) => rows.map((row) => row.data)

const readUsersFromDatabase = async () => {
  const result = await pool.query('SELECT data FROM users ORDER BY created_at DESC, id DESC')
  return normalizeUsers(mapRowData(result.rows))
}

const writeUsersToDatabase = async (users) => {
  const client = await pool.connect()
  try {
    await client.query('BEGIN')
    await replaceTable(client, 'users', users, (user) => ({
      text: 'INSERT INTO users (id, created_at, data) VALUES ($1, COALESCE($2::timestamptz, NOW()), $3::jsonb)',
      values: [String(user.id), String(user.createdAt ?? '') || null, JSON.stringify(user)]
    }))
    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

const readRegistersFromDatabase = async () => {
  const [issueResult, drawnResult] = await Promise.all([
    pool.query('SELECT data FROM issue_records ORDER BY created_at DESC, id DESC'),
    pool.query('SELECT data FROM drawn_records ORDER BY created_at DESC, id DESC')
  ])

  return {
    issueRecords: normalizeEntries(mapRowData(issueResult.rows)),
    drawnRecords: normalizeEntries(mapRowData(drawnResult.rows))
  }
}

const writeRegistersToDatabase = async (registers) => {
  const client = await pool.connect()
  try {
    await client.query('BEGIN')
    await replaceTable(client, 'issue_records', registers.issueRecords, (record) => ({
      text: 'INSERT INTO issue_records (id, sr_no, code_no, ulr_no, created_at, data) VALUES ($1, $2, $3, $4, COALESCE($5::timestamptz, NOW()), $6::jsonb)',
      values: [
        String(record.id),
        String(record.srNo ?? ''),
        String(record.codeNo ?? ''),
        String(record.ulrNo ?? ''),
        String(record.createdAt ?? '') || null,
        JSON.stringify(record)
      ]
    }))
    await replaceTable(client, 'drawn_records', registers.drawnRecords, (record) => ({
      text: 'INSERT INTO drawn_records (id, sr_no, report_code, ulr_no, created_at, data) VALUES ($1, $2, $3, $4, COALESCE($5::timestamptz, NOW()), $6::jsonb)',
      values: [
        String(record.id),
        String(record.srNo ?? ''),
        String(record.reportCode ?? ''),
        String(record.ulrNo ?? ''),
        String(record.createdAt ?? '') || null,
        JSON.stringify(record)
      ]
    }))
    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

const readAuditFromDatabase = async () => {
  const result = await pool.query('SELECT data FROM audit_entries ORDER BY created_at DESC, id DESC')
  return normalizeEntries(mapRowData(result.rows))
}

const writeAuditToDatabase = async (entries) => {
  const client = await pool.connect()
  try {
    await client.query('BEGIN')
    await replaceTable(client, 'audit_entries', entries, (entry) => ({
      text: 'INSERT INTO audit_entries (id, created_at, data) VALUES ($1, COALESCE($2::timestamptz, NOW()), $3::jsonb)',
      values: [String(entry.id), String(entry.createdAt ?? '') || null, JSON.stringify(entry)]
    }))
    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

const readRegisterHistoryFromDatabase = async () => {
  const result = await pool.query('SELECT data FROM register_history_entries ORDER BY created_at ASC, id ASC')
  return normalizeEntries(mapRowData(result.rows))
}

const writeRegisterHistoryToDatabase = async (entries) => {
  const client = await pool.connect()
  try {
    await client.query('BEGIN')
    await replaceTable(client, 'register_history_entries', entries, (entry) => ({
      text: 'INSERT INTO register_history_entries (id, created_at, data) VALUES ($1, COALESCE($2::timestamptz, NOW()), $3::jsonb)',
      values: [String(entry.id), String(entry.createdAt ?? '') || null, JSON.stringify(entry)]
    }))
    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

const readTestMasterFromDatabase = async () => {
  const result = await pool.query('SELECT section, data FROM test_master_sections')
  const sections = new Map(result.rows.map((row) => [row.section, row.data]))
  return {
    tests: normalizeEntries(sections.get('tests')),
    parameters: normalizeEntries(sections.get('parameters'))
  }
}

const writeTestMasterToDatabase = async (payload) => {
  const client = await pool.connect()
  try {
    await client.query('BEGIN')
    await replaceTable(
      client,
      'test_master_sections',
      [
        { section: 'tests', data: Array.isArray(payload?.tests) ? payload.tests : [] },
        { section: 'parameters', data: Array.isArray(payload?.parameters) ? payload.parameters : [] }
      ],
      (entry) => ({
        text: 'INSERT INTO test_master_sections (section, data) VALUES ($1, $2::jsonb)',
        values: [entry.section, JSON.stringify(entry.data)]
      })
    )
    await client.query('COMMIT')
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

export const storageMode = useDatabase ? 'postgres' : 'json'

export const initializeStorage = async () => {
  if (!pool) {
    return
  }

  await ensureSchema()
  await migrateFileDataToDatabase()
}

export const readUsers = async () => (pool ? readUsersFromDatabase() : readUsersFromFiles())
export const writeUsers = async (users) => (pool ? writeUsersToDatabase(users) : writeUsersToFiles(users))
export const readRegisters = async () => (pool ? readRegistersFromDatabase() : readRegistersFromFiles())
export const writeRegisters = async (registers) => (pool ? writeRegistersToDatabase(registers) : writeRegistersToFiles(registers))
export const readAudit = async () => (pool ? readAuditFromDatabase() : readAuditFromFiles())
export const writeAudit = async (entries) => (pool ? writeAuditToDatabase(entries) : writeAuditToFiles(entries))
export const readRegisterHistory = async () => (pool ? readRegisterHistoryFromDatabase() : readRegisterHistoryFromFiles())
export const writeRegisterHistory = async (entries) => (pool ? writeRegisterHistoryToDatabase(entries) : writeRegisterHistoryToFiles(entries))
export const readTestMaster = async () => (pool ? readTestMasterFromDatabase() : readTestMasterFromFiles())
export const writeTestMaster = async (payload) => (pool ? writeTestMasterToDatabase(payload) : writeTestMasterToFiles(payload))
