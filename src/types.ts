// Shared types for Labsoft
export type UserRole = 'admin' | 'staff' | 'customer';
export type ModuleKey = 'issue-entry' | 'issue-records' | 'drawn-entry' | 'drawn-records' | 'admin-panel';

export interface IssueRecord {
  id?: string;
  createdAt?: string;
  status?: 'Pending' | 'In Progress' | 'Reported';
  receivedByName?: string;
  srNo: string;
  codeNo: string;
  ulrNo?: string;
  sampleDescription: string;
  parameterToBeTested: string;
  issuedOn: string;
  issuedBy: string;
  issuedTo: string;
  reportDueOn: string;
  receivedBy: string;
  reportedOn: string;
  reportedByRemarks: string;
}

export interface DrawnRecord {
  id?: string;
  createdAt?: string;
  status?: 'Pending' | 'In Progress' | 'Reported';
  sampleReceivedByName?: string;
  srNo: string;
  reportCode?: string;
  ulrNo?: string;
  sampleDescription: string;
  sampleDrawnOn: string;
  sampleDrawnBy: string;
  customerNameAddress: string;
  parameterToBeTested: string;
  reportDueOn: string;
  sampleReceivedBy: string;
}

export interface Session {
  token: string;
  email: string;
  name?: string;
  role: UserRole;
  userCode: string;
}

// Placeholder types for missing declarations
export interface AdminUser {
  id: string;
  email: string;
  name: string;
  role: UserRole;
  userCode?: string;
  isActive: boolean;
  createdAt: string;
}

export interface AdminAlert {
  type: string;
  message: string;
  dueOn: string;
}

export interface AuditEntry {
  action: string;
  actor: string;
  createdAt: string;
}

export interface RegisterHistoryEntry {
  action: string;
  source: string;
  srNo: string;
  createdAt: string;
}
