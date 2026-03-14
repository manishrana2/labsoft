import { IssueRecord } from '../types'

function getIssueStatus(record: IssueRecord): 'Pending' | 'In Progress' | 'Reported' {
  if (record.status === 'Pending' || record.status === 'In Progress' || record.status === 'Reported') {
    return record.status
  }

  return String(record.reportedOn ?? '').trim() ? 'Reported' : 'Pending'
}

function getIssueStatusMeta(status: 'Pending' | 'In Progress' | 'Reported'): { label: string; className: string } {
  if (status === 'Reported') {
    return { label: 'Reported', className: 'status-reported' }
  }

  if (status === 'In Progress') {
    return { label: 'Under Process', className: 'status-progress' }
  }

  return { label: 'Pending', className: 'status-pending' }
}

// IssueTable component: renders the issue records table
export function renderIssueTable(records: IssueRecord[], canDelete: boolean, canEdit: boolean, canManageStatus: boolean) {
  if (!records || records.length === 0) {
    return '<p class="empty-state">No entries yet.</p>';
  }
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Sr.No.</th>
            <th>Report Number</th>
            <th>ULR No.</th>
            <th>Sample Description</th>
            <th>Parameter to Be Tested</th>
            <th>Issued On</th>
            <th>Issued By</th>
            <th>Issued To</th>
            <th>Report Due On</th>
            <th>Received By</th>
            <th>Reported On</th>
            <th>Reported By/Remarks</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${records
            .map(
              (item) => {
                const status = getIssueStatus(item)
                const statusMeta = getIssueStatusMeta(status)
                return `
            <tr>
              <td>${item.srNo}</td>
              <td>${item.codeNo}</td>
              <td>${item.ulrNo ?? ''}</td>
              <td>${item.sampleDescription}</td>
              <td>${item.parameterToBeTested}</td>
              <td>${item.issuedOn}</td>
              <td>${item.issuedBy}</td>
              <td>${item.issuedTo}</td>
              <td>${item.reportDueOn}</td>
              <td>${item.receivedByName ?? item.receivedBy}</td>
              <td>${item.reportedOn}</td>
              <td>${item.reportedByRemarks}</td>
              <td class="actions-col">
                <span class="status-chip ${statusMeta.className}${canManageStatus ? ' issue-status-interactive' : ''}" ${canManageStatus ? `data-issue-status-id="${item.id ?? ''}"` : ''}>${statusMeta.label}</span>
                ${canManageStatus ? `<div class="status-dropdown hidden" data-status-dropdown="${item.id ?? ''}">
                  <button type="button" data-status-option="In Progress">Under Process</button>
                  <button type="button" data-status-option="Reported">Analysis Completed</button>
                  <button type="button" data-status-option="Pending">Mark Pending</button>
                </div>` : ''}
                ${canEdit ? `<button class="table-action edit" data-issue-edit="${item.id ?? ''}" type="button">Edit</button>` : ''}
                ${canDelete ? `<button class="table-action delete" data-issue-delete="${item.id ?? ''}" type="button">Delete</button>` : ''}
              </td>
            </tr>
          `;
              }
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}
