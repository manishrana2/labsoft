import { DrawnRecord } from '../types'
// DrawnTable component: renders the drawn records table
export function renderDrawnTable(records: DrawnRecord[], canDelete: boolean, canEdit: boolean) {
  if (!records || records.length === 0) {
    return '<p class="empty-state">No entries yet.</p>';
  }
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Sr.No.</th>
            <th>Report Code</th>
            <th>ULR No.</th>
            <th>Sample Description</th>
            <th>Sample Drawn On</th>
            <th>Sample Drawn By</th>
            <th>Customer Name & Address</th>
            <th>Parameter to be Tested</th>
            <th>Report Due On</th>
            <th>Sample Received By</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${records
            .map(
              (item) => {
                return `
            <tr>
              <td>${item.srNo}</td>
              <td>${item.reportCode ?? ''}</td>
              <td>${item.ulrNo ?? ''}</td>
              <td>${item.sampleDescription}</td>
              <td>${item.sampleDrawnOn}</td>
              <td>${item.sampleDrawnBy}</td>
              <td>${item.customerNameAddress}</td>
              <td>${item.parameterToBeTested}</td>
              <td>${item.reportDueOn}</td>
              <td>${item.sampleReceivedByName ?? item.sampleReceivedBy}</td>
              <td class="actions-col">
                ${canEdit ? `<button class="table-action edit" data-drawn-edit="${item.id ?? ''}" type="button">Edit</button>` : ''}
                ${canDelete ? `<button class="table-action delete" data-drawn-delete="${item.id ?? ''}" type="button">Delete</button>` : ''}
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
