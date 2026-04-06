export type AuthorizedRequestorBusStatus =
  | 'Submitted'
  | 'Approved'
  | 'Rejected';

export interface SubmitToCcOwnerVars {
  ccOwnerName: string;
  ccOwnerEmail: string;
  requesterFullName: string;
  requestTypeText: string;
  requestId: string;
  editPageUrl: string;
  busStatus: 'Submitted';
}

export interface SubmitToEokmVars {
  requesterFullName: string;
  requestTypeText: string;
  requestId: string;
  editPageUrl: string;
  busStatus: 'Submitted';
}

export interface ApprovedToEokmVars {
  requesterFullName: string;
  requestTypeText: string;
  requestId: string;
  editPageUrl: string;
  busStatus: 'Approved';
}

export interface RejectedToEokmVars {
  requesterFullName: string;
  requestTypeText: string;
  requestId: string;
  editPageUrl: string;
  busStatus: 'Rejected';
}

