{
  "formKey": "letsFixIt_v2",
  "title": "Let's Fix It 2.0",
  "issueTypes": [
    {
      "key": "ERROR_404",
      "label": "404 Error",
      "description": "Unable to access an Online Help page or web page."
    },
    {
      "key": "INACCURATE_CONTENT",
      "label": "Inaccurate Content",
      "description": "Content with incorrect process/invalid information/typos/missing image."
    },
    {
      "key": "SEARCH_BROKEN",
      "label": "Search Broken",
      "description": "Unable to use the Search feature / incorrect Search results."
    },
    {
      "key": "TOPIC_INACCESSIBLE",
      "label": "Topic Inaccessible",
      "description": "Link to topic does not open; received removed or deleted message."
    },
    {
      "key": "TOPIC_LINKS_BROKEN",
      "label": "Topic Links Broken",
      "description": "Unable to view a link within a topic."
    },
    {
      "key": "WEB_NAVIGATION_BROKEN",
      "label": "Web Navigation Broken",
      "description": "Unable to navigate across pages (Tips, Let’s Fix It, etc.)."
    }
  ],
  "statusOptions": [
    { "key": "OPEN", "label": "Open" },
    { "key": "ASSIGNED_TO_TW", "label": "Assigned to TW" },
    { "key": "IN_PROGRESS", "label": "In Progress" },
    { "key": "REASSIGNED_TO_TW", "label": "Reassigned to TW" },
    { "key": "ROUTE_TO_TEAM_LEAD", "label": "Route to Team Lead" },
    { "key": "ROUTE_TO_SUBMITTER", "label": "Route to Submitter" },
    { "key": "RETURNED_FROM_ROUTE", "label": "Returned from Route" },
    { "key": "COMPLETE", "label": "Complete" },
    { "key": "CANCELLED", "label": "Cancelled" }
  ],
  "resolutionTypes": [
    {
      "group": "COMPLETE",
      "options": [
        { "key": "CONTENT_CHANGED_TOPIC_UPDATED", "label": "Content Changed / Topic Updated" },
        { "key": "BROKEN_LINKS_RESOLVED", "label": "Broken Links Resolved" },
        { "key": "EDUCATION_PROVIDED", "label": "Education Provided" },
        { "key": "TECHNICAL_ISSUE_RESOLVED", "label": "Technical Issue Resolved" }
      ]
    },
    {
      "group": "CANCELLED",
      "options": [
        { "key": "NOT_A_VALID_CHANGE_REQUEST", "label": "Not a Valid Change Request" },
        { "key": "NO_RESPONSE_FROM_SUBMITTER", "label": "No Response from Submitter" },
        { "key": "NO_RESPONSE_FROM_TEAM_LEAD", "label": "No Response from Team Lead" }
      ]
    }
  ],
  "fields": [
    {
      "key": "issueType",
      "label": "Issue Type",
      "type": "select",
      "required": true,
      "optionsFrom": "issueTypes"
    },
    {
      "key": "submitter",
      "label": "Submitter",
      "type": "person",
      "required": true,
      "readOnly": true,
      "autoPopulate": "currentUser"
    },
    {
      "key": "department",
      "label": "Department",
      "type": "text",
      "required": true,
      "readOnly": true,
      "autoPopulate": "currentUser.department"
    },
    {
      "key": "pageUrl",
      "label": "Page URL",
      "type": "url",
      "required": true
    },
    {
      "key": "topicName",
      "label": "Topic Name",
      "type": "text",
      "required": true
    },
    {
      "key": "topicErrorMessage",
      "label": "Topic Error Message",
      "type": "text",
      "requiredWhen": {
        "field": "issueType",
        "equals": "ERROR_404"
      }
    },
    {
      "key": "describeIssue",
      "label": "Describe the Issue",
      "type": "textarea",
      "required": true
    },

    {
      "key": "inaccurateContentDetails",
      "label": "Inaccurate Content",
      "type": "textarea",
      "visibleWhen": { "field": "issueType", "equals": "INACCURATE_CONTENT" },
      "requiredWhen": { "field": "issueType", "equals": "INACCURATE_CONTENT" }
    },

    {
      "key": "searchBrokenDetails",
      "label": "Search Broken Details",
      "type": "selectOrOther",
      "visibleWhen": { "field": "issueType", "equals": "SEARCH_BROKEN" },
      "requiredWhen": { "field": "issueType", "equals": "SEARCH_BROKEN" },
      "options": [
        { "key": "X_BUTTON_DOES_NOT_REFRESH", "label": "\"X\" button does not refresh search" },
        { "key": "OTHER", "label": "Other" }
      ],
      "otherKey": "searchBrokenOtherText"
    },

    {
      "key": "brokenLinkType",
      "label": "Broken Link Type",
      "type": "select",
      "visibleWhen": { "field": "issueType", "equals": "TOPIC_LINKS_BROKEN" },
      "requiredWhen": { "field": "issueType", "equals": "TOPIC_LINKS_BROKEN" },
      "options": [
        { "key": "DROPDOWN", "label": "Dropdown" },
        { "key": "POPUPS", "label": "Pop-ups" },
        { "key": "LINK_EXTERNAL_PAGE", "label": "Link to external page" },
        { "key": "LINK_OTHER_PAGE", "label": "Link to other page" },
        { "key": "LINK_DOCUMENT", "label": "Link to document" },
        { "key": "FAVORITES_LINK", "label": "Favorites link" },
        { "key": "OTHER", "label": "Other" }
      ]
    },
    {
      "key": "topicLinksDetails",
      "label": "Topic Links Details",
      "type": "textarea",
      "visibleWhen": { "field": "issueType", "equals": "TOPIC_LINKS_BROKEN" },
      "requiredWhen": { "field": "issueType", "equals": "TOPIC_LINKS_BROKEN" }
    },

    {
      "key": "comments",
      "label": "Comments",
      "type": "textarea",
      "required": false
    },
    {
      "key": "internalCommentsHistory",
      "label": "Internal Comments History",
      "type": "auditLog",
      "readOnly": true
    },
    {
      "key": "attachments",
      "label": "Supporting Documentation / Attachments",
      "type": "attachments",
      "required": false
    },

    {
      "key": "requestId",
      "label": "Request ID",
      "type": "text",
      "readOnly": true,
      "autoPopulate": "system.guid"
    },
    {
      "key": "submittedBy",
      "label": "Submitted By",
      "type": "person",
      "readOnly": true,
      "autoPopulate": "currentUser"
    },
    {
      "key": "overallRequestStatus",
      "label": "Overall Request Status",
      "type": "statusHistory",
      "readOnly": true
    },

    {
      "key": "assignedTo",
      "label": "Assigned To",
      "type": "person",
      "required": false
    },
    {
      "key": "assignToTeamLead",
      "label": "Assign to Team Lead",
      "type": "person",
      "requiredWhen": {
        "field": "status",
        "equals": "ROUTE_TO_TEAM_LEAD"
      }
    },

    {
      "key": "status",
      "label": "Status",
      "type": "select",
      "required": true,
      "optionsFrom": "statusOptions",
      "default": "OPEN"
    },
    {
      "key": "resolutionType",
      "label": "Resolution Type",
      "type": "select",
      "requiredWhenAny": [
        { "field": "status", "equals": "COMPLETE" },
        { "field": "status", "equals": "CANCELLED" }
      ],
      "optionsFrom": "resolutionTypes"
    }
  ],
  "routingRules": [
    {
      "ifIssueTypeIn": ["WEB_NAVIGATION_BROKEN", "SEARCH_BROKEN"],
      "thenAssignRole": "KS_DEVELOPER",
      "thenSetStatus": "ASSIGNED_TO_KS_DEV"
    },
    {
      "ifIssueTypeIn": ["ERROR_404", "TOPIC_INACCESSIBLE", "TOPIC_LINKS_BROKEN", "INACCURATE_CONTENT"],
      "thenAssignRole": "TECHNICAL_WRITER",
      "thenSetStatus": "ASSIGNED_TO_TW"
    }
  ],
  "validationRules": [
    {
      "name": "requireResolutionWhenClosed",
      "when": { "field": "status", "in": ["COMPLETE", "CANCELLED"] },
      "requireFields": ["resolutionType", "comments"]
    }
  ]
}
{
  "emailWorkflows": [
    {
      "status": "OPEN",
      "trigger": "ON_CREATE",
      "sendEmail": true,
      "recipients": {
        "to": ["proceduresAnalyst"],
        "cc": []
      },
      "email": {
        "subject": "New Let’s Fix It Request Submitted – {{requestId}}",
        "body": "A new Let’s Fix It request has been submitted.\n\nRequest ID: {{requestId}}\nIssue Type: {{issueType}}\nTopic: {{topicName}}\n\nPlease review and assign the request.\n\nView Request:\n{{formUrl}}"
      }
    },

    {
      "status": "ASSIGNED_TO_TW",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["assignedTo"],
        "cc": ["proceduresAnalyst"]
      },
      "email": {
        "subject": "Let’s Fix It Request Assigned to You – {{requestId}}",
        "body": "You have been assigned a Let’s Fix It request.\n\nRequest ID: {{requestId}}\nIssue Type: {{issueType}}\nTopic: {{topicName}}\n\nPlease begin review.\n\nView Request:\n{{formUrl}}"
      }
    },

    {
      "status": "REASSIGNED_TO_TW",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["assignedTo"],
        "cc": ["proceduresAnalyst"]
      },
      "email": {
        "subject": "Let’s Fix It Request Reassigned – {{requestId}}",
        "body": "A Let’s Fix It request has been reassigned to you.\n\nRequest ID: {{requestId}}\nTopic: {{topicName}}\n\nComments:\n{{comments}}\n\nView Request:\n{{formUrl}}"
      }
    },

    {
      "status": "ROUTE_TO_TEAM_LEAD",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["teamLead"],
        "cc": ["assignedTo", "proceduresAnalyst"]
      },
      "email": {
        "subject": "Action Required: Let’s Fix It Request – {{requestId}}",
        "body": "A Let’s Fix It request requires your review.\n\nRequest ID: {{requestId}}\nTopic: {{topicName}}\n\nPlease review and respond within 5 business days.\n\nComments:\n{{comments}}\n\nView Request:\n{{formUrl}}"
      }
    },

    {
      "status": "ROUTE_TO_SUBMITTER",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["submitter"],
        "cc": ["assignedTo"]
      },
      "email": {
        "subject": "More Information Needed – Let’s Fix It {{requestId}}",
        "body": "Additional information is needed to continue processing your Let’s Fix It request.\n\nRequest ID: {{requestId}}\nTopic: {{topicName}}\n\nRequested Details:\n{{comments}}\n\nPlease update the request using the link below.\n\n{{formUrl}}"
      }
    },

    {
      "status": "RETURNED_FROM_ROUTE",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["assignedTo"],
        "cc": []
      },
      "email": {
        "subject": "Let’s Fix It Request Updated by Business – {{requestId}}",
        "body": "The submitter has provided additional information for the Let’s Fix It request.\n\nRequest ID: {{requestId}}\nTopic: {{topicName}}\n\nYou may continue processing.\n\nView Request:\n{{formUrl}}"
      }
    },

    {
      "status": "COMPLETE",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["submitter"],
        "cc": ["assignedTo"]
      },
      "email": {
        "subject": "Let’s Fix It Request Completed – {{requestId}}",
        "body": "Your Let’s Fix It request has been completed.\n\nRequest ID: {{requestId}}\nTopic: {{topicName}}\nResolution: {{resolutionType}}\n\nComments:\n{{comments}}\n\nThank you."
      }
    },

    {
      "status": "CANCELLED",
      "trigger": "ON_STATUS_CHANGE",
      "sendEmail": true,
      "recipients": {
        "to": ["submitter"],
        "cc": ["assignedTo"]
      },
      "email": {
        "subject": "Let’s Fix It Request Cancelled – {{requestId}}",
        "body": "Your Let’s Fix It request has been cancelled.\n\nRequest ID: {{requestId}}\nTopic: {{topicName}}\nReason: {{resolutionType}}\n\nComments:\n{{comments}}\n\nIf needed, please submit a new request."
      }
    }
  ]
}
