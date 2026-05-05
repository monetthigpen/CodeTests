{
  "$schema": "./formRules.schema.json",
  "title": "Authorized Requestors Form Rules",
  "sourceType": "LIST",
  "_comment": "Simplified Authorized Requestors rules. No dependent status hide/show logic.",
  "items": [
    {
      "formMode": 8,
      "_comment": "New Form Mode",
      "items": {
        "globalDisableFields": {
          "items": []
        },
        "globalHiddenFields": {
          "items": [
            "Title",
            "Status",
            "Access_x0020_Granted"
          ]
        },
        "userPermsBased": {
          "ownerGroup": {
            "priority": 1,
            "groupName": "Request Form Owners",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ]
          },
          "internalUsers": {
            "priority": 2,
            "groupName": "Knowledge Management Internal Users",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ]
          },
          "externalUsers": {
            "priority": 3,
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ]
          }
        },
        "autoPopulate": {
          "items": [
            {
              "when": {
                "field": "Status",
                "equals": [
                  ""
                ]
              },
              "set": {
                "field": "Status",
                "value": "Submitted"
              }
            }
          ]
        }
      }
    },
    {
      "formMode": 6,
      "_comment": "Edit Form Mode",
      "items": {
        "globalDisableFields": {
          "items": [
            "Title"
          ]
        },
        "globalHiddenFields": {
          "items": [
            "Title"
          ]
        },
        "finalState": {
          "items": [
            {
              "fieldName": "Status",
              "fieldValues": [
                "Approved",
                "Rejected"
              ]
            }
          ]
        },
        "userPermsBased": {
          "ownerGroup": {
            "priority": 1,
            "groupName": "Request Form Owners",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ]
          },
          "internalUsers": {
            "priority": 2,
            "groupName": "Knowledge Management Internal Users",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ]
          },
          "externalUsers": {
            "priority": 3,
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "btnSubmit",
              "attachments"
            ]
          }
        }
      }
    },
    {
      "formMode": 4,
      "_comment": "View Form Mode",
      "items": {
        "globalDisableFields": {
          "items": []
        },
        "globalHiddenFields": {
          "items": [
            "Title"
          ]
        },
        "userPermsBased": {
          "ownerGroup": {
            "priority": 1,
            "groupName": "Request Form Owners",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "attachments"
            ]
          },
          "internalUsers": {
            "priority": 2,
            "groupName": "Knowledge Management Internal Users",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "attachments"
            ]
          },
          "externalUsers": {
            "priority": 3,
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Business_x0020_Area",
              "Cost_x0020_Center",
              "attachments"
            ]
          }
        }
      }
    }
  ]
}


