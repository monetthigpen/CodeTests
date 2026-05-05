{
  "$schema": "./formRules.schema.json",
  "title": "Authorized Requestors Form Rules",
  "sourceType": "LIST",
  "items": [
    {
      "formMode": 8,
      "_comment": "NEW FORM",
      "items": {
        "globalDisableFields": {
          "items": [
            "Business_x0020_Area"
          ]
        },
        "globalHiddenFields": {
          "items": [
            "Title",
            "Status",
            "Access_x0020_Granted"
          ]
        },
        "autoPopulate": {
          "items": [
            {
              "_comment": "Set default status",
              "when": {
                "field": "Status",
                "equals": [""]
              },
              "set": {
                "field": "Status",
                "value": "Submitted"
              }
            },
            {
              "_comment": "Map Cost Center → Business Area",
              "when": {
                "field": "Cost_x0020_Center",
                "equals": ["REPLACE_CC_1"]
              },
              "set": {
                "field": "Business_x0020_Area",
                "value": "REPLACE_AREA_1"
              }
            },
            {
              "when": {
                "field": "Cost_x0020_Center",
                "equals": ["REPLACE_CC_2"]
              },
              "set": {
                "field": "Business_x0020_Area",
                "value": "REPLACE_AREA_2"
              }
            }
          ]
        },
        "userPermsBased": {
          "ownerGroup": {
            "priority": 1,
            "groupName": "Request Form Owners",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ]
          },
          "internalUsers": {
            "priority": 2,
            "groupName": "Knowledge Management Internal Users",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ]
          },
          "externalUsers": {
            "priority": 3,
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ]
          }
        }
      }
    },
    {
      "formMode": 6,
      "_comment": "EDIT FORM",
      "items": {
        "globalDisableFields": {
          "items": [
            "Title",
            "Business_x0020_Area"
          ]
        },
        "userPermsBased": {
          "ownerGroup": {
            "priority": 1,
            "groupName": "Request Form Owners",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "btnSubmit",
              "attachments"
            ]
          },
          "internalUsers": {
            "priority": 2,
            "groupName": "Knowledge Management Internal Users",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "btnSubmit",
              "attachments"
            ]
          },
          "externalUsers": {
            "priority": 3,
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ],
            "defaultEditable": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "btnSubmit",
              "attachments"
            ]
          }
        }
      }
    },
    {
      "formMode": 4,
      "_comment": "VIEW FORM",
      "items": {
        "globalDisableFields": {
          "items": [
            "Business_x0020_Area"
          ]
        },
        "userPermsBased": {
          "ownerGroup": {
            "priority": 1,
            "groupName": "Request Form Owners",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "attachments"
            ]
          },
          "internalUsers": {
            "priority": 2,
            "groupName": "Knowledge Management Internal Users",
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "Status",
              "Access_x0020_Granted",
              "attachments"
            ]
          },
          "externalUsers": {
            "priority": 3,
            "defaultVisible": [
              "Authorized_x0020_Requestor",
              "Cost_x0020_Center",
              "Business_x0020_Area",
              "Document_x0020_Type",
              "Signer",
              "Comments",
              "Cost_x0020_Center_x0020_Owner_x0",
              "attachments"
            ]
          }
        }
      }
    }
  ]
}


