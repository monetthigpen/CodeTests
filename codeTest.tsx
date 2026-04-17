{
  "$schema": "./formRules.schema.json",
  "title": "Fields disable/hidden or user based permissions for Authorized Requestor list",
  "sourceType": "List",
  "type": "array",
  "items": [
    {
      "formMode": 8,
      "_comment": "Fields configuration for NewForm Mode",
      "items": {
        "disableFields": {
          "items": [
            "Title",
            "Department"
          ],
          "_comment": "List of fields that will be disabled for all users regardless of their permissions"
        },
        "hiddenFields": {
          "items": [
            "Title",
            "Department",
            "Status",
            "Resolution",
            "Internal_x0020__x0020_Comments",
            "Comment_x0020_History",
            "Complete_x0020_Date",
            "RequestTracker"
          ],
          "_comment": "List of fields that will be hidden for all users regardless of their permissions"
        },
        "userPermsBased": {
          "fieldBasedGeneral": {
            "items": [],
            "_comment": "List of fields that are allowed to be edited by all the users but ONLY based on a specific field and its value defined"
          },
          "fieldBasedCurUser": {
            "items": [],
            "_comment": "List of fields that are allowed to be edited by ONLY Current user but based on a specific field value defined and Current User Specific Field Reference"
          },
          "groupBased": {
            "items": []
          }
        }
      }
    },
    {
      "formMode": 6,
      "_comment": "Fields configuration for Edit Mode",
      "items": {
        "disableFields": {
          "items": [
            "Title",
            "Department"
          ],
          "_comment": "List of fields that will be disabled for all users regardless of their permissions"
        },
        "hiddenFields": {
          "items": [],
          "_comment": "List of fields that will be hidden fields for all users regardless of their permissions"
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
          "fieldBasedGeneral": {
            "items": [],
            "_comment": "List of fields that are allowed to be edited by all the users but ONLY based on a specific field and its value defined"
          },
          "fieldBasedCurUser": {
            "items": [
              {
                "fieldSource": {
                  "fieldName": "Status",
                  "fieldValue": "",
                  "referenceCurUsrField": "Cost_x0020_Center_x0020_OwnerId"
                },
                "editableOnly": {
                  "items": [
                    "Status",
                    "Internal_x0020__x0020_Comments",
                    "btnSubmit",
                    "attachments"
                  ]
                },
                "visibleOnly": {
                  "items": [
                    "Status",
                    "Internal_x0020__x0020_Comments",
                    "btnSubmit",
                    "attachments"
                  ]
                }
              }
            ],
            "_comment": "Fields editable by ONLY Current user based on Cost Center Owner people field"
          },
          "groupBased": {
            "items": [
              {
                "groupSource": {
                  "groupName": "Request Form Owners"
                },
                "editableOnly": {
                  "items": [
                    "Status",
                    "Internal_x0020__x0020_Comments",
                    "Complete_x0020_Date"
                  ]
                },
                "visibleOnly": {
                  "items": [
                    "Status",
                    "RequestTracker",
                    "Internal_x0020__x0020_Comments",
                    "Complete_x0020_Date"
                  ]
                }
              },
              {
                "groupSource": {
                  "groupName": "Knowledge Management Internal Users"
                },
                "editableOnly": {
                  "items": [
                    "Status",
                    "Internal_x0020__x0020_Comments",
                    "Complete_x0020_Date"
                  ]
                },
                "visibleOnly": {
                  "items": [
                    "Status",
                    "RequestTracker",
                    "Internal_x0020__x0020_Comments",
                    "Complete_x0020_Date"
                  ]
                }
              }
            ]
          }
        },
        "autoPopulate": {
          "items": [
            {
              "when": {
                "field": "Status",
                "equals": [
                  "Approved",
                  "Rejected"
                ]
              },
              "set": {
                "field": "Complete_x0020_Date",
                "value": "TODAY"
              }
            }
          ]
        }
      }
    },
    {
      "formMode": 4,
      "_comment": "Fields configuration for View Mode",
      "items": {
        "disableFields": {
          "items": [],
          "_comment": "List of fields that will be disabled for all users regardless of their permissions"
        },
        "hiddenFields": {
          "items": [],
          "_comment": "List of fields that will be hidden fields for all users regardless of their permissions"
        },
        "userPermsBased": {
          "fieldBasedGeneral": {
            "items": [],
            "_comment": "List of fields that are allowed to be edited by all the users but ONLY based on a specific field and its value defined"
          },
          "fieldBasedCurUser": {
            "items": [],
            "_comment": "List of fields that are allowed to be edited by ONLY Current user but based on a specific field value defined and Current User Specific Field Reference"
          },
          "groupBased": {
            "items": []
          }
        }
      }
    }
  ]
}