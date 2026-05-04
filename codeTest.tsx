{
  "$schema": "./formRules.schema.json",
  "title": "Fields disabled/hidden or user based permissions for Authorized Requestor list",
  "sourceType": "list",
  "type": "array",
  "items": [
    {
      "formMode": 8,
      "_comment": "New Form",
      "items": {
        "disableFields": {
          "items": []
        },
        "hiddenFields": {
          "items": []
        },
        "userPermsBased": {
          "fieldBasedGeneral": {
            "items": []
          },
          "fieldBasedCurUser": {
            "items": []
          },
          "groupBased": {
            "items": []
          }
        },
        "autoPopulate": {
          "items": [
            {
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
      "_comment": "Edit Form",
      "items": {
        "disableFields": {
          "items": [
            "Title",
            "Department"
          ]
        },
        "hiddenFields": {
          "items": []
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
            "items": []
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
                    "Attachments"
                  ]
                },
                "visibleOnly": {
                  "items": [
                    "Status",
                    "Internal_x0020__x0020_Comments",
                    "btnSubmit",
                    "Attachments"
                  ]
                }
              }
            ]
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
      "_comment": "View Form",
      "items": {
        "disableFields": {
          "items": []
        },
        "hiddenFields": {
          "items": []
        },
        "userPermsBased": {
          "fieldBasedGeneral": {
            "items": []
          },
          "fieldBasedCurUser": {
            "items": []
          },
          "groupBased": {
            "items": []
          }
        }
      }
    }
  ]
}