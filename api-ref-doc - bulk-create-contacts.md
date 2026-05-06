# Bulk Create Contacts

Use the Bulk Create Contacts endpoint to create up to 100 contacts in a single API request. This endpoint supports intelligent deduplication and returns separated arrays for newly created and existing contacts. <br><br>Important: This endpoint creates new contacts but does NOT update existing ones (except for placeholder contacts from email imports). Existing contacts that match the criteria will be returned in the existing_contacts array without modification. <br><br>To update existing contacts, use the Bulk Update Contacts endpoint.

# OpenAPI definition

```json
{
  "openapi": "3.1.0",
  "info": {
    "title": "apollo-rest-api",
    "version": "1.0"
  },
  "servers": [
    {
      "url": "https://api.apollo.io/api/v1"
    }
  ],
  "components": {
    "securitySchemes": {
      "apiKey": {
        "type": "apiKey",
        "in": "header",
        "name": "x-api-key",
        "description": "API key"
      },
      "bearerAuth": {
        "type": "http",
        "scheme": "bearer",
        "bearerFormat": "JWT",
        "description": "[Recommended] OAuth Access token"
      }
    }
  },
  "security": [
    {
      "bearerAuth": []
    },
    {
      "apiKey": []
    }
  ],
  "paths": {
    "/contacts/bulk_create": {
      "post": {
        "summary": "Bulk Create Contacts",
        "description": "Use the Bulk Create Contacts endpoint to create up to 100 contacts in a single API request. This endpoint supports intelligent deduplication and returns separate arrays for newly created and existing contacts. <br><br><strong>Important:</strong> This endpoint creates new contacts but does NOT update existing ones (except for placeholder contacts from email imports). Existing contacts that match the criteria will be returned in the existing_contacts array without modification. To update existing contacts, use the <a href=\"https://docs.apollo.io/reference/bulk-update-contacts\" target=\"_blank\">Bulk Update Contacts endpoint</a>. <br><br>The endpoint can operate in two modes: default mode (creates duplicates for non-email_import sources, merges with email_import placeholders only) or full deduplication mode (returns existing contacts without modifying them). <br><br>For creating individual contacts, use the <a href=\"https://docs.apollo.io/reference/create-a-contact\" target=\"_blank\">Create a Contact endpoint</a> instead.",
        "operationId": "bulk-create-contacts",
        "requestBody": {
          "required": true,
          "content": {
            "application/json": {
              "schema": {
                "type": "object",
                "required": [
                  "contacts"
                ],
                "properties": {
                  "contacts": {
                    "type": "array",
                    "description": "Array of contact objects to create (maximum 100 contacts per request)",
                    "maxItems": 100,
                    "items": {
                      "type": "object",
                      "properties": {
                        "first_name": {
                          "type": "string",
                          "description": "Contact's first name",
                          "example": "John"
                        },
                        "last_name": {
                          "type": "string",
                          "description": "Contact's last name",
                          "example": "Doe"
                        },
                        "email": {
                          "type": "string",
                          "description": "Contact's email address",
                          "example": "john.doe@example.com"
                        },
                        "title": {
                          "type": "string",
                          "description": "Contact's job title",
                          "example": "Senior Manager"
                        },
                        "primary_title": {
                          "type": "string",
                          "description": "Primary job title (takes precedence over title)",
                          "example": "VP of Sales"
                        },
                        "organization_name": {
                          "type": "string",
                          "description": "Company/organization name",
                          "example": "Acme Corporation"
                        },
                        "phone": {
                          "type": "string",
                          "description": "Phone number",
                          "example": "+1-555-0100"
                        },
                        "present_raw_address": {
                          "type": "string",
                          "description": "Physical address",
                          "example": "San Francisco, CA"
                        },
                        "linkedin_url": {
                          "type": "string",
                          "description": "LinkedIn profile URL",
                          "example": "https://www.linkedin.com/in/johndoe"
                        },
                        "facebook_url": {
                          "type": "string",
                          "description": "Facebook profile URL",
                          "example": "https://www.facebook.com/johndoe"
                        },
                        "twitter_url": {
                          "type": "string",
                          "description": "Twitter profile URL",
                          "example": "https://twitter.com/johndoe"
                        },
                        "photo_url": {
                          "type": "string",
                          "description": "Profile photo URL",
                          "example": "https://example.com/photo.jpg"
                        },
                        "account_id": {
                          "type": "string",
                          "description": "Associated account ID",
                          "example": "507f1f77bcf86cd799439011"
                        },
                        "organization_id": {
                          "type": "string",
                          "description": "Associated organization ID",
                          "example": "507f1f77bcf86cd799439012"
                        },
                        "owner_id": {
                          "type": "string",
                          "description": "Contact owner user ID (defaults to current user if not provided)",
                          "example": "507f1f77bcf86cd799439013"
                        },
                        "contact_stage_id": {
                          "type": "string",
                          "description": "Contact stage ID",
                          "example": "507f1f77bcf86cd799439014"
                        },
                        "salesforce_id": {
                          "type": "string",
                          "description": "Salesforce ID for matching and deduplication",
                          "example": "003xx000004TmiQAAS"
                        },
                        "hubspot_id": {
                          "type": "string",
                          "description": "HubSpot ID for matching and deduplication",
                          "example": "12345678"
                        },
                        "salesforce_lead_id": {
                          "type": "string",
                          "description": "Salesforce Lead ID",
                          "example": "00Qxx000001abcDEFG"
                        },
                        "salesforce_contact_id": {
                          "type": "string",
                          "description": "Salesforce Contact ID for matching",
                          "example": "003xx000004TmiQAAS"
                        },
                        "salesforce_account_id": {
                          "type": "string",
                          "description": "Salesforce Account ID",
                          "example": "001xx000003DGb2AAG"
                        },
                        "outreach_id": {
                          "type": "string",
                          "description": "Outreach.io ID",
                          "example": "98765"
                        },
                        "salesloft_id": {
                          "type": "string",
                          "description": "SalesLoft ID",
                          "example": "54321"
                        },
                        "phone_status_cd": {
                          "type": "string",
                          "description": "Phone validation status"
                        },
                        "typed_custom_fields": {
                          "type": "object",
                          "description": "Custom field values as key-value pairs where key is the field_id and value is the field_value",
                          "additionalProperties": {
                            "type": "string"
                          },
                          "example": {
                            "60c39ed82bd02f01154c470a": "2025-08-07"
                          }
                        },
                        "contact_emails": {
                          "type": "array",
                          "description": "Array of email objects with position",
                          "items": {
                            "type": "object",
                            "properties": {
                              "email": {
                                "type": "string",
                                "example": "john.doe@example.com"
                              },
                              "position": {
                                "type": "integer",
                                "example": 0
                              }
                            }
                          }
                        },
                        "phone_numbers": {
                          "type": "array",
                          "description": "Array of phone number objects",
                          "items": {
                            "type": "object",
                            "properties": {
                              "raw_number": {
                                "type": "string",
                                "example": "+1-555-0100"
                              },
                              "position": {
                                "type": "integer",
                                "example": 0
                              }
                            }
                          }
                        },
                        "contact_role_type_ids": {
                          "type": "array",
                          "description": "Array of contact role type IDs",
                          "items": {
                            "type": "string"
                          },
                          "example": [
                            "507f1f77bcf86cd799439020"
                          ]
                        }
                      }
                    }
                  },
                  "append_label_names": {
                    "type": "array",
                    "description": "Array of label names to add to ALL contacts in this request",
                    "items": {
                      "type": "string"
                    },
                    "example": [
                      "Hot Lead",
                      "Q1 2024"
                    ]
                  },
                  "run_dedupe": {
                    "type": "boolean",
                    "description": "Enable full deduplication across all sources. When false (default), creates duplicates for non-email_import sources and merges with email_import placeholders only. When true, returns existing contacts without modifying them (except email_import placeholders which are still merged). Matches by email, CRM IDs, or name + organization",
                    "default": false,
                    "example": true
                  }
                }
              },
              "examples": {
                "Basic Contact Creation": {
                  "value": {
                    "contacts": [
                      {
                        "first_name": "John",
                        "last_name": "Doe",
                        "email": "john.doe@example.com",
                        "title": "Senior Manager",
                        "organization_name": "Acme Corporation"
                      },
                      {
                        "first_name": "Jane",
                        "last_name": "Smith",
                        "email": "jane.smith@techstart.io",
                        "title": "VP of Sales",
                        "organization_name": "TechStart Inc",
                        "linkedin_url": "https://www.linkedin.com/in/janesmith"
                      }
                    ]
                  }
                },
                "With Full Deduplication (Returns Existing Without Updating)": {
                  "value": {
                    "contacts": [
                      {
                        "email": "existing@example.com",
                        "first_name": "New",
                        "last_name": "Name",
                        "title": "New Title"
                      }
                    ],
                    "run_dedupe": true
                  }
                },
                "With CRM IDs (Returns Existing Without Updating)": {
                  "value": {
                    "contacts": [
                      {
                        "salesforce_contact_id": "003xx000004TmiQAAS",
                        "first_name": "John",
                        "last_name": "Doe",
                        "email": "john.doe@example.com",
                        "phone": "+1-555-0200"
                      },
                      {
                        "hubspot_id": "12345678",
                        "first_name": "Jane",
                        "last_name": "Smith",
                        "email": "jane.smith@example.com"
                      }
                    ]
                  }
                },
                "With Custom Fields and Labels": {
                  "value": {
                    "contacts": [
                      {
                        "first_name": "Sarah",
                        "last_name": "Johnson",
                        "email": "sarah.johnson@enterprise.com",
                        "title": "CTO",
                        "organization_name": "Enterprise Corp",
                        "contact_stage_id": "507f1f77bcf86cd799439014",
                        "typed_custom_fields": {
                          "60c39ed82bd02f01154c470a": "2025-12-31",
                          "60c39ed82bd02f01154c470b": "High Priority"
                        }
                      }
                    ],
                    "append_label_names": [
                      "VIP",
                      "Decision Maker"
                    ]
                  }
                },
                "With Multiple Contact Details": {
                  "value": {
                    "contacts": [
                      {
                        "first_name": "Michael",
                        "last_name": "Brown",
                        "organization_name": "Tech Solutions",
                        "contact_emails": [
                          {
                            "email": "michael.brown@techsolutions.com",
                            "position": 0
                          },
                          {
                            "email": "mbrown@personal.com",
                            "position": 1
                          }
                        ],
                        "phone_numbers": [
                          {
                            "raw_number": "+1-555-0100",
                            "position": 0
                          },
                          {
                            "raw_number": "+1-555-0101",
                            "position": 1
                          }
                        ]
                      }
                    ]
                  }
                }
              }
            }
          }
        },
        "deprecated": false
      }
    }
  },
  "x-readme": {
    "headers": [
      {
        "key": "Cache-Control",
        "value": "no-cache"
      },
      {
        "key": "Content-Type",
        "value": "application/json"
      }
    ],
    "explorer-enabled": true,
    "proxy-enabled": true
  },
  "x-readme-fauxas": true
}
```