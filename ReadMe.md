# Microsoft 365 Copilot Connectors — Best Practices & Development Guide

> **Audience:** Developers building custom Microsoft 365 Copilot Connectors (formerly Microsoft Graph Connectors)  
> **Last updated:** March 2026  
> **Sources:** Official Microsoft Learn documentation, Microsoft Graph API reference

---

## Table of Contents

1. [Introduction & Overview](#1-introduction--overview)
2. [Choosing the Right Tools](#2-choosing-the-right-tools)
3. [Schema Design Best Practices](#3-schema-design-best-practices)
4. [Data Format & the Content Property](#4-data-format--the-content-property)
5. [Types of Data & Enterprise Scenarios](#5-types-of-data--enterprise-scenarios)
6. [Data Chunking & Indexing Strategies](#6-data-chunking--indexing-strategies)
7. [Handling Large Payloads](#7-handling-large-payloads)
8. [Handling Delimited Data (CSV / TSV / PSV)](#8-handling-delimited-data-csv--tsv--psv)
9. [Data Aggregation Across Multiple Items](#9-data-aggregation-across-multiple-items)
10. [Surfacing Data in Microsoft 365 Copilot](#10-surfacing-data-in-microsoft-365-copilot)
11. [Access Control & Security](#11-access-control--security)
12. [Monitoring, Testing & Troubleshooting](#12-monitoring-testing--troubleshooting)
13. [Quick-Reference Checklists](#13-quick-reference-checklists)
14. [Appendix A: Cross-Reference Reconciliation](#appendix-a-cross-reference-reconciliation-gpt-54--codex-53)
15. [Appendix B: Key Links & References](#appendix-b-key-links--references)

---

## 1. Introduction & Overview

### What Are Copilot Connectors?

Microsoft 365 Copilot Connectors (formerly Microsoft Graph Connectors) bring external data into the Microsoft Graph semantic index. Once indexed, this data becomes discoverable and usable across the Microsoft 365 ecosystem:

| Experience | What It Enables | Status |
|---|---|---|
| **Microsoft 365 Copilot** | Natural-language Q&A, summarization, and reasoning over your external data | GA |
| **Microsoft Search** | Full-text and semantic search in Office.com, SharePoint, Bing at Work | GA |
| **Context IQ** | In-context suggestions in Outlook on the web when composing emails | Preview |

> 📌 **People connectors** are a distinct connector pattern with profile-card integration, tenant-wide visibility, and information-barrier considerations. If you are indexing employee/people profile data, refer to the [People connectors documentation](https://learn.microsoft.com/graph/connecting-external-content-experiences) for specific requirements and limitations.

### How It Works

```
┌──────────────┐     ┌───────────────────┐     ┌──────────────────────┐
│ Your Data    │────▶│ Copilot Connector │────▶│ Microsoft Graph      │
│ Source       │     │ (SDK / REST API)  │     │ Semantic Index       │
└──────────────┘     └───────────────────┘     └──────────┬───────────┘
                                                          │
                              ┌────────────────────────────┤
                              ▼                            ▼
                     ┌─────────────────┐       ┌───────────────────┐
                     │ M365 Copilot    │       │ Microsoft Search  │
                     │ (Chat, BizChat) │       │ (SharePoint, Bing)│
                     └─────────────────┘       └───────────────────┘
```

1. **Create a connection** — Register your external data source with Microsoft Graph.
2. **Register a schema** — Define the properties, data types, and search attributes.
3. **Ingest items** — Push external items (with ACLs, properties, and content) into the index.
4. **Configure experiences** — Enable search verticals, result types, and Copilot inline results.

### Licensing & Item Quota

Items ingested through connectors consume your organization's **item quota**. Every Microsoft 365 license that includes Microsoft Search also includes connector item quota. Additional quota can be purchased through the **Microsoft Graph Connectors Capacity** add-on. Check your current usage in the M365 Admin Center under **Search & intelligence > Connectors**.

### Prerequisites

- Microsoft 365 tenant with appropriate licenses
- Application registration in Microsoft Entra ID with the following permissions:
  - **Connection & schema management**: `ExternalConnection.ReadWrite.OwnedBy` (or `.All` for broader access)
  - **Item ingestion**: `ExternalItem.ReadWrite.OwnedBy` (or `.All`)
  - Both permissions are typically needed for end-to-end connector development
- Admin consent for the application
- For SDK-based connectors: Microsoft Graph Connector Agent installed on a Windows machine

---

## 2. Choosing the Right Tools

### Tool Comparison Matrix

| Tool | Language | Best For | Complexity | On-Prem Support |
|---|---|---|---|---|
| **Microsoft 365 Agents Toolkit** | TypeScript/C# | New custom connectors with integrated agent development | Medium | ❌ Cloud-only |
| **Microsoft Graph REST API** | Any (HTTP) | Maximum flexibility, polyglot teams, serverless architectures | Medium-High | ✅ If ingestion app runs on-prem/hybrid |
| **Federated connectors** (Preview) | Any (HTTP) | Real-time, non-indexed access to live external data | Medium | Varies |
| **Pre-built connectors** | No code | 100+ supported data sources (ServiceNow, Salesforce, Jira, etc.) | Low | Varies |
| **Power Platform connectors** | Low-code | Connector *actions* and extensibility (not primary content ingestion) | Low | Via on-prem gateway |
| **Copilot Connectors SDK** | C# (primary), any gRPC language | Production-grade custom connectors with full crawl management | Medium | ✅ Via connector agent |
| **Graph Connector Agent** | N/A (agent Windows Only) | On-premises data sources, bridge to cloud indexing | Low-Medium | ✅ Primary purpose |


> 📌 **Synced vs. Federated connectors**: Synced connectors (SDK, REST API) **index** content into Microsoft Graph for offline search and Copilot reasoning. Federated connectors (preview) provide **live, read-only** access without indexing — ideal for sensitive, fast-changing, or aggregation-heavy data where pre-computation is impractical.

### Decision Flowchart

```
Is there a pre-built connector for your data source?
├── YES → Use the pre-built connector (M365 Admin Center)
└── NO
    ├── Do you need indexed content or real-time live access?
    │   ├── LIVE ACCESS → Use Federated Connectors (Preview)
    │   └── INDEXED CONTENT
    │       ├── Is your data source on-premises?
    │       │   ├── YES → Use the Copilot Connectors SDK + Connector Agent
    │       │   └── NO
    │       │       ├── Do you need full crawl management (scheduling, dedup, change detection)?
    │       │       │   ├── YES → Use the Copilot Connectors SDK or Agents Toolkit
    │       │       │   └── NO → Use the Microsoft Graph REST API directly
    │       │       └── Are you building a Declarative Agent with connector integration?
    │       │           └── YES → Consider Microsoft 365 Agents Toolkit
```

### Copilot Connectors SDK — Key Components

1. **Custom connector template** — Download from Visual Studio Marketplace; provides scaffolding for your C# connector project.
2. **gRPC contracts** — Protocol buffer files defining the interface between the connector agent and your code. Available in [GitHub](https://github.com/microsoftgraph/msgraph-connectors-sdk) for non-C# languages.
3. **Microsoft Graph connector agent** — Lightweight Windows service that orchestrates crawls, manages ingestion into Microsoft Graph, and handles identity mapping. [Download latest version](https://aka.ms/gca/).
4. **SDK test utility** — Pre-built test scenarios to validate your connector code before deployment.

### Agent Capabilities

The connector agent handles:

- **Crawl scheduling**: Full crawls and incremental crawls at configurable intervals
- **Change detection**: Hash-based comparison to only re-index changed items
- **Duplicate/cycle detection**: Skips previously-seen or circular references (especially for web-based sources)
- **Identity mapping**: Stamps ACLs using Microsoft Entra ID or external group mappings
- **Microsoft Graph ingestion**: Handles the actual HTTP calls to push items into the index

### Using the Graph REST API Directly

If you prefer direct API control (e.g., running in Azure Functions, AWS Lambda, or any cloud platform):

```http
PUT /external/connections/{connectionId}/items/{itemId}
Content-Type: application/json

{
  "acl": [...],
  "properties": {...},
  "content": {
    "value": "Your content text here",
    "type": "text"
  }
}
```

This approach gives you full flexibility but requires you to implement your own:
- Crawl scheduling and orchestration
- Change detection and delta handling
- Retry logic and throttle management
- ACL resolution

---

## 3. Schema Design Best Practices

The schema is a flat list of properties that defines the structure of your external items. It determines how your data is searched, filtered, displayed, and understood by Copilot.

### Schema Limits

| Limit | Value |
|---|---|
| Maximum properties per schema | **128** |
| Maximum external groups per tenant | 100,000 |
| Maximum external groups per user (for search) | 10,000 |

### Property Naming Conventions

| ✅ Do | ❌ Don't |
|---|---|
| `parentOrganizationName` | `orgName` |
| `incidentRootCause` | `dataBlob` |
| `qualifiedSalesLead` | `ftxInvIsLead` |
| `departmentName` | `brOrgName` |

**Rules:**
- Use **clear, descriptive names** that convey semantic meaning
- Avoid abbreviations, acronyms, or cryptic identifiers
- Add **property descriptions** to help Copilot and Declarative Agents understand the semantic meaning of each property. The `description` field is available on schema properties via the Graph API — use it for all properties, especially queryable ones.
- Property names help Copilot understand and match content to user queries

### Data Types

| Type | Use For | Example Properties |
|---|---|---|
| `String` | Free-text values, identifiers, names | `title`, `description`, `assignee` |
| `Int64` | Whole numbers, counts, priorities | `priority`, `itemCount`, `severity` |
| `Double` | Decimal numbers, scores, percentages | `score`, `price`, `completionRate` |
| `DateTime` | Timestamps | `createdDate`, `lastModified`, `dueDate` |
| `Boolean` | True/false flags | `isResolved`, `isActive`, `isPublic` |
| `StringCollection` | Multi-value text (tags, categories) | `tags`, `categories`, `skills` |
| `Int64Collection` | Multi-value integers | `relatedItemIds`, `scores` |
| `DoubleCollection` | Multi-value decimals | `measurements`, `ratings` |
| `DateTimeCollection` | Multi-value timestamps | `milestones`, `reviewDates` |

> 📌 The Graph API also supports `principal` and `principalCollection` types for representing Microsoft Entra ID users and groups directly as property values. Use these when a property inherently represents an identity (e.g., `approver`, `teamMembers`) instead of storing raw email addresses as strings.

### Schema Attribute Configuration

Each property can be configured with the following attributes. Choose carefully — they directly impact search behavior and performance.

#### Searchable

When marked searchable, the property value is added to the **full-text index**.

- ✅ Mark searchable: `title`, `description`, `tags`, `createdBy`, `assignedTo`
- ❌ Don't mark searchable: Large binary fields, refinable fields (mutually exclusive)
- 📌 **Critical for Copilot**: The `searchable` attribute defines which properties Copilot can match against in user prompts

#### Queryable

Enables filtering with **KQL (Keyword Query Language)** in search queries.

- ✅ Mark queryable: `status`, `priority`, `assignedTo`, `category`, `ticketId`
- ❌ Don't mark queryable: Large text fields like `description`
- 💡 Combine with `retrievable: true` so filtered properties appear in results
- 💡 Supports prefix matching with wildcard `*` (suffix matching not supported)

#### Retrievable

The property value is returned in search results and available for display templates.

- ✅ Mark retrievable: `title`, `summary`, `status`, `assignedTo`, `createdDateTime`
- ❌ Don't over-mark: Too many or large retrievable properties increase search latency
- 📌 **Required** for properties mapped to semantic labels

#### Refinable

The property appears as a filter control (dropdown, checkbox) in Microsoft Search.

- ✅ Mark refinable: `tags`, `status`, `priority`, `category`, `type`
- ⚠️ **Mutually exclusive with searchable** — a property cannot be both
- ⚠️ Only `String`, numeric, and `DateTime` types can be refinable
- ⚠️ Too many refinable properties impact performance

#### ExactMatchRequired

When `true`, the full string value is indexed without tokenization.

- ✅ Use for: GUIDs, ticket IDs, SKUs, part numbers
- ⚠️ Can only be applied to properties that are **not searchable**
- Example: `ticketId:CTS-ce913b61` matches only `CTS-ce913b61`, not `CTS`

### Semantic Labels

Semantic labels are the **most important schema feature for Copilot integration**. They tell Microsoft 365 the semantic role of each property.

| Label | Description | Map To Properties Like |
|---|---|---|
| `title` | Main name/heading of the item | `documentTitle`, `ticketSubject` |
| `url` | Direct link to open the item | `documentLink`, `itemUrl` |
| `createdBy` | User who created the item | `authorEmail`, `submittedBy` |
| `lastModifiedBy` | User who last edited the item | `editorEmail`, `updatedBy` |
| `authors` | All collaborators on the item | `authorName`, `contributors` |
| `createdDateTime` | When the item was created | `createdOn`, `submissionDate` |
| `lastModifiedDateTime` | When the item was last modified | `lastUpdated`, `modifiedOn` |
| `fileName` | Name of the file | `documentName`, `attachmentName` |
| `fileExtension` | File extension | `documentType`, `format` |
| `iconUrl` | URL of an icon/thumbnail | `thumbnailUrl`, `logo` |
| `containerName` | Name of the parent container | `projectName`, `folderName` |
| `containerUrl` | URL of the parent container | `projectUrl`, `folderLink` |

> 📌 **Additional labels exist** beyond those listed above, including `assignedTo`, `dueDate`, `closedDate`, `closedBy`, `reportedBy`, `sprintName`, `severity`, `state`, `priority`, `secondaryId`, `itemParentId`, `parentUrl`, `tags`, `itemType`, `itemPath`, and `numReactions`. People connectors have additional people-specific labels. Always check the [current semantic labels documentation](https://learn.microsoft.com/graph/connecting-external-content-manage-schema#semantic-labels) for the latest list.

**Best Practices:**
- **Always assign `title`, `url`, and `iconUrl`** — these are required for content to surface in Copilot
- Map as many labels as are relevant and accurate
- Do **not** assign a label if it doesn't match the property's actual purpose
- Properties must be marked **retrievable** before they can receive labels
- Labels listed in order of impact on discovery: `title` → `lastModifiedDateTime` → `lastModifiedBy` → `url` → `fileName` → `fileExtension`

### Aliases

Aliases are friendly names that users can use in search queries.

| Property | Suggested Aliases |
|---|---|
| `createdBy` | `author`, `owner`, `submittedBy` |
| `title` | `subject`, `heading` |
| `tags` | `labels`, `categories` |
| `fileName` | `documentName`, `file` |
| `summary` | `description`, `abstract` |

**Tips:**
- Use aliases for common synonyms and domain-specific terminology
- Keep aliases short and intuitive
- Avoid overly generic or ambiguous aliases

### Example Schema: Work Ticket System

```json
{
  "baseType": "microsoft.graph.externalItem",
  "properties": [
    {
      "name": "ticketId",
      "type": "String",
      "isSearchable": false,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": false,
      "isExactMatchRequired": true,
      "aliases": ["ID"]
    },
    {
      "name": "title",
      "type": "String",
      "isSearchable": true,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": false,
      "labels": ["title"]
    },
    {
      "name": "description",
      "type": "String",
      "isSearchable": true,
      "isQueryable": false,
      "isRetrievable": false,
      "isRefinable": false
    },
    {
      "name": "status",
      "type": "String",
      "isSearchable": false,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": true,
      "aliases": ["state"]
    },
    {
      "name": "priority",
      "type": "Int64",
      "isSearchable": false,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": true
    },
    {
      "name": "assignedTo",
      "type": "String",
      "isSearchable": true,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": false,
      "aliases": ["assignee", "owner"]
    },
    {
      "name": "createdDate",
      "type": "DateTime",
      "isSearchable": false,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": true,
      "labels": ["createdDateTime"]
    },
    {
      "name": "lastModifiedDate",
      "type": "DateTime",
      "isSearchable": false,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": true,
      "labels": ["lastModifiedDateTime"]
    },
    {
      "name": "tags",
      "type": "StringCollection",
      "isSearchable": false,
      "isQueryable": true,
      "isRetrievable": true,
      "isRefinable": true,
      "isExactMatchRequired": true,
      "aliases": ["labels", "categories"]
    },
    {
      "name": "itemUrl",
      "type": "String",
      "isSearchable": false,
      "isQueryable": false,
      "isRetrievable": true,
      "isRefinable": false,
      "labels": ["url"]
    }
  ]
}
```

---

## 4. Data Format & the Content Property

### Understanding the Content Property

The `content` property is a **built-in, default property** in every Copilot Connector schema. Unlike other properties, you do **not** define it in the schema — it is included directly in the item payload during ingestion.

The content property is:
- **Semantically indexed** for full-text search
- Used to generate **dynamic snippets** in search results
- Available to Copilot for **summarization and semantic reasoning**

### Supported Content Types

| Type | Format | When to Use |
|---|---|---|
| `text` | Plain text, no markup | Simple data, compliance scenarios, structured records, short content |
| `html` | HTML markup | Rich documents, formatted content, content with headings/lists/tables |

> ⚠️ **Markdown is NOT a supported content type.** If your source data is in Markdown, you must either:
> - **Convert it to HTML** before ingestion (recommended for rich content)
> - **Strip it to plain text** (acceptable for simpler content)

> ⚠️ **For compliance-enabled connections** (where `enabledContentExperience` is set to `compliance`), you **must** use `text` — HTML is not supported for compliance scenarios.

### Content Format Best Practices

#### When to Use `text`

```json
{
  "content": {
    "value": "Payment gateway error. Root cause: SSL certificate expired on gateway-prod-03. Resolution: Certificate renewed and services restarted. Downtime: 45 minutes.",
    "type": "text"
  }
}
```

Use plain text when:
- Content is short and factual (< 1,000 words)
- No structural formatting adds value
- Data comes from structured records (database rows, API responses)
- Compliance features are enabled on the connection

#### When to Use `html`

```json
{
  "content": {
    "value": "<html><body><h1>Quarterly Sales Report</h1><h2>Executive Summary</h2><p>Q3 revenue increased by 15% year-over-year, driven primarily by enterprise segment growth.</p><h2>Key Metrics</h2><ul><li>Total Revenue: $12.4M</li><li>New Customers: 147</li><li>Churn Rate: 2.1%</li></ul><h2>Regional Breakdown</h2><table><tr><th>Region</th><th>Revenue</th></tr><tr><td>North America</td><td>$7.2M</td></tr><tr><td>EMEA</td><td>$3.8M</td></tr><tr><td>APAC</td><td>$1.4M</td></tr></table></body></html>",
    "type": "html"
  }
}
```

Use HTML when:
- Content has inherent structure (headings, lists, tables)
- Documents are long-form (articles, reports, wiki pages)
- Preserving structure improves Copilot's comprehension
- Content was originally in a rich format (Word, HTML pages, wikis)

#### HTML Best Practices

- Use **semantic HTML tags**: `<h1>`-`<h6>`, `<p>`, `<ul>`, `<ol>`, `<table>`, `<strong>`, `<em>`
- **Strip unnecessary elements**: Remove `<script>`, `<style>`, navigation, footers, ads
- **Keep it clean**: No inline CSS, no class attributes, no data attributes — focus on content structure
- **Tables**: Use `<table>` with `<th>` and `<td>` for tabular data — Copilot can reason over well-structured tables
- **Avoid deeply nested HTML**: Flatten when possible to reduce token overhead

### Structuring Content for Copilot Summarization

Copilot performs best when content is **information-dense and well-organized**:

1. **Lead with the most important information** — Place key facts, summaries, and conclusions at the beginning of the content
2. **Use descriptive headings** — Help Copilot identify sections and locate relevant information
3. **Include contextual metadata inline** — If a property is important for understanding, include it in the content too
4. **Concatenate related text fields** — Merge `summary`, `description`, `rootCause`, `resolution` into a single content value

#### Content Concatenation Pattern

```csharp
// Build rich content from multiple source fields
var contentBuilder = new StringBuilder();
contentBuilder.AppendLine($"Title: {ticket.Title}");
contentBuilder.AppendLine($"Status: {ticket.Status} | Priority: {ticket.Priority}");
contentBuilder.AppendLine($"Assigned to: {ticket.AssignedTo}");
contentBuilder.AppendLine();
contentBuilder.AppendLine($"Description: {ticket.Description}");

if (!string.IsNullOrEmpty(ticket.RootCause))
    contentBuilder.AppendLine($"Root Cause: {ticket.RootCause}");

if (!string.IsNullOrEmpty(ticket.Resolution))
    contentBuilder.AppendLine($"Resolution: {ticket.Resolution}");

// Add comments for additional context
foreach (var comment in ticket.Comments.OrderByDescending(c => c.Date))
    contentBuilder.AppendLine($"[{comment.Author} - {comment.Date:yyyy-MM-dd}]: {comment.Text}");

var externalItem = new ExternalItem
{
    Content = new ExternalItemContent
    {
        Value = contentBuilder.ToString(),
        Type = ExternalItemContentType.Text
    },
    Properties = new Properties
    {
        AdditionalData = new Dictionary<string, object>
        {
            { "title", ticket.Title },
            { "status", ticket.Status },
            { "priority", ticket.Priority },
            { "assignedTo", ticket.AssignedTo }
        }
    }
};
```

> 💡 **Key principle**: Put unstructured, searchable text in `content`. Keep structured, filterable values as separate schema properties. Duplicate into `content` only what helps Copilot understand the item contextually.

---

## 5. Types of Data & Enterprise Scenarios

### Scenario Reference Guide

#### 5.1 Knowledge Bases & Wiki Articles

**Examples:** Confluence pages, SharePoint wikis, internal documentation, FAQs

| Schema Property | Type | Attributes | Label |
|---|---|---|---|
| `title` | String | Searchable, Queryable, Retrievable | `title` |
| `articleUrl` | String | Retrievable | `url` |
| `author` | String | Searchable, Queryable, Retrievable | `createdBy` |
| `lastEditor` | String | Queryable, Retrievable | `lastModifiedBy` |
| `lastModified` | DateTime | Queryable, Retrievable, Refinable | `lastModifiedDateTime` |
| `space` | String | Queryable, Retrievable, Refinable | `containerName` |
| `tags` | StringCollection | Queryable, Retrievable, Refinable | — |

**Content strategy:** Use `html` type. Preserve headings, lists, and tables. Strip navigation chrome, sidebars, and boilerplate.

#### 5.2 Tickets & Work Items

**Examples:** ServiceNow incidents, Jira issues, Azure DevOps work items, Zendesk tickets

| Schema Property | Type | Attributes | Label |
|---|---|---|---|
| `ticketId` | String | Queryable, ExactMatch | — |
| `title` | String | Searchable, Queryable, Retrievable | `title` |
| `status` | String | Queryable, Retrievable, Refinable | — |
| `priority` | Int64 | Queryable, Retrievable, Refinable | — |
| `assignedTo` | String | Searchable, Queryable, Retrievable | — |
| `createdBy` | String | Searchable, Queryable, Retrievable | `createdBy` |
| `createdDate` | DateTime | Queryable, Retrievable, Refinable | `createdDateTime` |
| `lastModified` | DateTime | Queryable, Retrievable | `lastModifiedDateTime` |
| `itemUrl` | String | Retrievable | `url` |

**Content strategy:** Use `text` type. Concatenate description + root cause + resolution + recent comments. Prefix each section with a label.

#### 5.3 CRM Records

**Examples:** Salesforce opportunities, Dynamics 365 leads, HubSpot contacts

| Schema Property | Type | Attributes | Label |
|---|---|---|---|
| `accountName` | String | Searchable, Queryable, Retrievable | `title` |
| `contactEmail` | String | Queryable, Retrievable, ExactMatch | — |
| `dealStage` | String | Queryable, Retrievable, Refinable | — |
| `dealValue` | Double | Queryable, Retrievable | — |
| `industry` | String | Queryable, Retrievable, Refinable | — |
| `lastActivity` | DateTime | Queryable, Retrievable | `lastModifiedDateTime` |
| `recordUrl` | String | Retrievable | `url` |

**Content strategy:** Use `text`. Include recent activity notes, deal history, and key contact information in content.

#### 5.4 HR & People Data

**Examples:** Employee profiles, skills directories, org chart data

> ⚠️ **People connectors** are a distinct connector type with specific behavior: profile-card integration, tenant-wide visibility, read-only profiles, precedence rules, and information-barrier considerations. If you are building a connector specifically to surface employee profile data in people cards, refer to the [People connectors documentation](https://learn.microsoft.com/graph/connecting-external-content-experiences). The schema below applies to general HR/people data indexed as standard external items.

| Schema Property | Type | Attributes | Label |
|---|---|---|---|
| `employeeName` | String | Searchable, Queryable, Retrievable | `title` |
| `department` | String | Queryable, Retrievable, Refinable | `containerName` |
| `jobTitle` | String | Searchable, Queryable, Retrievable | — |
| `skills` | StringCollection | Queryable, Retrievable, Refinable | — |
| `location` | String | Queryable, Retrievable, Refinable | — |
| `profileUrl` | String | Retrievable | `url` |

**Content strategy:** Use `text`. Include bio, skills summary, project history, and certifications.

#### 5.5 Financial & Compliance Records

**Examples:** Audit reports, regulatory filings, policy documents, SOX controls

| Schema Property | Type | Attributes | Label |
|---|---|---|---|
| `documentTitle` | String | Searchable, Queryable, Retrievable | `title` |
| `regulatoryBody` | String | Queryable, Retrievable, Refinable | — |
| `complianceStatus` | String | Queryable, Retrievable, Refinable | — |
| `effectiveDate` | DateTime | Queryable, Retrievable, Refinable | `createdDateTime` |
| `documentUrl` | String | Retrievable | `url` |

**Content strategy:** Use `text` (required for compliance). Include full document text. Consider chunking for long regulatory documents.

> ⚠️ **Compliance note:** If you enable the connection for compliance (`enabledContentExperience: "compliance"`), you **must** use `text` as the content type.

#### 5.6 Product Catalogs

**Examples:** SKUs, product specifications, pricing sheets

| Schema Property | Type | Attributes | Label |
|---|---|---|---|
| `productName` | String | Searchable, Queryable, Retrievable | `title` |
| `sku` | String | Queryable, ExactMatch | — |
| `category` | String | Queryable, Retrievable, Refinable | — |
| `price` | Double | Queryable, Retrievable | — |
| `availability` | String | Queryable, Retrievable, Refinable | — |
| `productUrl` | String | Retrievable | `url` |

**Content strategy:** Use `text` or `html`. Include full product description, specifications table, and compatibility notes.

#### 5.7 Custom LOB Applications

**Examples:** ERP records, inventory systems, project management tools, custom databases

Design your schema to mirror the most-queried fields in your LOB system. Use the content property for free-text descriptions and notes. Apply semantic labels wherever a natural mapping exists.

#### 5.8 File Repositories (Parsed Content)

**Examples:** Network file shares, S3 buckets, FTP servers, document management systems

For binary files (PDF, DOCX, PPTX, images):
1. **Parse to text** before ingestion — use libraries like Apache Tika, iTextSharp, or Azure AI Document Intelligence
2. **Apply OCR** for scanned documents and images
3. Set `fileName` and `fileExtension` semantic labels
4. Use `html` content type if the parsed output preserves structure

---

## 6. Data Chunking & Indexing Strategies

### The 4 MB Item Size Limit

Each external item's request body (the full PUT payload including ACL, properties, and content) is limited to **4 MB**. This translates to approximately:

- **600,000–700,000 words** of plain text
- **~1,400 pages** (at 500 words per page)

> 📌 The 4 MB limit refers to **parsed text content**, which is typically ~10% of the original file size for common formats (DOCX, PPTX, PDF).

### When Do You Need to Chunk?

| Scenario | Chunking Needed? |
|---|---|
| Short records (tickets, CRM entries, product listings) | ❌ No — each record is one item |
| Medium articles (wiki pages, FAQs, 1–10 pages) | ❌ Usually no |
| Long documents where **serialized request body exceeds ~3.5 MB** | ✅ Yes |
| Entire databases/datasets | ✅ Yes — one item per row/record |
| Large binary files (parsed to text) | Measure first — serialize to JSON and check byte size |

> 📌 **Base your chunking decision on serialized request-body size, not page counts.** The 4 MB limit refers to parsed text content, which is typically ~10% of the original file size. A 200-page PDF may only produce 200 KB of text. Always measure `Encoding.UTF8.GetByteCount(json)` on the serialized payload before assuming you need to chunk.

### Chunking Strategies

#### Strategy 1: Logical Section Chunking (Recommended)

Split at natural document boundaries: chapters, sections, or headings.

```
Original: 200-page technical manual
├── Chunk 1: "Chapter 1: Introduction" (pages 1–15)
├── Chunk 2: "Chapter 2: Installation" (pages 16–40)
├── Chunk 3: "Chapter 3: Configuration" (pages 41–80)
├── Chunk 4: "Chapter 4: API Reference" (pages 81–150)
└── Chunk 5: "Chapter 5: Troubleshooting" (pages 151–200)
```

**Advantages:** Preserves semantic coherence; each chunk is self-contained and meaningful.

#### Strategy 2: Fixed-Size Chunking with Overlap

Split at fixed character/token boundaries with overlap to preserve context across boundaries.

```
Chunk size: 50,000 characters
Overlap: 5,000 characters

Chunk 1: characters 0–50,000
Chunk 2: characters 45,000–95,000
Chunk 3: characters 90,000–140,000
...
```

**Advantages:** Predictable sizing; ensures no content falls through boundary cracks.  
**Disadvantages:** May split mid-sentence or mid-paragraph; overlap increases total indexed volume.

#### Strategy 3: Semantic Boundary Chunking

Split at paragraph or sentence boundaries, accumulating content until approaching the size limit.

```python
def chunk_by_paragraphs(text, max_size=3_500_000):  # Leave room for metadata
    paragraphs = text.split('\n\n')
    chunks, current_chunk = [], []
    current_size = 0
    
    for para in paragraphs:
        para_size = len(para.encode('utf-8'))
        if current_size + para_size > max_size and current_chunk:
            chunks.append('\n\n'.join(current_chunk))
            current_chunk = []
            current_size = 0
        current_chunk.append(para)
        current_size += para_size
    
    if current_chunk:
        chunks.append('\n\n'.join(current_chunk))
    return chunks
```

### Linking Chunks Together

When you split a document into multiple items, maintain relationships with properties:

```json
{
  "properties": {
    "title": "Technical Manual — Chapter 3: Configuration",
    "parentDocumentId": "DOC-12345",
    "parentDocumentTitle": "Technical Manual v4.2",
    "chunkIndex": 3,
    "totalChunks": 5,
    "sectionPath": "Chapter 3 > Configuration > Network Settings"
  }
}
```

### Making Chunks Self-Contained

Each chunk should be independently understandable by Copilot. Add a **contextual header** to every chunk:

```
Document: Technical Manual v4.2
Section: Chapter 3 — Configuration > Network Settings
Part 3 of 5

[Actual section content begins here...]
```

This ensures Copilot can correctly attribute and contextualize information even when it retrieves a single chunk.

### Indexing Best Practices

1. **Use unique, deterministic item IDs** — Derive from source primary key + chunk index: `DOC-12345_chunk_3`
2. **Set the `title` to include section context** — Don't use the same title for every chunk
3. **Apply all relevant semantic labels to every chunk** — Each chunk is an independent item in the index
4. **Include the `url` that deep-links to the specific section** when possible
5. **Track chunk metadata** — Store `parentDocId`, `chunkIndex`, and `totalChunks` as queryable properties for administration

---

## 7. Handling Large Payloads

### Size Limits and Throttling

| Limit | Value |
|---|---|
| Item request body size | **4 MB** |
| Activities per call | **20** |
| Concurrent operations per connection | **25** |
| Throttle response | **HTTP 429** with `Retry-After` header |

### Throttle-Resilient Ingestion Pattern

```csharp
public async Task IngestWithRetry(ExternalItem item, int maxRetries = 5)
{
    for (int attempt = 0; attempt <= maxRetries; attempt++)
    {
        try
        {
            await graphClient.External.Connections[connectionId]
                .Items[item.Id]
                .PutAsync(item);
            return; // Success
        }
        catch (ServiceException ex) when (ex.ResponseStatusCode == 429)
        {
            var retryAfter = ex.ResponseHeaders?
                .RetryAfter?.Delta ?? TimeSpan.FromSeconds(Math.Pow(2, attempt));
            await Task.Delay(retryAfter);
        }
    }
    throw new Exception($"Failed to ingest item after {maxRetries} retries");
}
```

### Batch Ingestion Strategies

1. **Sequential with rate limiting** — Process items one at a time with a delay between requests. Simplest approach, lowest throughput.

2. **Concurrent with semaphore** — Use a semaphore to limit concurrent requests (recommended: **4–8 simultaneous calls**, but never exceeding the **25 concurrent operations per connection** platform limit). Balance throughput with throttle avoidance.

3. **Queue-based** — Push items to an Azure Queue / Service Bus, process with Azure Functions. Best for large-scale, fault-tolerant ingestion.

```csharp
// Concurrent ingestion with controlled parallelism
var semaphore = new SemaphoreSlim(4); // Max 4 concurrent requests
var tasks = items.Select(async item =>
{
    await semaphore.WaitAsync();
    try
    {
        await IngestWithRetry(item);
    }
    finally
    {
        semaphore.Release();
    }
});
await Task.WhenAll(tasks);
```

### Crawl Strategy Selection

| Strategy | Mechanism | Best For |
|---|---|---|
| **Full crawl** | Re-crawl entire data source | Initial load, periodic reconciliation, data sources without change tracking |
| **Incremental crawl** | Only detect additions/changes since last checkpoint | Ongoing sync, reducing API calls and processing time |
| **Event-based** | Push updates on source events (webhook, change feed) | Dynamic/sensitive data (ticket status changes, real-time updates) |
| **Scheduled** | Push at regular intervals | Content-rich, infrequently updated data (wiki pages, documentation) |

### Handling Items Near the 4 MB Limit

If a single item approaches the 4 MB limit:

1. **Measure the payload size** before sending — serialize to JSON and check byte length
2. **Trim non-essential content** — Remove boilerplate, repeated headers, navigation text
3. **Move verbose content to separate chunks** — Split the item using the chunking strategies from Section 6
4. **Compress whitespace** — Remove excessive newlines, redundant spaces
5. **Strip HTML of non-semantic elements** — Remove `<div>`, `<span>`, `class`, `id`, `style` attributes

### Ingestion Gotchas

- **`@odata.type` required for collection properties**: When sending `StringCollection` or other collection types in the item payload, include `@odata.type` annotations or the API may reject the request
- **Non-ASCII characters**: Ensure proper UTF-8 encoding; some languages inflate byte size significantly compared to character count
- **Item ID restrictions**: Item IDs must be URL-safe; avoid special characters like `#`, `?`, `&`, `/`
- **Property value limits**: While there is no explicit per-property size limit, excessively large property values can impact performance

```csharp
// Check payload size before ingestion
var json = JsonSerializer.Serialize(externalItem);
var byteSize = Encoding.UTF8.GetByteCount(json);

if (byteSize > 3_800_000) // Leave 200KB buffer
{
    // Chunk the content
    var chunks = ChunkContent(externalItem.Content.Value, maxChunkSize: 3_500_000);
    for (int i = 0; i < chunks.Count; i++)
    {
        var chunkedItem = CloneItemWithChunk(externalItem, chunks[i], i, chunks.Count);
        await IngestWithRetry(chunkedItem);
    }
}
else
{
    await IngestWithRetry(externalItem);
}
```

---

## 8. Handling Delimited Data (CSV / TSV / PSV)

### Pre-Built CSV Connector

Microsoft provides a **built-in CSV connector** in the M365 Admin Center for straightforward CSV ingestion. The CSV connector currently supports files stored in **SharePoint document libraries** or **Azure Data Lake Storage (ADLS)**:

1. Point to a `.csv` file in a supported source (SharePoint or ADLS)
2. Configure column-to-property mappings
3. Set a **unique identifier** column
4. Configure **multi-item delimiters** (e.g., semicolons within cells)
5. Set access control — file-level or item-level via `AllowedUsers`/`AllowedGroups` columns (must contain **Microsoft Entra object IDs**)
6. Assign semantic labels and configure schema

> ⚠️ If both file-level and item-level ACLs are configured, **file-level ACL takes precedence**.

> 💡 Use the pre-built CSV connector when your data is already in a clean CSV format in SharePoint or ADLS and you don't need custom transformation logic.

### Custom Connector: Parsing Delimited Data

When the pre-built connector doesn't meet your needs (complex transformations, multiple files, API-sourced CSV), parse the data in your custom connector:

#### Row-Per-Item Mapping

Each row in your CSV becomes one `externalItem`. Map columns to schema properties:

```csharp
// Parse CSV and create external items
using var reader = new StreamReader("data.csv");
using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
var records = csv.GetRecords<dynamic>();

foreach (var record in records)
{
    var item = new ExternalItem
    {
        Id = record.Id,
        Acl = new List<Acl> { /* ... */ },
        Properties = new Properties
        {
            AdditionalData = new Dictionary<string, object>
            {
                { "title", record.Title },
                { "category", record.Category },
                { "status", record.Status },
                { "tags", ParseMultiValue(record.Tags, ";") }
            }
        },
        Content = new ExternalItemContent
        {
            Value = BuildContentFromRow(record),
            Type = ExternalItemContentType.Text
        }
    };
    await IngestWithRetry(item);
}
```

#### Handling Multi-Value Delimiters

When a single CSV cell contains multiple values (e.g., `"Finance;HR;Engineering"`):

```csharp
// Parse semicolon-delimited values into StringCollection
private List<string> ParseMultiValue(string value, string delimiter = ";")
{
    if (string.IsNullOrWhiteSpace(value)) return new List<string>();
    return value.Split(delimiter)
                .Select(v => v.Trim())
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();
}
```

Map these to `StringCollection` properties in your schema with `isRefinable: true` and `isExactMatchRequired: true` for precise filtering.

#### Deciding What Goes into `content` vs. Properties

| Put in `content` | Put in properties (only) |
|---|---|
| Free-text descriptions, notes, comments | Identifiers (IDs, SKUs) |
| Combined narrative from multiple columns | Status values, categories |
| Any text users would naturally search for | Dates, numbers for filtering |
| Context that helps Copilot understand the item | URLs, email addresses |

**Pattern:** Build content by concatenating descriptive columns with labels:

```csharp
private string BuildContentFromRow(dynamic record)
{
    var sb = new StringBuilder();
    sb.AppendLine($"Product: {record.ProductName}");
    sb.AppendLine($"Category: {record.Category}");
    sb.AppendLine($"Description: {record.Description}");
    if (!string.IsNullOrEmpty(record.Notes))
        sb.AppendLine($"Notes: {record.Notes}");
    if (!string.IsNullOrEmpty(record.Specifications))
        sb.AppendLine($"Specifications: {record.Specifications}");
    return sb.ToString();
}
```

#### Handling Large CSV Files

For CSV files with millions of rows:

1. **Stream the file** — Don't load the entire file into memory
2. **Batch in groups** of 100–500 items with concurrent ingestion
3. **Implement checkpointing** — Track the last successfully ingested row number
4. **Handle duplicate rows** — Use deterministic item IDs derived from unique columns
5. **Consider partitioning** — Split very large CSVs by date, category, or region

---

## 9. Data Aggregation Across Multiple Items

### The Challenge

Large Language Models (including M365 Copilot) are **not reliable at performing aggregation operations across multiple retrieved items**. When a user asks:

- *"How many open tickets are assigned to the engineering team?"*
- *"What is the total revenue across all Q3 deals?"*
- *"Which product category has the most support incidents?"*

Copilot retrieves a subset of relevant items and attempts to reason over them. However, it may:
- Miss items that weren't retrieved (search returns a limited result set)
- Miscalculate counts, sums, or averages
- Provide confident-sounding but incorrect aggregate answers

### Strategy 1: Pre-Computed Summary Items (Recommended)

Ingest **dedicated summary items** alongside your detail items. These items contain pre-calculated aggregations:

```json
{
  "id": "summary-engineering-tickets-2026-03",
  "properties": {
    "title": "Engineering Team Ticket Summary — March 2026",
    "reportType": "monthly-summary",
    "team": "Engineering",
    "period": "2026-03",
    "openTickets": 47,
    "closedTickets": 123,
    "avgResolutionHours": 18.5,
    "topCategories": "Infrastructure (32%), Application Bugs (28%), Security (15%)"
  },
  "content": {
    "value": "Engineering Team Ticket Summary for March 2026. Total open tickets: 47. Total closed tickets: 123. Average resolution time: 18.5 hours. Top categories: Infrastructure (32%), Application Bugs (28%), Security (15%). P1 incidents: 3 (all resolved). Trend: 12% reduction in open tickets compared to February 2026.",
    "type": "text"
  }
}
```

**When to create summaries:**
- At regular intervals (daily, weekly, monthly)
- Per logical grouping (team, department, product, region)
- For commonly-asked aggregate questions
- During ingestion pipeline execution

### Strategy 2: Roll-Up Properties

Add pre-computed aggregate values as properties on individual items:

```json
{
  "id": "project-alpha",
  "properties": {
    "title": "Project Alpha",
    "totalTasks": 156,
    "completedTasks": 98,
    "completionPercentage": 62.8,
    "totalBudgetSpent": 450000,
    "teamSize": 12,
    "daysRemaining": 45
  }
}
```

Compute these values during your ingestion pipeline so Copilot can retrieve them directly without needing to aggregate.

### Strategy 3: Declarative Agent Instructions

When using a **Declarative Agent (DA)** with your connector, include explicit instructions about aggregation limitations and available summary data:

```
## Data Interpretation Instructions

This connector contains individual support tickets and pre-computed summary items.

When users ask aggregate questions (counts, totals, averages, trends):
1. Look for summary items first (property: reportType = "monthly-summary" or "weekly-summary")
2. Reference the pre-computed values in summary items rather than counting individual tickets
3. If no summary item exists for the requested aggregation, clearly state that the data shown represents a sample and may not be comprehensive
4. Never present a count derived from search results as an exact total

Available summary item types:
- Monthly team summaries (reportType: "monthly-summary")
- Weekly status reports (reportType: "weekly-summary")
- Quarterly trend analyses (reportType: "quarterly-trend")
```

### Strategy 4: Connector Actions + Power Automate

For **real-time aggregation needs**, use connector actions that call back to your source system's API:

1. **Define a connector action** in Copilot Studio that calls your API
2. The API performs the aggregation query directly against the source database
3. Return the calculated result to Copilot

This approach is ideal when:
- Aggregation data changes frequently
- Pre-computation is impractical (too many dimensions)
- Users need exact, real-time numbers

### Strategy 5: Federated Connectors (Preview)

For scenarios where data is too sensitive, too dynamic, or too aggregation-heavy to index:

- **Federated connectors** provide live, read-only access to external data without indexing
- Queries are executed in real-time against the source system
- The source system can perform aggregation natively (SQL queries, API calls)
- Ideal for financial dashboards, real-time metrics, and sensitive compliance data

> 📌 Federated connectors are in preview. Use them when pre-computation is impractical and real-time accuracy is critical.

### Strategy 6: Hybrid Approach

Combine strategies for the best coverage:

```
┌─────────────────────────────────────────────────────────┐
│                    Ingestion Pipeline                    │
├─────────────────────────────────────────────────────────┤
│  1. Ingest individual items (detail records)            │
│  2. Compute and ingest summary items (daily/weekly)     │
│  3. Add roll-up properties to parent items              │
│  4. Configure DA instructions for interpretation        │
│  5. Set up connector actions for real-time aggregation   │
│  6. Consider federated connectors for live aggregation   │
└─────────────────────────────────────────────────────────┘
```

### Anti-Patterns to Avoid

| ❌ Anti-Pattern | Why It Fails |
|---|---|
| Relying on Copilot to count items | Search returns a limited subset; counts will be inaccurate |
| Asking Copilot to sum values across results | Same issue — incomplete data leads to wrong totals |
| Expecting Copilot to rank or sort all items | Copilot sees a window of results, not the full dataset |
| Ingesting raw data without summaries | Users will ask aggregate questions; without summaries, answers will be unreliable |
| Using Copilot as a database query engine | Copilot is a semantic reasoning tool, not a SQL engine |

---

## 10. Surfacing Data in Microsoft 365 Copilot

### Step-by-Step Enablement

#### 1. Configure Connection Description

When creating the connection, provide a rich description that answers:
- **What** kind of content does this connection have?
- **Who** uses this content and how do they refer to it?
- **When** in their workflow do users access this content?
- **What** are notable characteristics of the content?

```http
POST /external/connections
{
  "id": "contoso-helpdesk",
  "name": "Contoso Helpdesk",
  "description": "Internal IT helpdesk tickets from the Contoso Helpdesk system. Contains incident reports, service requests, and change requests. Used by IT support staff and employees to track and resolve technical issues. Tickets include descriptions, root causes, resolutions, priority levels, and assignment information."
}
```

> 💡 Also consider setting the **`contentCategory`** property on your connection (e.g., `"contentCategory": "howTo"` or `"contentCategory": "reference"`). This helps the platform categorize your content for relevance purposes.

#### 2. Apply Semantic Labels

At minimum, assign: **`title`**, **`url`**, and **`iconUrl`**.

For maximum Copilot relevance (in order of impact):
1. `title`
2. `lastModifiedDateTime`
3. `lastModifiedBy`
4. `url`
5. `fileName`
6. `fileExtension`

#### 3. Mark Properties as Searchable

The `searchable` attribute is the **most critical for Copilot** — it determines which properties Copilot can match against. Mark all content-rich, text-based properties as searchable.

> ⚠️ **Current limitation**: Only the `title` semantic label can currently be used directly in Copilot prompts. However, all `searchable` properties and the `content` property are used for semantic matching. Apply all applicable labels now to future-proof your schema.

#### 4. Configure Rank Hints

For **searchable** properties that are **not mapped to semantic labels**, configure rank hints to prioritize certain properties in search results. Set importance from `default` to `veryHigh` in the M365 Admin Center under **Search & intelligence > Customization > Relevance tuning**.

#### 5. Add urlToItemResolver

Enable URL detection so Copilot recognizes when users share links to your external content:

```http
PATCH /external/connections/contoso-helpdesk
{
  "activitySettings": {
    "urlToItemResolvers": [
      {
        "@odata.type": "#microsoft.graph.externalConnectors.itemIdResolver",
        "urlMatchInfo": {
          "baseUrls": ["https://helpdesk.contoso.com"],
          "urlPattern": "/tickets/(?<itemId>[0-9]+)"
        }
      }
    ]
  }
}
```

#### 6. Send User Activities

Activities boost item relevance. Supported activity types are: `created`, `modified`, `commented`, and `viewed`.

```http
POST /external/connections/contoso-helpdesk/items/SR00145/addActivities
{
  "activities": [
    {
      "@odata.type": "#microsoft.graph.externalConnectors.externalActivity",
      "type": "commented",
      "startDateTime": "2026-03-23T10:00:00Z",
      "performedBy": {
        "@odata.type": "#microsoft.graph.externalConnectors.identity",
        "id": "18948b93-d3ed-4307-9981-10fc36a08a52",
        "type": "user"
      }
    }
  ]
}
```

> ⚠️ Activities older than **7 days** don't surface in the Microsoft 365 app.

#### 7. Enable Inline Results

In the M365 Admin Center:
1. Go to **Search & intelligence** > **Customizations** > **Verticals**
2. Select the **All** vertical
3. Select **Manage connector result**
4. Ensure **Show results inline** is checked
5. Select the connections to enable for Copilot and Search

#### 8. Configure Result Types (Optional)

Create custom result layouts using **Adaptive Cards** for richer search result presentation. Note: For the **All vertical**, inline results will render with default or custom result types. For **custom verticals**, a result type/layout is generally required for proper rendering.

```json
{
  "type": "AdaptiveCard",
  "body": [
    {
      "type": "TextBlock",
      "text": "${title}",
      "weight": "Bolder",
      "size": "Medium"
    },
    {
      "type": "TextBlock",
      "text": "Status: ${status} | Priority: ${priority}",
      "spacing": "Small"
    },
    {
      "type": "TextBlock",
      "text": "${description}",
      "wrap": true,
      "maxLines": 3
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Open Ticket",
      "url": "${url}"
    }
  ]
}
```

### Declarative Agents with Connectors

When building a **Declarative Agent** that uses your connector as a knowledge source:

1. **Include property descriptions** in the agent's instruction set — this helps the agent understand the semantic meaning of each field
2. **Specify available query patterns** — tell the agent which properties support filtering
3. **Document aggregation boundaries** — explain what summary data is available (see Section 9)
4. **Provide example queries** — include sample natural-language questions and expected data sources

---

## 11. Access Control & Security

### ACL Model

Every external item must have an **Access Control List (ACL)** that specifies who can see the item. ACL values must use **Microsoft Entra object IDs** (GUIDs), not email addresses or UPNs:

```json
{
  "acl": [
    {
      "type": "user",
      "value": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      "accessType": "grant"
    },
    {
      "type": "group",
      "value": "b2c3d4e5-f6a7-8901-bcde-f12345678901",
      "accessType": "grant"
    },
    {
      "type": "everyone",
      "value": "everyone",
      "accessType": "grant"
    }
  ]
}
```

### Key Rules

- **`deny` overrides `grant`** — If a user is granted access via `everyone` but explicitly denied, the effective permission is `deny`
- **ACL types**: `user` (Entra ID user), `group` (Entra ID group), `everyone` (all tenant users), `everyoneExceptGuests` (all tenant users except guests), `externalGroup` (custom group)
- **Security trimming happens at query time** — Users only see items they have access to
- **Always use Entra object IDs** — If your source system uses UPNs or emails, translate them to Entra object IDs during ingestion

### External Groups

For data sources with non-Entra ID permissions (e.g., custom role systems):

1. **Create external groups** using the group sync APIs
2. Map your source system's roles/groups to external groups
3. External groups can contain: other external groups, Entra ID users, Entra ID groups

> ⚠️ **Avoid expanding group membership directly into individual item ACLs** — This leads to excessive item updates when group membership changes. Use external groups instead.

### Best Practices

- Mirror source system permissions as closely as possible
- Use `everyone` only for truly public content; prefer `everyoneExceptGuests` when guest access should be excluded
- Implement group-based ACLs over user-based ACLs for maintainability
- Translate non-Entra ID users to Entra ID object IDs in your ACL mapping
- Test ACL behavior with users of different permission levels

---

## 12. Monitoring, Testing & Troubleshooting

### SDK Test Utility

The Copilot Connectors SDK includes a **test utility** with pre-built scenarios:
- Connection creation/validation
- Schema registration verification
- Item ingestion with various content types
- ACL validation
- Crawl simulation

### Admin Center Monitoring

Monitor connector health in the **M365 Admin Center** under **Search & intelligence > Connectors**:
- Connection status (active, paused, failed)
- Item count and quota usage
- Crawl history and errors
- Throughput metrics

### Common Issues and Resolutions

| Issue | Cause | Resolution |
|---|---|---|
| Items not appearing in search | Schema not registered or items not indexed | Verify schema registration; check item status in Admin Center |
| HTTP 429 errors | Throttle limits exceeded | Implement exponential backoff with `Retry-After` header |
| Content not surfacing in Copilot | Missing semantic labels or `searchable` attribute | Add `title`, `url`, `iconUrl` labels; mark text properties as searchable |
| ACL errors | Invalid user/group IDs | Verify Entra ID resolution; check external group configuration |
| Schema update fails | Attempting incompatible changes | Some changes require creating a new connection; see schema update rules |
| Items showing to wrong users | Incorrect ACL configuration | Audit ACLs; test with specific user accounts |

### Schema Update Rules

| Operation | Supported? | Notes |
|---|---|---|
| Add a new property | ✅ Yes | Reingestion recommended |
| Add/remove search capability | ✅ Yes | Reingestion **required** |
| Add refinable attribute | ❌ Not via update | Requires new connection or initial schema |
| Add/remove alias | ✅ Yes | Auto-created aliases for refinable properties can't be removed |
| Add/remove semantic label | ✅ Yes | May affect search and Copilot experiences |

> 📌 **After any schema update, reindex items to align with the latest schema.** Without reingestion, item behavior may be inconsistent.

---

## 13. Quick-Reference Checklists

### Pre-Launch Checklist

- [ ] Schema registered with all required properties
- [ ] Semantic labels assigned: `title`, `url`, `iconUrl` (minimum)
- [ ] Text properties marked as `searchable`
- [ ] Filter properties marked as `queryable` and/or `refinable`
- [ ] Display properties marked as `retrievable`
- [ ] Content property populated with rich, descriptive text
- [ ] ACLs configured and tested with multiple user roles
- [ ] Connection description is detailed and descriptive
- [ ] urlToItemResolver configured for URL-based boosting
- [ ] User activities being sent for relevant interactions
- [ ] Inline results enabled in the "All" vertical
- [ ] Throttle handling and retry logic implemented
- [ ] Incremental crawl strategy defined for ongoing sync
- [ ] Items within 4 MB size limit (chunked if necessary)

### Copilot Optimization Checklist

- [ ] Content is information-dense and well-structured
- [ ] Content leads with the most important information
- [ ] Multiple text fields concatenated into content with labels
- [ ] Summary items ingested for aggregate data queries
- [ ] Declarative Agent instructions include property descriptions
- [ ] Declarative Agent instructions address aggregation limitations
- [ ] Properties have clear, descriptive names (not abbreviations)
- [ ] Aliases defined for common synonyms

### Security Checklist

- [ ] ACLs mirror source system permissions
- [ ] External groups used for non-Entra ID permissions
- [ ] Group memberships not expanded into individual ACLs
- [ ] `deny` entries used sparingly and intentionally
- [ ] Compliance content type set to `text` (if applicable)
- [ ] Sensitive data excluded or properly access-controlled

---

## Appendix A: Key Links & References

| Resource | URL |
|---|---|
| Copilot Connectors API Overview | https://learn.microsoft.com/graph/connecting-external-content-connectors-api-overview |
| Schema Registration & Best Practices | https://learn.microsoft.com/graph/connecting-external-content-manage-schema |
| Creating & Managing Items | https://learn.microsoft.com/graph/connecting-external-content-manage-items |
| API Limits | https://learn.microsoft.com/graph/connecting-external-content-api-limits |
| Copilot Connector Experiences | https://learn.microsoft.com/graph/connecting-external-content-experiences |
| Connectors SDK Overview | https://learn.microsoft.com/graph/custom-connector-sdk-overview |
| SDK Sample Connectors (GitHub) | https://github.com/microsoftgraph/msgraph-connectors-sdk |
| Build Custom Connector in C# | https://learn.microsoft.com/graph/custom-connector-sdk-sample-create |
| CSV Connector | https://learn.microsoft.com/microsoft-365/copilot/connectors/csv-connector |
| Result Layout Design | https://learn.microsoft.com/microsoftsearch/customize-results-layout |
| Adaptive Cards Designer | https://adaptivecards.io/designer/ |
| M365 Copilot Extensibility Overview | https://learn.microsoft.com/microsoft-365-copilot/extensibility/ |
