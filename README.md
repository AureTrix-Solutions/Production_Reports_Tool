# Production Reports Tool

This repository contains a text-based source export of the Microsoft Access
database **`Production Reports Tool.accdb`**.

The export was produced by the MSAccess VCS add-in so that Access objects can be
reviewed and versioned with Git. It is not a standalone web application and
cannot be run directly on Replit.

## Repository contents

| Path | Contents |
| --- | --- |
| `modules/` | Standard VBA modules (`.bas`) |
| `forms/` | Access form definitions, metadata, and code-behind |
| `reports/` | Access report definitions, metadata, and code-behind |
| `queries/` | Saved query SQL and query metadata |
| `tbldefs/` | Table-definition SQL and XML metadata |
| `relations/` | Database relationship definitions |
| `imexspecs/` | Import/export specifications |
| `themes/` | Access theme assets |
| `project.json` | General Access project metadata |
| `vbe-project.json` | VBA project metadata |
| `dbs-properties.json` | Database properties |
| `documents.json` | Access document metadata |
| `nav-pane-groups.json` | Navigation Pane configuration |

`AGENTS.md` contains detailed guidance for safely editing the exported source,
including the required file formats and object-specific conventions.

## Working with the project

### Review source

The text export can be searched and reviewed directly:

- VBA logic is primarily in `modules/` and in form/report `.cls` files.
- Saved SQL is available in `queries/`.
- Table definitions are available in `tbldefs/`.
- Form and report layouts are stored in Access `SaveAsText`-style files.

### Build or run the database

To create a working Access database, use the MSAccess VCS add-in from Microsoft
Access on Windows to import these source files into an `.accdb` file. Microsoft
Access and the add-in are required; Replit does not provide the Access runtime
needed to execute or visually test the database.

### Make changes

When editing exported Access objects:

1. Preserve UTF-8 BOM encoding and CRLF line endings.
2. Do not alter required VBA headers such as `Attribute VB_Name`.
3. Do not edit generated binary/index files such as `vcs-index.idx`.
4. Import the changed source into Microsoft Access and test it there.
5. Re-export from Access when appropriate so the repository remains synchronized.

See `AGENTS.md` before making changes for the complete safety rules.

## Replit status

No run workflow is configured because this repository is an Access source
export rather than a Replit-runnable application. The imported project has
otherwise been left unchanged.

## Documentation opportunities

Useful project-specific documentation can be added as the database is explored,
for example:

- the purpose and intended users of the Production Reports Tool;
- required linked databases, network paths, or external data sources;
- setup and first-run instructions inside Microsoft Access;
- descriptions of major forms, reports, and workflows;
- release, import, export, and validation procedures.
