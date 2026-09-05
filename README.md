# Production Reports Tool

The **Production Reports Tool** is a Microsoft Access desktop reporting and
operational-analysis application for aviation maintenance and supply data. Its
source indicates that it is designed for personnel working with NALCOMIS IMA
data, including production, maintenance, inventory, requisition, and support
records.

The database brings recurring operational data into one interface where users
can run NALCOMIS AdHoc queries, review Maintenance Action Form (MAF) records,
search by aircraft and supply identifiers, generate production reports, and
export selected results to Excel.

This repository contains a text-based source export of
**`Production Reports Tool.accdb`**. The export was produced by the MSAccess VCS
add-in so that Access objects can be reviewed and versioned with Git. It is not
a standalone web application and cannot be run directly on Replit.

## What the database does

### Automates NALCOMIS AdHoc reports

The database includes VBA automation that works with the NALCOMIS IMA desktop
application. Based on configured AdHoc definitions, it:

1. Locates an open NALCOMIS session.
2. Opens the NALCOMIS Reports and Ad Hoc Query interfaces.
3. Loads configured `.ah` query files.
4. Runs each query and waits for its results.
5. Saves the output as CSV.
6. Records whether each run succeeded, failed, or returned fewer rows than
   expected.

Separate runners appear to support active, historical, AK0/AK7, selected, and
all configured AdHoc query sets.

### Produces maintenance and supply reports

The application groups maintenance information into several recurring report
families. The exported forms, queries, and reports include:

- maintenance levels 400 through 900;
- PMS, DIFM, BSY, and BSR views;
- 500/600 Priority, In Work (IW), and discrepancy reports;
- Dog House reports by maintenance level;
- Exrep and Refer With Stock views;
- MAF Rejects, Buffer Status, PMI, MOC, SEPMS, and MMCO Daily views;
- Cost Savings, Raw Data, and Requisition views;
- AK0/AK7 and all-BSR-with-DDSN outputs.

Some of these are formal Access report objects, while others are query-backed
forms or export workflows.

### Supports MAF research

Users can review MAF details together with related notes, parts, support, and
requisition information. Search tools allow records to be researched by fields
such as:

- FGC and TEC;
- NIIN and DDSN;
- MCN and JCN;
- BUNO and serial number;
- part and equipment serial numbers;
- management and project codes;
- current-status and initiated dates.

The database's core maintenance dataset includes fields for aircraft,
equipment, maintenance status, work centers, parts, priorities, dates, and
other production-tracking information.

### Provides a central Access navigation interface

The database starts on an Access form named `Navigation Form`. This form acts as
the main menu for report groups, research tools, MAF detail views, raw data,
configuration, and AdHoc execution. Selected query results can be exported to
Excel, and formal Access reports can be opened in preview mode.

## Setup and external dependencies

On first run, the database creates a local `Production Reports` folder
structure, writes configured AdHoc templates and empty CSV files, and can prompt
the user to run the initial AdHoc set.

The complete runtime environment is not contained in this repository. The
application depends on:

- Microsoft Access on Windows;
- the NALCOMIS IMA desktop application for automated AdHoc execution;
- configured `.ah` templates and CSV destinations;
- locally imported or linked maintenance, inventory, requisition, and reference
  data;
- Windows UI automation, including window-title detection and simulated
  keyboard input.

The exact upstream data-loading process, site-specific paths, user permissions,
and production operating procedures should be confirmed in the original Access
environment.

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
