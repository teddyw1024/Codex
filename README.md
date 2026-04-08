# Lenovo + Dell Warranty Lookup Tool

This repository contains a PowerShell warranty checker for Lenovo serial numbers and Dell service tags.

## Files

- `warranty_lookup.ps1`: CLI tool (JSON and text output)
- `warranty_lookup_ui.ps1`: desktop Windows Forms UI

## Features

- Auto-detects Lenovo vs Dell from input
- Returns warranty start date, expiration date, and status
- Returns model when available
- Returns spec availability and spec URL when available
- Provides a result URL for manual confirmation

## Run (UI)

```powershell
powershell -ExecutionPolicy Bypass -File .\warranty_lookup_ui.ps1
```

## Run (CLI)

```powershell
powershell -ExecutionPolicy Bypass -File .\warranty_lookup.ps1 -Serial YOUR_SERIAL_OR_TAG
```

## JSON output

```powershell
powershell -ExecutionPolicy Bypass -File .\warranty_lookup.ps1 -Serial YOUR_SERIAL_OR_TAG -AsJson
```
