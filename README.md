# Conciliación Bancaria — Chilean SME Edition

🔗 **[Try it live →](https://claude.ai/public/artifacts/7b05494d-ede2-4b63-8aba-9810f93434a0)** — No install, no login, runs in browser.

Multi-pass bank reconciliation engine that matches bank statement movements 
(cartolas) against SII accounting sources: Libro de Compras (LC), Libro de 
Ventas (LV), Remuneraciones, and Boletas de Honorarios (BHE). Built for the 
specific messiness of Chilean SME accounting.

## The problem

Chilean accountants reconcile bank statements manually every month. A typical 
SME has 200–500 monthly movements across multiple sources. Matching them by 
hand takes 4–8 hours per client. Existing tools either assume clean data or 
require perfect RUT/reference fields — which real-world cartolas don't have.

## How it works

The engine runs 4 sequential matching passes. Bank movements are always the 
base row; accounting sources are referenced from them, never the reverse.

| Pass | Criterion | Use case |
|------|-----------|----------|
| 1. RUT match | Exact tax ID match | Highest-confidence matches |
| 2. Keyword match | Fuzzy match on descriptions | Entries without RUT |
| 3. Amount match | Amount + date range | Handles split payments |
| 4. Date match | Wider date window | Last resort fallback |

Each match gets a dynamic score. Unmatched entries are flagged for manual 
review instead of force-matched.

## Stack

- React (single-file component)
- All matching logic runs client-side

## Status

v4 — in active use at SUR Consulting for client reconciliations.

## Part of ReportIA

This tool is one piece of a broader automation pipeline for Chilean SME 
financial closing. See also:
- [clasificador-facturas-ifrs](https://github.com/lucasplvd/clasificador-facturas-ifrs)
- [cobranza-whatsapp-agent](https://github.com/lucasplvd/cobranza-whatsapp-agent)
