# Stylecolab (SC) Expense Classification — David Jones Amex

**Purpose:** Identify which David Jones American Express transactions are **Stylecolab business
expenses** (paid on the personal Amex card), and split them into **financial years for tax**.

This file is the durable record of Vicki's classification so it can be applied to future
statement batches. The workbook
`DavidJones_Amex_2022-23_SC-MARKED_by-Vicki_GROUND-TRUTH.xlsx` in this folder is the
**authoritative source** — it contains Vicki's own yellow highlights and `SC` column for the
2022–23 set. Machine-readable rules are in `stylecolab_sc_rules.json`.

## Financial year (for the tax split)
- **Australian FY: 1 July → 30 June.** Split by **transaction date**, not statement date.
- Batches to process: **FY2020-21**, **FY2021-22**, **FY2022-23** (2022-23 statements already held).
- Note: a statement dated the 22nd of a month contains transactions spanning ~23rd of the prior
  month to the 22nd, so month-end transactions land in the correct FY by their own date.

## Scope
- SC applies to the **primary cardholder, Vicki Zertopoulos**, only.
- **George Doufas'** supplementary card (…61025) is treated as **personal** unless told otherwise.

## 2022–23 summary (ground truth)
- **250** transaction rows flagged SC · total SC debits **$33,914.54**.

## Tier 1 — Auto-flag SC (flagged 100% of the time / clear business SaaS & services)
Software, subscriptions, suppliers and services that are always business:

Xero · Shopify (SHOPIFYCOMM) · Clickfunnels · Squarespace · ActiveCampaign · Netregistry ·
Kindle Unltd · Scribd · Spotify · Microsoft (MSFT) · LastPass · Canva · Upwork · Lawpath ·
Savvy Shopkeeper · The Boutique Hub · Salvos Stores · Kmart (PayPal KMARTAUSTRA) · Cabcharge ·
SP Email By Design · Bespoke Collective · CM Style Squad (CMSTYLESQUA) · Magic Tailor ·
Looksmart Alterations · Academy of Professional · Supply and Demand · Lombard The Paper People

## Tier 2 — Review (mixed: only *some* purchases are business — Vicki decides per transaction)
These merchants appear as **both** business and personal, so each transaction needs Vicki's call.
On a new batch I will pre-mark these **REVIEW** rather than auto-flagging them:

David Jones (only 20 of 51 were SC) · Country Road · Witchery · Zara · Oroton · Kathmandu ·
Aesop · Bed Bath N Table · Readings · The Sage Method · Meetup · Otter.ai · PayPal Paddle ·
PayPal Bellharr · PayPal Savvyshopke · Coles · Uber · Grill'd · Leaf Hawthorn · TK Maxx ·
Amazon (AMZN Digital / Marketplace) · SP Milligram · Centr · PayPal Goat · Donna Cameron ·
Magshop · Mamamia · Secondo · plus the Feb-2023 Sydney work trip (Airport, ICC Sydney, JB Hi-Fi,
Barangaroo House) · Cultivate Nursery · Big W · eBay

## How this gets applied to a new batch
1. Extract & reconcile the statements (as done for 2022-23).
2. Add the `SC` column on Vicki's transactions:
   - **Auto-flag** rows matching Tier 1 keywords.
   - **Mark "REVIEW"** rows matching Tier 2 keywords.
   - Leave the rest blank.
3. Split output into AU financial-year tabs/files by transaction date.
4. Vicki reviews the REVIEW rows and any edits update this rule set.

_Vicki's manual edits always override these rules._
