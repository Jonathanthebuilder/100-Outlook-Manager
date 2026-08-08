# Outlook Manager Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local-first web app that imports Outlook mailbox TXT records, separates available and used accounts, supports grouped selection and remarking, and exports all still-valid accounts as plain text.

**Architecture:** Use a React + Vite single-page app with browser-local persistence. Keep mailbox parsing, grouping, mutation, and export logic in a pure TypeScript module covered by Vitest tests, while the React app handles file import, grouping filters, selection, remark forms, and downloads.

**Tech Stack:** React, TypeScript, Vite, Vitest, CSS modules via plain CSS.

---

### Task 1: Project Skeleton

**Files:**
- Create: `package.json`
- Create: `index.html`
- Create: `tsconfig.json`
- Create: `tsconfig.node.json`
- Create: `vite.config.ts`
- Create: `src/main.tsx`
- Create: `src/vite-env.d.ts`

- [ ] **Step 1: Create package and Vite config**

Add React, Vite, TypeScript, and Vitest scripts. Use `npm install` after files are created.

- [ ] **Step 2: Create TypeScript and HTML entry files**

Create the Vite HTML entry and TS config files.

- [ ] **Step 3: Create minimal React entry**

Create `src/main.tsx` so the project can compile before app implementation.

- [ ] **Step 4: Install dependencies**

Run: `npm install`
Expected: dependencies install and `package-lock.json` is generated.

### Task 2: Mailbox Ledger Core

**Files:**
- Create: `src/lib/mailboxLedger.ts`
- Create: `src/lib/mailboxLedger.test.ts`

- [ ] **Step 1: Write failing parsing test**

Test that plain lines become available records, lines ending with `【备注】` become used records, malformed lines are reported, and comments are preserved without being included in export text.

- [ ] **Step 2: Verify the test fails**

Run: `npm test -- src/lib/mailboxLedger.test.ts`
Expected: FAIL because the module is not implemented.

- [ ] **Step 3: Implement parsing and serialization**

Export functions:
- `parseMailboxText(text: string): ParsedMailboxFile`
- `exportAvailableMailboxes(records: MailboxRecord[]): string`

- [ ] **Step 4: Verify parsing test passes**

Run: `npm test -- src/lib/mailboxLedger.test.ts`
Expected: PASS.

- [ ] **Step 5: Add grouping and marking tests**

Test grouping by first letter and domain, deterministic random selection from a group, and marking a record used with remark.

- [ ] **Step 6: Verify new tests fail**

Run: `npm test -- src/lib/mailboxLedger.test.ts`
Expected: FAIL because grouping/marking functions are missing.

- [ ] **Step 7: Implement grouping and marking helpers**

Export functions:
- `getMailboxGroups(records: MailboxRecord[]): MailboxGroup[]`
- `pickRandomAvailable(records: MailboxRecord[], filters: PickFilters, random?: () => number): MailboxRecord | undefined`
- `markMailboxUsed(records: MailboxRecord[], id: string, remark: string, usedAt?: string): MailboxRecord[]`
- `getInventoryStats(records: MailboxRecord[]): InventoryStats`

- [ ] **Step 8: Verify all core tests pass**

Run: `npm test -- src/lib/mailboxLedger.test.ts`
Expected: PASS.

### Task 3: App UI and Persistence

**Files:**
- Create: `src/App.tsx`
- Create: `src/styles.css`
- Modify: `src/main.tsx`

- [ ] **Step 1: Build app shell**

Implement a dashboard-style app with import/export controls, stats, grouped available pool, selected mailbox details, remark form, used pool, and raw import preview errors.

- [ ] **Step 2: Add browser-local persistence**

Save parsed/edited records to `localStorage` under `outlook-manager-records-v1` and restore them on load.

- [ ] **Step 3: Add TXT import**

Use a file input and `FileReader` to parse TXT content, then replace current records after confirmation if records already exist.

- [ ] **Step 4: Add export download**

Create a Blob from `exportAvailableMailboxes(records)` and download `available-outlook-mailboxes.txt`.

- [ ] **Step 5: Wire random pick and mark-used flow**

Allow selecting a group, randomly picking an available mailbox, copying its fields, entering a remark, and marking it used.

### Task 4: Verification

**Files:**
- No new files expected.

- [ ] **Step 1: Run tests**

Run: `npm test -- --run`
Expected: PASS.

- [ ] **Step 2: Run production build**

Run: `npm run build`
Expected: PASS.

- [ ] **Step 3: Start dev server**

Run: `npm run dev -- --host 127.0.0.1`
Expected: local URL is available.

- [ ] **Step 4: Browser smoke test**

Open the local app, verify the initial screen renders, sample data can be imported through the paste/demo path or file import where possible, a mailbox can be marked used, and export produces available-only text.
