// @ts-nocheck
import { downloadReportAsDocx, downloadReportAsXlsx } from './documentExporters';

export function initReportBuilderEngine(onResetCallback?: () => void) {
    let partners = [
        { id: 1, name: 'Tariqul Islam', role: 'Managing Partner & Strategy', ratio: 60, open: 3000000, add: 500000, draw: 450000 },
        { id: 2, name: 'Nusrat Jahan', role: 'Partner - Finance & Operations', ratio: 40, open: 2000000, add: 250000, draw: 300000 }
    ];

    let pnlExpenses = [
        { id: 1, label: 'Selling and distribution expenses', note: 'Note 4', curr: 1450000, prev: 1220000 },
        { id: 2, label: 'Administrative expenses', note: 'Note 17', curr: 2380000, prev: 2050000 },
        { id: 3, label: 'Other operating expenses', note: 'Note 4', curr: 210000, prev: 160000 }
    ];

    let cfInvesting = [
        { id: 1, label: 'Purchase of property, plant and equipment', amount: -1100000 },
        { id: 2, label: 'Proceeds from sale of property, plant and equipment', amount: 150000 }
    ];

    let cfFinancing = [
        { id: 1, label: 'Proceeds from borrowings', amount: 400000 },
        { id: 2, label: 'Repayment of borrowings', amount: -650000 }
    ];

    let bsItems = {
        nca: [
            { id: 1, label: 'Property, plant and equipment', note: 'Note 8', curr: 4650000, prev: 3900000 },
            { id: 2, label: 'Intangible assets', note: 'Note 4', curr: 350000, prev: 400000 }
        ],
        ca: [
            { id: 3, label: 'Inventories', note: 'Note 9', curr: 2850000, prev: 2300000 },
            { id: 4, label: 'Trade and other receivables', note: 'Note 10', curr: 2150000, prev: 1850000 },
            { id: 5, label: 'Cash and cash equivalents', note: 'Note 11', curr: 2000000, prev: 1450000, isCash: true }
        ],
        ncl: [
            { id: 6, label: 'Long-term borrowings', note: 'Note 15', curr: 2200000, prev: 2450000 }
        ],
        cl: [
            { id: 7, label: 'Trade and other payables', note: 'Note 16', curr: 1750000, prev: 1650000 },
            { id: 8, label: 'Short-term borrowings', note: 'Note 15', curr: 1050000, prev: 800000 }
        ]
    };

    let disclosureNotes = [
        { id: 26, title: 'Note 26 — Additional Regulatory Compliance', text: 'All statutory filings and trade licences were duly renewed during the reporting period.' }
    ];

    let structuredNotes8to14 = {
        n8: {
            visible: true,
            title: 'Note 8 — Property, Plant, Equipment and Intangible Assets (IAS 16, IAS 38)',
            rows: [
                { id: 1, cat: 'Freehold land and building', cost: 2500000, dep: -300000 },
                { id: 2, cat: 'Cold-storage plant and machinery', cost: 2300000, dep: -800000 },
                { id: 3, cat: 'Furniture and fixtures', cost: 700000, dep: -280000 },
                { id: 4, cat: 'Office and distribution equipment', cost: 900000, dep: -370000 },
                { id: 5, cat: 'Total property, plant and equipment', cost: 6400000, dep: -1750000, isTotal: true },
                { id: 6, cat: 'Intangible assets (software and licences)', cost: 700000, dep: -350000 }
            ],
            footer: 'Additions during the year: property, plant and equipment BDT 1,100,000. Depreciation and amortisation for the year amounted to BDT 520,000 and BDT 50,000 respectively.'
        },
        n9: {
            visible: true,
            title: 'Note 9 — Inventories (IAS 2)',
            rows: [
                { id: 1, label: 'Finished goods & commodities', curr: 1850000, prev: 1520000 },
                { id: 2, label: 'Packaging and distribution supplies', curr: 620000, prev: 540000 },
                { id: 3, label: 'Goods in transit', curr: 380000, prev: 240000 },
                { id: 4, label: 'Total inventories (at lower of cost and NRV)', curr: 2850000, prev: 2300000, isTotal: true }
            ]
        },
        n10: {
            visible: true,
            title: 'Note 10 — Trade and Other Receivables (IFRS 9, IFRS 7)',
            rows: [
                { id: 1, label: 'Trade receivables', curr: 2350000, prev: 2020000 },
                { id: 2, label: 'Less: Loss allowance (expected credit losses)', curr: -200000, prev: -170000 },
                { id: 3, label: 'Net trade receivables', curr: 2150000, prev: 1850000, isTotal: true }
            ]
        },
        n11: {
            visible: true,
            title: 'Note 11 — Cash and Cash Equivalents (IAS 7)',
            rows: [
                { id: 1, label: 'Cash in hand', curr: 120000, prev: 90000 },
                { id: 2, label: 'Cash at bank — current accounts', curr: 1880000, prev: 1360000 },
                { id: 3, label: 'Total cash and cash equivalents', curr: 2000000, prev: 1450000, isTotal: true }
            ]
        },
        n12: {
            visible: true,
            title: "Note 12 — Partners' Capital Accounts (IAS 1, IAS 32)",
            footer: 'Illustrative figures only. This example assumes the balances qualify for equity classification under the terms of the partnership agreement and IAS 32.'
        },
        n13: {
            visible: true,
            title: "Note 13 — Partners' Current Accounts and Drawings (IAS 1)",
            text: "Drawings during the year are recorded as a direct reduction of the respective partner's capital account and are not recognised as an expense of the Firm. The amounts are disclosed in the Statement of Changes in Partners' Equity and in Note 12."
        },
        n14: {
            visible: true,
            title: 'Note 14 — Profit or Loss Allocation Among Partners (IAS 1)',
            text: ''
        }
    };

    let structuredNotes15to25 = {
        n15: {
            visible: true,
            title: 'Note 15 — Borrowings and Loans (IFRS 9, IFRS 7)',
            rows: [
                { id: 1, label: 'Non-current', curr: null, prev: null, isHeader: true },
                { id: 2, label: 'Bank term loan (secured)', curr: 2200000, prev: 2450000 },
                { id: 3, label: 'Current', curr: null, prev: null, isHeader: true },
                { id: 4, label: 'Short-term bank borrowings (working capital facilities)', curr: 1050000, prev: 800000 },
                { id: 5, label: 'Total borrowings', curr: 3250000, prev: 3250000, isTotal: true }
            ],
            footer: "The term loan is repayable in quarterly instalments and carries interest at a floating rate linked to the bank's lending rate. Security comprises a registered charge over the Firm's plant and machinery."
        },
        n16: {
            visible: true,
            title: 'Note 16 — Trade and Other Payables (IAS 1)',
            rows: [
                { id: 1, label: 'Trade payables', curr: 1330000, prev: 1300000 },
                { id: 2, label: 'Accrued expenses', curr: 260000, prev: 220000 },
                { id: 3, label: 'Other payables', curr: 160000, prev: 130000 },
                { id: 4, label: 'Total trade and other payables', curr: 1750000, prev: 1650000, isTotal: true }
            ]
        },
        n17: {
            visible: true,
            title: 'Note 17 — Employee Benefits (IAS 19)',
            text: 'Employee benefit expense recognised during the year amounted to BDT 1,720,000 (2025: BDT 1,480,000), comprising salaries and wages, allowances and statutory contributions. The Firm operates no defined benefit plan.'
        },
        n18: {
            visible: true,
            title: 'Note 18 — Finance Costs (IAS 23)',
            rows: [
                { id: 1, label: 'Interest on bank term loan', curr: 200000, prev: 220000 },
                { id: 2, label: 'Interest on short-term borrowings and bank charges', curr: 120000, prev: 130000 },
                { id: 3, label: 'Total finance costs', curr: 320000, prev: 350000, isTotal: true }
            ]
        },
        n19: {
            visible: true,
            title: 'Note 19 — FY 2025-2026 (BDT) - CURRENT AND DEFERRED TAXATION (IAS 12 & S.R.O. 324/2026)',
            extraRows: []
        },
        n20: {
            visible: true,
            title: 'Note 20 — Related Party Transactions (IAS 24)',
            text: "Transactions with related parties comprise partners' drawings, capital contributions and profit allocations as disclosed in Notes 12 to 14. No other material transactions with related parties were entered into during the year."
        },
        n21: {
            visible: true,
            title: 'Note 21 — Financial Risk Management (IFRS 7)',
            text: 'The Firm is exposed to credit risk, liquidity risk and market risk (including interest rate and foreign exchange risk). These risks are monitored by management on an ongoing basis. Credit risk is managed through credit limits and periodic ageing reviews; liquidity risk through committed banking facilities and cash-flow forecasting; and interest rate and foreign exchange risk through the terms of borrowing arrangements and forward cover.'
        },
        n22: {
            visible: true,
            title: 'Note 22 — Commitments and Contingencies (IAS 37)',
            text: 'At the reporting date the Firm had capital commitments of BDT 420,000 in respect of warehouse equipment (2025: BDT 250,000). There were no material contingent liabilities requiring recognition or disclosure.'
        },
        n23: {
            visible: true,
            title: 'Note 23 — Capital Management (IAS 1)',
            text: "The Firm's capital comprises partners' capital accounts. The objective of capital management is to safeguard the Firm's ability to continue as a going concern and to maintain an appropriate capital structure. The debt-to-equity ratio at the reporting date was 0.71 times (2025: 0.98 times)."
        },
        n24: {
            visible: true,
            title: 'Note 24 — Events After the Reporting Period (IAS 10)',
            text: 'No adjusting or non-adjusting events requiring disclosure occurred between the reporting date and the date on which the financial statements were authorised for issue.'
        },
        n25: {
            visible: true,
            title: 'Note 25 — Comparative Information (IAS 1)',
            text: "Comparative information has been presented for the previous reporting period in respect of all amounts reported in the primary statements and notes, in accordance with IAS 1. Where necessary, comparative figures have been reclassified to conform to the current period's presentation."
        }
    };

    let nextUid = 100;

    function fmt(num, showParens = false) {
        if (isNaN(num) || num === null) return '0';
        const rounded = Math.round(num);
        const absStr = Math.abs(rounded).toLocaleString('en-US');
        if (rounded < 0 || showParens) {
            return rounded === 0 ? '0' : `(${absStr})`;
        }
        return absStr;
    }

    function fmtDec(num, showParens = false) {
        if (isNaN(num) || num === null) return '0.00';
        const absStr = Math.abs(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        if (num < 0 || showParens) {
            return Math.abs(num) < 0.005 ? '0.00' : `(${absStr})`;
        }
        return absStr;
    }

    function escapeHtml(str) {
        return String(str ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function val(id) {
        const el = document.getElementById(id);
        return el ? (parseFloat(el.value) || 0) : 0;
    }

    function strVal(id) {
        const el = document.getElementById(id);
        return el ? el.value : '';
    }

    function setText(id, text) {
        const el = document.getElementById(id);
        if (el) el.textContent = text;
    }

    function setClassText(className, text) {
        document.querySelectorAll('.' + className).forEach(el => {
            el.textContent = text;
        });
    }

    // Partners CRUD
    window.addPartner = function() {
        const idx = partners.length + 1;
        partners.push({
            id: nextUid++,
            name: `Partner ${idx}`,
            role: 'Partner',
            ratio: 0,
            open: 1000000,
            add: 0,
            draw: 0
        });
        renderPartnersEditor();
        window.updateReport();
    };

    window.removePartner = function(id) {
        if (partners.length <= 1) return;
        partners = partners.filter(p => p.id !== id);
        renderPartnersEditor();
        window.updateReport();
    };

    window.updatePartnerField = function(id, field, value) {
        const p = partners.find(item => item.id === id);
        if (!p) return;
        p[field] = (field === 'name' || field === 'role') ? value : (parseFloat(value) || 0);
        window.updateReport();
    };

    function renderPartnersEditor() {
        const container = document.getElementById('partners-editor-list');
        if (!container) return;
        container.innerHTML = partners.map((p, index) => `
            <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200 space-y-2.5">
                <div class="flex justify-between items-center">
                    <span class="text-xs font-bold text-slate-700">Partner ${index + 1}</span>
                    ${partners.length > 1 ? `<button type="button" onclick="removePartner(${p.id})" class="text-[11px] text-rose-600 hover:text-rose-800 font-medium">Remove</button>` : ''}
                </div>
                <div class="grid grid-cols-2 gap-2">
                    <div>
                        <label class="block text-[11px] text-slate-500">Full Name</label>
                        <input type="text" value="${escapeHtml(p.name)}" oninput="updatePartnerField(${p.id}, 'name', this.value)" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                    </div>
                    <div>
                        <label class="block text-[11px] text-slate-500">Role / Responsibility</label>
                        <input type="text" value="${escapeHtml(p.role)}" oninput="updatePartnerField(${p.id}, 'role', this.value)" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                    </div>
                </div>
                <div class="grid grid-cols-3 gap-2">
                    <div>
                        <label class="block text-[11px] text-slate-500">Profit Share (%)</label>
                        <input type="number" value="${p.ratio}" oninput="updatePartnerField(${p.id}, 'ratio', this.value)" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                    </div>
                    <div>
                        <label class="block text-[11px] text-slate-500">Opening Cap (BDT)</label>
                        <input type="number" value="${p.open}" oninput="updatePartnerField(${p.id}, 'open', this.value)" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                    </div>
                    <div>
                        <label class="block text-[11px] text-slate-500">Added Cap (BDT)</label>
                        <input type="number" value="${p.add}" oninput="updatePartnerField(${p.id}, 'add', this.value)" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                    </div>
                </div>
                <div>
                    <label class="block text-[11px] text-slate-500">Drawings During Year (BDT)</label>
                    <input type="number" value="${p.draw}" oninput="updatePartnerField(${p.id}, 'draw', this.value)" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                </div>
            </div>
        `).join('');
    }

    // P&L Operating Expenses CRUD
    window.addPnlExpense = function() {
        pnlExpenses.push({
            id: nextUid++,
            label: 'New operating expense',
            note: 'Note 4',
            curr: 0,
            prev: 0
        });
        renderPnlExpensesEditor();
        window.updateReport();
    };

    window.removePnlExpense = function(id) {
        pnlExpenses = pnlExpenses.filter(e => e.id !== id);
        renderPnlExpensesEditor();
        window.updateReport();
    };

    window.updatePnlExpense = function(id, field, value) {
        const item = pnlExpenses.find(e => e.id === id);
        if (!item) return;
        item[field] = (field === 'label' || field === 'note') ? value : (parseFloat(value) || 0);
        window.updateReport();
    };

    function renderPnlExpensesEditor() {
        const container = document.getElementById('pnl-expenses-list');
        if (!container) return;
        container.innerHTML = pnlExpenses.map(e => `
            <div class="grid grid-cols-12 gap-1 items-center">
                <input type="text" value="${escapeHtml(e.label)}" oninput="updatePnlExpense(${e.id}, 'label', this.value)" class="col-span-5 px-2 py-1.5 text-xs border border-slate-300 rounded" title="Account Head">
                <input type="text" value="${escapeHtml(e.note)}" oninput="updatePnlExpense(${e.id}, 'note', this.value)" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded" title="Note Ref">
                <input type="number" value="${e.curr}" oninput="updatePnlExpense(${e.id}, 'curr', this.value)" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded" title="Current Year">
                <input type="number" value="${e.prev}" oninput="updatePnlExpense(${e.id}, 'prev', this.value)" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50" title="Prior Year">
                <button type="button" onclick="removePnlExpense(${e.id})" class="col-span-1 text-center text-rose-500 hover:text-rose-700 font-bold text-sm" title="Delete Line">×</button>
            </div>
        `).join('');
    }

    // Cash Flow CRUD
    window.addCfItem = function(section) {
        const list = section === 'investing' ? cfInvesting : cfFinancing;
        list.push({
            id: nextUid++,
            label: section === 'investing' ? 'Other investing cash flow' : 'Other financing cash flow',
            amount: 0
        });
        renderCfEditors();
        window.updateReport();
    };

    window.removeCfItem = function(section, id) {
        if (section === 'investing') {
            cfInvesting = cfInvesting.filter(i => i.id !== id);
        } else {
            cfFinancing = cfFinancing.filter(i => i.id !== id);
        }
        renderCfEditors();
        window.updateReport();
    };

    window.updateCfItem = function(section, id, field, value) {
        const list = section === 'investing' ? cfInvesting : cfFinancing;
        const item = list.find(i => i.id === id);
        if (!item) return;
        item[field] = field === 'label' ? value : (parseFloat(value) || 0);
        window.updateReport();
    };

    function renderCfEditors() {
        const renderList = (section, list, containerId) => {
            const el = document.getElementById(containerId);
            if (!el) return;
            el.innerHTML = list.map(item => `
                <div class="grid grid-cols-12 gap-1.5 items-center">
                    <input type="text" value="${escapeHtml(item.label)}" oninput="updateCfItem('${section}', ${item.id}, 'label', this.value)" class="col-span-8 px-2 py-1.5 text-xs border border-slate-300 rounded">
                    <input type="number" value="${item.amount}" oninput="updateCfItem('${section}', ${item.id}, 'amount', this.value)" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded">
                    <button type="button" onclick="removeCfItem('${section}', ${item.id})" class="col-span-1 text-center text-rose-500 hover:text-rose-700 font-bold text-sm" title="Delete">×</button>
                </div>
            `).join('');
        };
        renderList('investing', cfInvesting, 'cf-investing-list');
        renderList('financing', cfFinancing, 'cf-financing-list');
    }

    // Balance Sheet CRUD
    window.addBsItem = function(group) {
        bsItems[group].push({
            id: nextUid++,
            label: 'New account head',
            note: 'Note 4',
            curr: 0,
            prev: 0
        });
        renderBsEditors();
        window.updateReport();
    };

    window.removeBsItem = function(group, id) {
        bsItems[group] = bsItems[group].filter(i => i.id !== id);
        renderBsEditors();
        window.updateReport();
    };

    window.updateBsItem = function(group, id, field, value) {
        const item = bsItems[group].find(i => i.id === id);
        if (!item) return;
        item[field] = (field === 'label' || field === 'note') ? value : (parseFloat(value) || 0);
        window.updateReport();
    };

    function renderBsEditors() {
        ['nca', 'ca', 'ncl', 'cl'].forEach(group => {
            const container = document.getElementById(`bs-${group}-list`);
            if (!container) return;
            container.innerHTML = bsItems[group].map(item => `
                <div class="grid grid-cols-12 gap-1 items-center">
                    <input type="text" value="${escapeHtml(item.label)}" oninput="updateBsItem('${group}', ${item.id}, 'label', this.value)" class="col-span-5 px-2 py-1.5 text-xs border border-slate-300 rounded" title="Account Head">
                    <input type="text" value="${escapeHtml(item.note)}" oninput="updateBsItem('${group}', ${item.id}, 'note', this.value)" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded" title="Note Ref">
                    <input type="number" value="${item.curr}" oninput="updateBsItem('${group}', ${item.id}, 'curr', this.value)" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded" title="Current Year">
                    <input type="number" value="${item.prev}" oninput="updateBsItem('${group}', ${item.id}, 'prev', this.value)" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50" title="Prior Year">
                    <button type="button" onclick="removeBsItem('${group}', ${item.id})" class="col-span-1 text-center text-rose-500 hover:text-rose-700 font-bold text-sm" title="Delete Line">×</button>
                </div>
            `).join('');
        });
    }

    // Disclosure Notes CRUD
    window.addDisclosureNote = function() {
        const num = disclosureNotes.length + 26;
        disclosureNotes.push({
            id: nextUid++,
            title: `Note ${num}: Additional Disclosure`,
            text: 'Enter explanatory accounting policy or quantitative disclosure details here.'
        });
        renderNotesEditor();
        window.updateReport();
    };

    window.removeDisclosureNote = function(id) {
        disclosureNotes = disclosureNotes.filter(n => n.id !== id);
        renderNotesEditor();
        window.updateReport();
    };

    window.toggleStructuredNote = function(group, key, show) {
        if (group === '8_14') structuredNotes8to14[key].visible = show;
        if (group === '15_25') structuredNotes15to25[key].visible = show;
        window.updateReport();
    };

    window.addStructuredNoteRow = function(group, key) {
        if (key === 'n8') {
            structuredNotes8to14.n8.rows.push({ id: nextUid++, cat: 'New asset category', cost: 0, dep: 0 });
        } else if (key === 'n19') {
            structuredNotes15to25.n19.extraRows.push({ id: nextUid++, sl: '', item: 'Additional tax adjustment', ref: 'ITA 2023', base: '—', amt: 0 });
        } else {
            const target = group === '8_14' ? structuredNotes8to14[key] : structuredNotes15to25[key];
            target.rows.push({ id: nextUid++, label: 'New line item', curr: 0, prev: 0 });
        }
        window.updateReport();
    };

    window.removeStructuredNoteRow = function(group, key, id) {
        if (key === 'n8') {
            structuredNotes8to14.n8.rows = structuredNotes8to14.n8.rows.filter(r => r.id !== id);
        } else if (key === 'n19') {
            structuredNotes15to25.n19.extraRows = structuredNotes15to25.n19.extraRows.filter(r => r.id !== id);
        } else {
            const target = group === '8_14' ? structuredNotes8to14[key] : structuredNotes15to25[key];
            target.rows = target.rows.filter(r => r.id !== id);
        }
        window.updateReport();
    };

    window.updateStructuredNoteRow = function(group, key, id, field, rawVal) {
        if (key === 'n8') {
            const r = structuredNotes8to14.n8.rows.find(x => x.id === id);
            if (!r) return;
            if (field === 'cat') r.cat = rawVal;
            else r[field] = parseFloat(String(rawVal).replace(/,/g, '')) || 0;
        } else if (key === 'n19') {
            const r = structuredNotes15to25.n19.extraRows.find(x => x.id === id);
            if (!r) return;
            if (field === 'amt') r.amt = parseFloat(String(rawVal).replace(/,/g, '')) || 0;
            else r[field] = rawVal;
        } else {
            const target = group === '8_14' ? structuredNotes8to14[key] : structuredNotes15to25[key];
            const r = target.rows.find(x => x.id === id);
            if (!r) return;
            if (field === 'label') r.label = rawVal;
            else r[field] = parseFloat(String(rawVal).replace(/,/g, '')) || 0;
        }
        window.updateReport();
    };

    window.updateStructuredNoteText = function(group, key, field, rawVal) {
        const target = group === '8_14' ? structuredNotes8to14[key] : structuredNotes15to25[key];
        if (target) target[field] = rawVal;
        window.updateReport();
    };

    window.updateDisclosureNote = function(id, field, value) {
        const note = disclosureNotes.find(n => n.id === id);
        if (!note) return;
        note[field] = value;
        window.updateReport();
    };

    function renderNotesEditor() {
        const container = document.getElementById('notes-editor-list');
        if (!container) return;
        container.innerHTML = disclosureNotes.map(n => `
            <div class="p-2.5 bg-slate-50 border border-slate-200 rounded-lg space-y-1.5">
                <div class="flex items-center gap-2">
                    <input type="text" value="${escapeHtml(n.title)}" oninput="updateDisclosureNote(${n.id}, 'title', this.value)" class="flex-1 px-2 py-1 text-xs font-semibold border border-slate-300 rounded bg-white">
                    <button type="button" onclick="removeDisclosureNote(${n.id})" class="text-xs text-rose-600 hover:text-rose-800 font-medium px-1">Delete</button>
                </div>
                <textarea rows="2" oninput="updateDisclosureNote(${n.id}, 'text', this.value)" class="w-full px-2 py-1 text-xs border border-slate-300 rounded bg-white">${escapeHtml(n.text)}</textarea>
            </div>
        `).join('');
    }

    // Tab & Mobile View Navigation
    window.switchTab = function(tabName) {
        ['firm', 'pnl', 'bs', 'mgmt', 'tax', 'it116'].forEach(t => {
            const panel = document.getElementById('panel-' + t);
            const btn = document.getElementById('tab-' + t);
            if (panel) panel.classList.add('hidden');
            if (btn) {
                btn.classList.remove('border-emerald-600', 'text-emerald-700', 'bg-white');
                btn.classList.add('border-transparent', 'text-slate-600');
            }
        });

        const activePanel = document.getElementById('panel-' + tabName);
        const activeBtn = document.getElementById('tab-' + tabName);
        if (activePanel) activePanel.classList.remove('hidden');
        if (activeBtn) {
            activeBtn.classList.add('border-emerald-600', 'text-emerald-700', 'bg-white');
            activeBtn.classList.remove('border-transparent', 'text-slate-600');
        }
    };

    function initIt116TinBoxes() {
        document.querySelectorAll('.it116-tin-container').forEach(container => {
            container.innerHTML = '';
            for (let i = 0; i < 12; i++) {
                const inp = document.createElement('input');
                inp.type = 'text';
                inp.maxLength = 1;
                inp.className = 'it116-tin-box';
                inp.addEventListener('input', e => {
                    const boxes = Array.from(container.querySelectorAll('.it116-tin-box'));
                    if (e.target.value.length === 1 && i < boxes.length - 1) {
                        boxes[i + 1].focus();
                    }
                    const combined = boxes.map(b => b.value || '').join('');
                    const tinEl = document.getElementById('it116Tin');
                    if (tinEl) tinEl.value = combined;
                    window.syncTinFromSidebar(combined, container);
                });
                inp.addEventListener('keydown', e => {
                    const boxes = Array.from(container.querySelectorAll('.it116-tin-box'));
                    if (e.key === 'Backspace' && e.target.value.length === 0 && i > 0) {
                        boxes[i - 1].focus();
                    }
                });
                container.appendChild(inp);
            }
        });
    }

    window.syncTinFromSidebar = function(tinStr, skipContainer = null) {
        const digits = String(tinStr || '').replace(/\D/g, '').slice(0, 12);
        document.querySelectorAll('.it116-tin-container').forEach(container => {
            if (container === skipContainer) return;
            const boxes = container.querySelectorAll('.it116-tin-box');
            boxes.forEach((box, idx) => {
                box.value = digits[idx] || '';
            });
        });
    };

    window.loadBwdbTaxSample = function() {
        document.getElementById('taxPeriodTitle').value = 'FY 2025-2026 (As at 30 June 2026)';
        document.getElementById('taxTurnoverSlab').value = '0.5';
        document.getElementById('taxGrossReceipts').value = '940434.00';
        document.getElementById('taxInvoiceRef').value = 'Contract Invoices (BWDB Bill 1952)';
        document.getElementById('taxableProfit').value = '139143.50';
        document.getElementById('regularTaxRate').value = '27.50';
        document.getElementById('taxPbtNote').value = 'PBT (Note 11.00)';
        document.getElementById('taxAitPaid').value = '47022.00';
        document.getElementById('taxAitRef').value = 'BWDB Bill 1952 Token 00324450';
        document.getElementById('deferredTaxAmt').value = '0.00';
        document.getElementById('taxClosingLabel').value = 'TOTAL CURRENT TAX LIABILITY AS AT 30 JUNE 2026';
        document.getElementById('taxStatutoryNote').value = 'Regular tax @ 27.5% on taxable profit is BDT 38,264.46 (higher than 0.50% minimum tax of BDT 4,702.17). Advance Income Tax (AIT / TDS) of BDT 47,022.00 was deducted at source by BWDB under Section 89 of the Income Tax Act 2023. Under Section 163(2)(b), tax deducted at source serves as the minimum tax, leaving a Net Current Tax Liability of BDT 0.00 at June 30, 2026.';
        window.updateReport();
    };

    let mobileShowingPreview = false;
    window.toggleMobileView = function() {
        const sidebar = document.getElementById('editor-sidebar');
        const preview = document.getElementById('preview-container');
        const btn = document.getElementById('toggle-view-btn');

        mobileShowingPreview = !mobileShowingPreview;
        if (mobileShowingPreview) {
            sidebar.classList.add('hidden');
            preview.classList.remove('hidden');
            btn.textContent = 'Back to Editor';
        } else {
            sidebar.classList.remove('hidden');
            preview.classList.add('hidden');
            btn.textContent = 'Switch to Preview';
        }
    };

    window.resetDemoData = function() {
        if (onResetCallback) {
            onResetCallback();
        } else {
            window.location.reload();
        }
    };

    let printTimeoutId = null;

    function showExportToast(titleText, bodyText, durationMs = 2400) {
        const toast = document.getElementById('print-toast-notification');
        const titleEl = document.getElementById('print-toast-title');
        const bodyEl = document.getElementById('print-toast-body');
        if (titleEl) titleEl.textContent = titleText;
        if (bodyEl) bodyEl.textContent = bodyText;
        if (toast) toast.classList.remove('hidden');

        if (printTimeoutId) clearTimeout(printTimeoutId);
        if (durationMs > 0) {
            printTimeoutId = setTimeout(() => {
                if (toast) toast.classList.add('hidden');
            }, durationMs);
        }
    }

    function getVisibleA4Sheets() {
        const allSheets = Array.from(document.querySelectorAll('#preview-container .a4-sheet'));
        return allSheets.filter(sheet => {
            const el = sheet as HTMLElement;
            if (el.style.display === 'none') return false;
            const parentWrapper = el.closest('#it116-return-wrapper') as HTMLElement | null;
            if (parentWrapper && parentWrapper.style.display === 'none') return false;
            return true;
        }) as HTMLElement[];
    }

    function prepareClonedSheetForExport(originalSheet: HTMLElement): HTMLElement {
        const clone = originalSheet.cloneNode(true) as HTMLElement;

        // Remove all non-printing UI controls and action columns
        clone.querySelectorAll('.no-print, button.no-print, th.no-print, td.no-print').forEach(el => el.remove());
        clone.querySelectorAll('[contenteditable]').forEach(el => el.removeAttribute('contenteditable'));

        // Sync live input/textarea/select values from the original sheet into the clone
        const origInputs = Array.from(originalSheet.querySelectorAll('input, textarea, select'));
        const cloneInputs = Array.from(clone.querySelectorAll('input, textarea, select'));

        origInputs.forEach((origEl, idx) => {
            const cloneEl = cloneInputs[idx];
            if (!cloneEl) return;

            if (origEl.tagName === 'INPUT') {
                const inp = origEl as HTMLInputElement;
                if (inp.type === 'checkbox') {
                    const span = document.createElement('span');
                    span.textContent = inp.checked ? ' [√] ' : ' [ ] ';
                    span.style.fontWeight = 'bold';
                    cloneEl.replaceWith(span);
                } else if (inp.classList.contains('it116-tin-box')) {
                    const span = document.createElement('span');
                    span.textContent = inp.value || ' ';
                    span.style.display = 'inline-block';
                    span.style.width = '18pt';
                    span.style.height = '20pt';
                    span.style.lineHeight = '20pt';
                    span.style.border = '1pt solid #000000';
                    span.style.textAlign = 'center';
                    span.style.fontWeight = 'bold';
                    span.style.fontFamily = 'Courier New, monospace';
                    cloneEl.replaceWith(span);
                } else {
                    const span = document.createElement('span');
                    span.textContent = inp.value || '';
                    if (inp.classList.contains('it116-line-input') || inp.classList.contains('it11cola-input')) {
                        span.style.borderBottom = '1pt dotted #000000';
                        span.style.padding = '0 4pt';
                        span.style.fontWeight = inp.classList.contains('font-semibold') || inp.classList.contains('font-bold') ? 'bold' : 'normal';
                    }
                    cloneEl.replaceWith(span);
                }
            } else if (origEl.tagName === 'TEXTAREA') {
                const ta = origEl as HTMLTextAreaElement;
                const div = document.createElement('div');
                div.style.whiteSpace = 'pre-wrap';
                div.textContent = ta.value || '';
                cloneEl.replaceWith(div);
            } else if (origEl.tagName === 'SELECT') {
                const sel = origEl as HTMLSelectElement;
                const span = document.createElement('span');
                span.textContent = sel.options[sel.selectedIndex]?.text || sel.value || '';
                cloneEl.replaceWith(span);
            }
        });

        return clone;
    }

    function triggerBlobDownload(blob: Blob, filename: string) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 1500);
    }

    function buildCurrentExportPayload() {
        const firmName = (strVal('firmName') || 'Partnership Firm').trim();
        const reportDate = (strVal('reportDate') || '31 December 2026').trim();
        const prevDate = (strVal('prevDate') || '31 December 2025').trim();
        const regNo = strVal('regNo');
        const estDate = strVal('estDate');
        const address = strVal('address');
        const businessNature = strVal('businessNature');
        const approvalDate = strVal('approvalDate');
        const includeAudit = Boolean(document.getElementById('includeAudit')?.checked);
        const auditorFirm = strVal('auditorFirm');
        const auditorPartner = strVal('auditorPartner');

        const revLabel = strVal('revLabel') || 'Revenue';
        const revNote = strVal('revNote');
        const revCurr = val('revCurr'), revPrev = val('revPrev');
        const cosLabel = strVal('cosLabel') || 'Cost of sales';
        const cosNote = strVal('cosNote');
        const cosCurr = val('cosCurr'), cosPrev = val('cosPrev');
        const gpCurr = revCurr - cosCurr, gpPrev = revPrev - cosPrev;
        const othIncLabel = strVal('othIncLabel') || 'Other income';
        const othIncNote = strVal('othIncNote');
        const othIncCurr = val('othIncCurr'), othIncPrev = val('othIncPrev');

        let totOpExpCurr = 0, totOpExpPrev = 0;
        pnlExpenses.forEach(e => {
            totOpExpCurr += e.curr;
            totOpExpPrev += e.prev;
        });

        const opCurr = gpCurr + othIncCurr - totOpExpCurr;
        const opPrev = gpPrev + othIncPrev - totOpExpPrev;
        const finLabel = strVal('finLabel') || 'Finance costs';
        const finNote = strVal('finNote');
        const finCurr = val('finCostCurr'), finPrev = val('finCostPrev');
        const pbtCurr = opCurr - finCurr, pbtPrev = opPrev - finPrev;

        const includeTax = Boolean(document.getElementById('includeTaxSchedule')?.checked);
        const syncTaxToPnl = Boolean(document.getElementById('syncTaxToPnl')?.checked);
        const taxPeriodTitle = strVal('taxPeriodTitle');
        const slabRatePct = parseFloat(strVal('taxTurnoverSlab')) || 0;
        const grossReceipts = val('taxGrossReceipts');
        const taxInvoiceRef = strVal('taxInvoiceRef');
        const taxableProfit = val('taxableProfit');
        const regTaxRatePct = val('regularTaxRate');
        const taxPbtNote = strVal('taxPbtNote');
        const aitPaid = val('taxAitPaid');
        const taxAitRef = strVal('taxAitRef');
        const deferredTaxAmt = val('deferredTaxAmt');
        const taxClosingLabel = strVal('taxClosingLabel');
        const taxStatutoryNote = strVal('taxStatutoryNote');

        const minTaxA = +(grossReceipts * (slabRatePct / 100)).toFixed(2);
        const regTaxB = +(Math.max(0, taxableProfit) * (regTaxRatePct / 100)).toFixed(2);
        const applicableTax = Math.max(minTaxA, regTaxB);
        const netCurrentTaxPayable = Math.max(0, +(applicableTax - aitPaid).toFixed(2));
        const totalTaxExpenseIas12 = +(applicableTax + deferredTaxAmt).toFixed(2);

        const taxLabel = strVal('taxLabel') || 'Income tax expense';
        const taxNote = strVal('taxNote');
        const taxCurr = syncTaxToPnl ? Math.round(totalTaxExpenseIas12) : val('taxCurr');
        const taxPrev = val('taxPrev');

        const npCurr = pbtCurr - taxCurr, npPrev = pbtPrev - taxPrev;
        const ociLabel = strVal('ociLabel') || 'Other comprehensive income';
        const ociCurr = val('ociCurr'), ociPrev = val('ociPrev');
        const tciCurr = npCurr + ociCurr, tciPrev = npPrev + ociPrev;

        const ratioSum = partners.reduce((acc, p) => acc + p.ratio, 0);
        let allocatedSoFar = 0;
        const computedPartners = partners.map((p, idx) => {
            let profShare = 0;
            if (ratioSum > 0) {
                if (idx === partners.length - 1 && ratioSum === 100) {
                    profShare = tciCurr - allocatedSoFar;
                } else {
                    profShare = Math.round(tciCurr * (p.ratio / 100));
                    allocatedSoFar += profShare;
                }
            }
            const close = p.open + p.add + profShare - p.draw;
            return { ...p, prof: profShare, close };
        });

        const totOpen = computedPartners.reduce((s, p) => s + p.open, 0);
        const totAdd = computedPartners.reduce((s, p) => s + p.add, 0);
        const totProf = computedPartners.reduce((s, p) => s + p.prof, 0);
        const totDraw = computedPartners.reduce((s, p) => s + p.draw, 0);
        const totClose = computedPartners.reduce((s, p) => s + p.close, 0);

        const ncaCurr = bsItems.nca.reduce((s, i) => s + i.curr, 0);
        const ncaPrev = bsItems.nca.reduce((s, i) => s + i.prev, 0);
        const caCurr = bsItems.ca.reduce((s, i) => s + i.curr, 0);
        const caPrev = bsItems.ca.reduce((s, i) => s + i.prev, 0);
        const totAssetsCurr = ncaCurr + caCurr, totAssetsPrev = ncaPrev + caPrev;
        const eqCurr = totClose, eqPrev = totOpen;
        const nclCurr = bsItems.ncl.reduce((s, i) => s + i.curr, 0);
        const nclPrev = bsItems.ncl.reduce((s, i) => s + i.prev, 0);
        const clCurr = bsItems.cl.reduce((s, i) => s + i.curr, 0);
        const clPrev = bsItems.cl.reduce((s, i) => s + i.prev, 0);
        const totLiabCurr = nclCurr + clCurr, totLiabPrev = nclPrev + clPrev;
        const totEqLiabCurr = eqCurr + totLiabCurr, totEqLiabPrev = eqPrev + totLiabPrev;

        const cfOpsGen = val('cfOpsGen');
        const netOps = cfOpsGen - finCurr - taxCurr;
        const netInv = cfInvesting.reduce((s, i) => s + i.amount, 0);
        const otherFin = cfFinancing.reduce((s, i) => s + i.amount, 0);
        const netFin = totAdd - totDraw + otherFin;
        const netCashChange = netOps + netInv + netFin;
        const cfOpenCash = val('cfOpenCash');
        const cfCloseCash = cfOpenCash + netCashChange;

        const gpm = revCurr > 0 ? ((gpCurr / revCurr) * 100).toFixed(1) : '0.0';
        const npm = revCurr > 0 ? ((npCurr / revCurr) * 100).toFixed(1) : '0.0';
        const currRatio = clCurr > 0 ? (caCurr / clCurr).toFixed(2) : '0.00';
        const roe = eqCurr > 0 ? ((npCurr / eqCurr) * 100).toFixed(1) : '0.0';

        const includeIt116 = Boolean(document.getElementById('includeIt116')?.checked);
        const rentInc = val('it116RentInc');
        const agriInc = val('it116AgriInc');
        const bizInc = pbtCurr;
        const capGain = val('it116CapGain');
        const finInc = val('it116FinInc');
        const othSrcInc = 0;
        const foreignInc = val('it116ForeignInc');
        const totalIncome = rentInc + agriInc + bizInc + capGain + finInc + othSrcInc + foreignInc;
        const regularTaxOnTotalInc = taxCurr;
        const minTaxOnTds = val('it116Tds');
        const turnoverTaxCalc = Math.round(revCurr * (slabRatePct / 100));
        const taxPayableHigher = Math.max(regularTaxOnTotalInc, minTaxOnTds, turnoverTaxCalc);
        const envSurcharge = val('it116EnvSurcharge');
        const penaltyAmt = val('it116Penalty');
        const totalTaxPayable = taxPayableHigher + envSurcharge + penaltyAmt;
        const tdsPaid = val('it116Tds');
        const aitPaidReturn = val('it116Ait');
        const balanceTaxWithReturn = Math.max(0, totalTaxPayable - tdsPaid - aitPaidReturn);
        const totalTaxPaid = tdsPaid + aitPaidReturn + balanceTaxWithReturn;
        const excessPayment = Math.max(0, totalTaxPaid - totalTaxPayable);

        return {
            firmName,
            reportDate,
            prevDate,
            regNo,
            estDate,
            address,
            businessNature,
            approvalDate,
            includeAudit,
            auditorFirm,
            auditorPartner,
            txtOps: strVal('txtOps'),
            txtRisks: strVal('txtRisks'),
            txtFuture: strVal('txtFuture'),
            txtResp: strVal('txtResp'),
            partners: computedPartners,
            totals: { ratioSum, totOpen, totAdd, totProf, totDraw, totClose },
            pnl: {
                revLabel, revNote, revCurr, revPrev,
                cosLabel, cosNote, cosCurr, cosPrev,
                gpCurr, gpPrev,
                othIncLabel, othIncNote, othIncCurr, othIncPrev,
                expenses: pnlExpenses,
                opCurr, opPrev,
                finLabel, finNote, finCurr, finPrev,
                pbtCurr, pbtPrev,
                taxLabel, taxNote, taxCurr, taxPrev,
                npCurr, npPrev,
                ociLabel, ociCurr, ociPrev,
                tciCurr, tciPrev
            },
            bs: {
                nca: bsItems.nca,
                ca: bsItems.ca,
                ncl: bsItems.ncl,
                cl: bsItems.cl,
                ncaCurr, ncaPrev,
                caCurr, caPrev,
                totAssetsCurr, totAssetsPrev,
                eqCurr, eqPrev,
                nclCurr, nclPrev,
                clCurr, clPrev,
                totLiabCurr, totLiabPrev,
                totEqLiabCurr, totEqLiabPrev
            },
            cf: {
                cfOpsGen, netOps,
                investing: cfInvesting, netInv,
                financing: cfFinancing, netFin,
                netCashChange, cfOpenCash, cfCloseCash
            },
            ratios: { gpm, npm, currRatio, roe },
            notes8to14: structuredNotes8to14,
            notes15to25: structuredNotes15to25,
            disclosureNotes,
            tax: {
                includeTax, taxPeriodTitle, slabRatePct, grossReceipts, taxInvoiceRef,
                taxableProfit, regTaxRatePct, taxPbtNote, aitPaid, taxAitRef,
                deferredTaxAmt, taxClosingLabel, taxStatutoryNote,
                minTaxA, regTaxB, applicableTax, netCurrentTaxPayable
            },
            it116: {
                includeIt116,
                tin: strVal('it116Tin'),
                bin: strVal('it116Bin'),
                circle: strVal('it116Circle'),
                zone: strVal('it116Zone'),
                taxYear: strVal('it116TaxYear'),
                category: strVal('it116Category'),
                resident: strVal('it116Resident'),
                regDate: strVal('it116RegDate'),
                phone: strVal('it116Phone'),
                email: strVal('it116Email'),
                rentInc, agriInc, bizInc, capGain, finInc, othSrcInc, foreignInc, totalIncome,
                regularTaxOnTotalInc, minTaxOnTds, turnoverTaxCalc, taxPayableHigher,
                envSurcharge, penaltyAmt, totalTaxPayable,
                tdsPaid, aitPaidReturn, balanceTaxWithReturn, totalTaxPaid, excessPayment,
                exemptInc: val('it116ExemptInc'),
                docsList: strVal('it116DocsList')
            }
        };
    }

    window.exportReportToWord = async function() {
        if (document.activeElement && typeof (document.activeElement as HTMLElement).blur === 'function') {
            (document.activeElement as HTMLElement).blur();
        }
        window.updateReport();

        showExportToast(
            'Generating Word Document (.docx)…',
            'Mapping Financial Statements, Notes, IAS 12 Tax Schedule & IT-116 Return into an A4 .docx file.',
            2600
        );

        const payload = buildCurrentExportPayload();
        await downloadReportAsDocx(payload);
    };

    window.exportReportToExcel = function() {
        if (document.activeElement && typeof (document.activeElement as HTMLElement).blur === 'function') {
            (document.activeElement as HTMLElement).blur();
        }
        window.updateReport();

        showExportToast(
            'Generating Excel Workbook (.xlsx)…',
            'Exporting multi-sheet .xlsx workbook (Overview, Balance Sheet, P&L, Cash Flow, Notes, IAS 12 Tax & IT-116).',
            2600
        );

        const payload = buildCurrentExportPayload();
        downloadReportAsXlsx(payload);
    };

    window.printReportToPdf = function() {
        // 1. Commit any active inline edits and remove focus rings
        if (document.activeElement && typeof document.activeElement.blur === 'function') {
            document.activeElement.blur();
        }

        // 2. Ensure latest calculations are rendered
        window.updateReport();

        // 3. Show pre-print reminder notification briefly before opening the print dialog
        showExportToast(
            'Preparing A4 PDF Print Dialog…',
            "Ensure printer settings are set to A4 portrait and 'Background graphics' are enabled for best results.",
            0
        );
        const toast = document.getElementById('print-toast-notification');

        printTimeoutId = setTimeout(() => {
            // 4. Set clean PDF filename via document.title
            const originalTitle = document.title;
            const firmName = (strVal('firmName') || 'Partnership_Firm').trim();
            const reportDate = (strVal('reportDate') || 'FY2026').trim();
            document.title = `${firmName} - IFRS Financial Report & IT-116 (${reportDate})`;

            // 5. Ensure preview container is visible even if triggered from mobile editor view
            const preview = document.getElementById('preview-container');
            const wasPreviewHidden = preview ? preview.classList.contains('hidden') : false;
            if (preview && wasPreviewHidden) {
                preview.classList.remove('hidden');
            }

            // 6. Mark the last visible .a4-sheet so no trailing blank page is generated when optional sections are hidden
            const allSheets = Array.from(document.querySelectorAll('.a4-sheet'));
            allSheets.forEach(sheet => {
                sheet.style.pageBreakAfter = '';
                sheet.style.breakAfter = '';
            });
            const visibleSheets = getVisibleA4Sheets();
            if (visibleSheets.length > 0) {
                const lastVisible = visibleSheets[visibleSheets.length - 1] as HTMLElement;
                lastVisible.style.setProperty('page-break-after', 'auto', 'important');
                lastVisible.style.setProperty('break-after', 'auto', 'important');
            }

            const cleanupAfterPrint = () => {
                document.title = originalTitle;
                if (toast) {
                    toast.classList.add('hidden');
                }
                if (preview && wasPreviewHidden && !mobileShowingPreview) {
                    preview.classList.add('hidden');
                }
                window.removeEventListener('afterprint', cleanupAfterPrint);
            };

            window.addEventListener('afterprint', cleanupAfterPrint);

            // 7. Trigger browser native print-to-PDF dialog
            window.focus();
            window.print();

            // Fallback hide for the toast in case afterprint fires immediately
            setTimeout(() => {
                if (toast) toast.classList.add('hidden');
            }, 500);
        }, 1200);
    };

    // Core Calculation & Preview Render Engine
    window.updateReport = function() {
        const firmName = strVal('firmName') || 'Partnership Firm';
        const reportDate = strVal('reportDate');
        const prevDate = strVal('prevDate');
        const regNo = strVal('regNo');
        const estDate = strVal('estDate');
        const address = strVal('address');
        const businessNature = strVal('businessNature');
        const approvalDate = strVal('approvalDate');

        setText('doc-cover-firm', firmName);
        setText('doc-cover-date', reportDate);
        setText('doc-cover-reg', regNo);
        setText('doc-cover-addr', address);
        setText('doc-est-date', estDate);
        setText('doc-biz-nature', businessNature);

        setClassText('doc-firm-name', firmName);
        setClassText('doc-report-date', reportDate);
        setClassText('doc-prev-date', prevDate);
        setClassText('doc-reg-no', regNo);
        setClassText('doc-addr', address);
        setClassText('doc-est-date', estDate);
        setClassText('doc-biz-nature', businessNature);
        setClassText('doc-approval-date', approvalDate);

        setText('doc-txt-ops', strVal('txtOps'));
        setText('doc-txt-risks', strVal('txtRisks'));
        setText('doc-txt-future', strVal('txtFuture'));
        setText('doc-txt-resp', strVal('txtResp'));

        const includeAuditEl = document.getElementById('includeAudit');
        const includeAudit = includeAuditEl ? includeAuditEl.checked : true;
        const docAuditBlock = document.getElementById('doc-audit-block');
        const tocAuditSection = document.getElementById('toc-audit-section');
        const auditFields = document.getElementById('audit-fields');
        if (docAuditBlock) docAuditBlock.style.display = includeAudit ? 'flex' : 'none';
        if (tocAuditSection) tocAuditSection.style.display = includeAudit ? 'block' : 'none';
        if (auditFields) auditFields.style.opacity = includeAudit ? '1' : '0.4';
        setText('doc-auditor-firm', strVal('auditorFirm'));
        setText('doc-auditor-partner', strVal('auditorPartner'));

        // IAS 12 & S.R.O. 324/2026 Corporate Tax Assessment Engine
        const includeTaxEl = document.getElementById('includeTaxSchedule');
        const syncTaxEl = document.getElementById('syncTaxToPnl');
        const includeTax = includeTaxEl ? includeTaxEl.checked : true;
        const syncTaxToPnl = syncTaxEl ? syncTaxEl.checked : false;
        const docTaxPage = document.getElementById('doc-tax-page');
        const tocTaxRow = document.getElementById('toc-tax-row');
        if (docTaxPage) docTaxPage.style.display = includeTax ? 'flex' : 'none';
        if (tocTaxRow) tocTaxRow.style.display = includeTax ? 'flex' : 'none';

        const taxPeriodTitle = strVal('taxPeriodTitle');
        const slabRatePct = parseFloat(strVal('taxTurnoverSlab')) || 0;
        const grossReceipts = val('taxGrossReceipts');
        const taxInvoiceRef = strVal('taxInvoiceRef');
        const taxableProfit = val('taxableProfit');
        const regTaxRatePct = val('regularTaxRate');
        const taxPbtNote = strVal('taxPbtNote');
        const aitPaid = val('taxAitPaid');
        const taxAitRef = strVal('taxAitRef');
        const deferredTaxAmt = val('deferredTaxAmt');
        const taxClosingLabel = strVal('taxClosingLabel');
        const taxStatutoryNote = strVal('taxStatutoryNote');

        const minTaxA = +(grossReceipts * (slabRatePct / 100)).toFixed(2);
        const regTaxB = +(Math.max(0, taxableProfit) * (regTaxRatePct / 100)).toFixed(2);
        const applicableTax = Math.max(minTaxA, regTaxB);
        const netCurrentTaxPayable = Math.max(0, +(applicableTax - aitPaid).toFixed(2));
        const advanceTaxCarryForward = aitPaid > applicableTax ? +(aitPaid - applicableTax).toFixed(2) : 0;
        const totalTaxExpenseIas12 = +(applicableTax + deferredTaxAmt).toFixed(2);

        setText('doc-tax-period', taxPeriodTitle);
        setText('doc-tax-statutory-note', taxStatutoryNote);

        const slabNameMap = {
            '0': 'first slab (Up to Tk. 20M), qualifying for a 0.00%',
            '0.5': 'second slab (> Tk. 20M to Tk. 40M), subjecting it to a 0.50%',
            '1': 'third slab (> Tk. 40M), subjecting it to a 1.00%',
            '1.0': 'third slab (> Tk. 40M), subjecting it to a 1.00%'
        };
        const slabDesc = slabNameMap[String(slabRatePct)] || `selected slab (${slabRatePct.toFixed(2)}%)`;
        const slabDescEl = document.getElementById('doc-tax-slab-desc');
        if (slabDescEl) {
            slabDescEl.innerHTML = `The entity's contract turnover falls into the ${slabDesc} minimum turnover tax. Corporate income tax is determined as the higher of Minimum Tax (A) and Regular Tax @ ${regTaxRatePct.toFixed(2)}% on Taxable Profit (B).`;
        }

        const taxTbody = document.getElementById('tax-schedule-tbody');
        if (taxTbody) {
            taxTbody.innerHTML = `
                <tr>
                    <td class="border border-slate-300 px-2 py-1.5 text-center font-medium">1</td>
                    <td class="border border-slate-300 px-2 py-1.5 font-medium">Gross Receipts / Turnover from Contracts with Customers</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">S.R.O. 324/2026</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">${escapeHtml(taxInvoiceRef)}</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono">${fmtDec(grossReceipts)}</td>
                </tr>
                <tr>
                    <td class="border border-slate-300 px-2 py-1.5 text-center font-medium">2</td>
                    <td class="border border-slate-300 px-2 py-1.5 font-medium">Minimum Income Tax @ ${slabRatePct.toFixed(2)}% of Gross Receipts (A)</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">Section 163(6) / S.R.O. 324</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">${fmtDec(grossReceipts)} @ ${slabRatePct.toFixed(2)}%</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono font-semibold">${fmtDec(minTaxA)}</td>
                </tr>
                <tr>
                    <td class="border border-slate-300 px-2 py-1.5 text-center font-medium">3</td>
                    <td class="border border-slate-300 px-2 py-1.5 font-medium">Taxable Business Profit for the Period</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">Income Statement</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">${escapeHtml(taxPbtNote)}</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono">${fmtDec(taxableProfit)}</td>
                </tr>
                <tr>
                    <td class="border border-slate-300 px-2 py-1.5 text-center font-medium">4</td>
                    <td class="border border-slate-300 px-2 py-1.5 font-medium">Tax Applied on Taxable Profit @ ${regTaxRatePct.toFixed(1)}% (B)</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">Regular Tax Rate @ ${regTaxRatePct.toFixed(2)}%</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">${fmtDec(taxableProfit)} @ ${regTaxRatePct.toFixed(2)}%</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono font-semibold">${fmtDec(regTaxB)}</td>
                </tr>
                <tr class="bg-slate-100 font-bold">
                    <td class="border border-slate-300 px-2 py-1.5 text-center">—</td>
                    <td class="border border-slate-300 px-2 py-1.5 uppercase">Applicable Corporate Income Tax (Higher of A and B)</td>
                    <td class="border border-slate-300 px-2 py-1.5">Higher of (A) & (B)</td>
                    <td class="border border-slate-300 px-2 py-1.5">${regTaxB >= minTaxA ? 'Regular Tax (B) Applies' : 'Minimum Tax (A) Applies'}</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono">${fmtDec(applicableTax)}</td>
                </tr>
                <tr>
                    <td class="border border-slate-300 px-2 py-1.5 text-center font-medium">5</td>
                    <td class="border border-slate-300 px-2 py-1.5 font-medium">Less: Advance Tax Paid (AIT) / TDS Deducted at Source</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">Section 163 / AIT Challan</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">${escapeHtml(taxAitRef)}</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono">${fmtDec(aitPaid, true)}</td>
                </tr>
                ${deferredTaxAmt !== 0 ? `
                <tr>
                    <td class="border border-slate-300 px-2 py-1.5 text-center font-medium">6</td>
                    <td class="border border-slate-300 px-2 py-1.5 font-medium">Deferred Tax Expense / (Income) Movement (IAS 12)</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">IAS 12 Temporary Diff.</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-slate-600">Deferred Tax Recon.</td>
                    <td class="border border-slate-300 px-2 py-1.5 text-right font-mono">${fmtDec(deferredTaxAmt)}</td>
                </tr>` : ''}
                <tr class="bg-slate-900 text-white font-bold">
                    <td class="border border-slate-400 px-2 py-2 text-center">—</td>
                    <td class="border border-slate-400 px-2 py-2 uppercase">${escapeHtml(taxClosingLabel)}</td>
                    <td class="border border-slate-400 px-2 py-2 font-normal">Net Payable to NBR</td>
                    <td class="border border-slate-400 px-2 py-2 font-normal">${netCurrentTaxPayable === 0 ? 'Balance Due (TDS covers liability)' : 'Net Tax Payable'}</td>
                    <td class="border border-slate-400 px-2 py-2 text-right font-mono">${fmtDec(netCurrentTaxPayable)}</td>
                </tr>
                ${advanceTaxCarryForward > 0 ? `
                <tr class="bg-slate-50 text-slate-700 text-[10px]">
                    <td class="border border-slate-300 px-2 py-1 text-center">i</td>
                    <td class="border border-slate-300 px-2 py-1 italic" colspan="3">Memorandum: Excess AIT / TDS u/s 163(2)(b) (Treated as Minimum Tax / Unadjusted Credit)</td>
                    <td class="border border-slate-300 px-2 py-1 text-right font-mono">${fmtDec(advanceTaxCarryForward)}</td>
                </tr>` : ''}
            `;
        }

        const taxInputEl = document.getElementById('taxCurr');
        const taxSyncHint = document.getElementById('tax-sync-hint');
        if (taxInputEl && taxSyncHint) {
            if (syncTaxToPnl) {
                taxInputEl.value = Math.round(totalTaxExpenseIas12);
                taxInputEl.readOnly = true;
                taxInputEl.classList.add('bg-emerald-50', 'font-semibold', 'text-emerald-900');
                taxSyncHint.classList.remove('hidden');
            } else {
                taxInputEl.readOnly = false;
                taxInputEl.classList.remove('bg-emerald-50', 'font-semibold', 'text-emerald-900');
                taxSyncHint.classList.add('hidden');
            }
        }

        // Profit or Loss Calculations
        const revLabel = strVal('revLabel') || 'Revenue';
        const revNote = strVal('revNote');
        const revCurr = val('revCurr'), revPrev = val('revPrev');

        const cosLabel = strVal('cosLabel') || 'Cost of sales';
        const cosNote = strVal('cosNote');
        const cosCurr = val('cosCurr'), cosPrev = val('cosPrev');

        const gpCurr = revCurr - cosCurr;
        const gpPrev = revPrev - cosPrev;

        const othIncLabel = strVal('othIncLabel') || 'Other income';
        const othIncNote = strVal('othIncNote');
        const othIncCurr = val('othIncCurr'), othIncPrev = val('othIncPrev');

        let totOpExpCurr = 0, totOpExpPrev = 0;
        pnlExpenses.forEach(e => {
            totOpExpCurr += e.curr;
            totOpExpPrev += e.prev;
        });

        const opCurr = gpCurr + othIncCurr - totOpExpCurr;
        const opPrev = gpPrev + othIncPrev - totOpExpPrev;

        const finLabel = strVal('finLabel') || 'Finance costs';
        const finNote = strVal('finNote');
        const finCurr = val('finCostCurr'), finPrev = val('finCostPrev');

        const pbtCurr = opCurr - finCurr;
        const pbtPrev = opPrev - finPrev;

        const taxLabel = strVal('taxLabel') || 'Income tax expense';
        const taxNote = strVal('taxNote');
        const taxCurr = val('taxCurr'), taxPrev = val('taxPrev');

        const npCurr = pbtCurr - taxCurr;
        const npPrev = pbtPrev - taxPrev;

        const ociLabel = strVal('ociLabel') || 'Other comprehensive income for the year';
        const ociCurr = val('ociCurr'), ociPrev = val('ociPrev');

        const tciCurr = npCurr + ociCurr;
        const tciPrev = npPrev + ociPrev;

        const plTbody = document.getElementById('pl-preview-tbody');
        if (plTbody) {
            plTbody.innerHTML = `
                <tr>
                    <td class="py-2 px-2 font-semibold">${escapeHtml(revLabel)}</td>
                    <td class="py-2 px-2 text-center text-slate-500">${escapeHtml(revNote)}</td>
                    <td class="py-2 px-2 text-right font-mono font-semibold">${fmt(revCurr)}</td>
                    <td class="py-2 px-2 text-right font-mono">${fmt(revPrev)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 px-2">Less: ${escapeHtml(cosLabel)}</td>
                    <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(cosNote)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(cosCurr, true)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(cosPrev, true)}</td>
                </tr>
                <tr class="font-bold bg-slate-50">
                    <td class="py-2 px-2">Gross Profit</td>
                    <td></td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(gpCurr)}</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(gpPrev)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 px-2">${escapeHtml(othIncLabel)}</td>
                    <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(othIncNote)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(othIncCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(othIncPrev)}</td>
                </tr>
                ${pnlExpenses.map(e => `
                    <tr>
                        <td class="py-1.5 px-2">${escapeHtml(e.label)}</td>
                        <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(e.note)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(e.curr, true)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(e.prev, true)}</td>
                    </tr>
                `).join('')}
                <tr class="font-bold bg-slate-50">
                    <td class="py-2 px-2">Operating Profit</td>
                    <td></td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(opCurr)}</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(opPrev)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 px-2">${escapeHtml(finLabel)}</td>
                    <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(finNote)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(finCurr, true)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(finPrev, true)}</td>
                </tr>
                <tr class="font-bold">
                    <td class="py-2 px-2">Profit Before Tax</td>
                    <td></td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(pbtCurr)}</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(pbtPrev)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 px-2">${escapeHtml(taxLabel)}</td>
                    <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(taxNote)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(taxCurr, true)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(taxPrev, true)}</td>
                </tr>
                <tr class="font-bold text-slate-900 bg-slate-100">
                    <td class="py-2 px-2">Profit for the Year</td>
                    <td></td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-500">${fmt(npCurr)}</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-500">${fmt(npPrev)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 px-2">${escapeHtml(ociLabel)}</td>
                    <td></td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(ociCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(ociPrev)}</td>
                </tr>
                <tr class="font-bold text-slate-900 bg-slate-100 border-y-2 border-slate-900">
                    <td class="py-2.5 px-2 uppercase">Total Comprehensive Income for the Year</td>
                    <td></td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(tciCurr)}</td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(tciPrev)}</td>
                </tr>
            `;
        }

        // Dynamic Partners' Calculations & Tables
        const ratioSum = partners.reduce((acc, p) => acc + p.ratio, 0);
        const badge = document.getElementById('ratio-badge');
        if (badge) {
            badge.textContent = `Total Ratio: ${ratioSum}%`;
            badge.className = ratioSum === 100
                ? 'text-xs font-semibold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800'
                : 'text-xs font-semibold px-2 py-0.5 rounded bg-rose-100 text-rose-800';
        }

        let allocatedSoFar = 0;
        const computedPartners = partners.map((p, idx) => {
            let profShare = 0;
            if (ratioSum > 0) {
                if (idx === partners.length - 1 && ratioSum === 100) {
                    profShare = tciCurr - allocatedSoFar;
                } else {
                    profShare = Math.round(tciCurr * (p.ratio / 100));
                    allocatedSoFar += profShare;
                }
            }
            const close = p.open + p.add + profShare - p.draw;
            return { ...p, prof: profShare, close };
        });

        const totOpen = computedPartners.reduce((s, p) => s + p.open, 0);
        const totAdd = computedPartners.reduce((s, p) => s + p.add, 0);
        const totProf = computedPartners.reduce((s, p) => s + p.prof, 0);
        const totDraw = computedPartners.reduce((s, p) => s + p.draw, 0);
        const totClose = computedPartners.reduce((s, p) => s + p.close, 0);

        const mrStructEl = document.getElementById('mr-structure-table');
        if (mrStructEl) mrStructEl.innerHTML = `
            <thead class="bg-slate-100">
                <tr>
                    <th class="border border-slate-300 px-2.5 py-1 text-left">Particulars</th>
                    ${computedPartners.map(p => `<th class="border border-slate-300 px-2.5 py-1 text-right">${escapeHtml(p.name)}</th>`).join('')}
                    <th class="border border-slate-300 px-2.5 py-1 text-right">Total</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="border border-slate-300 px-2.5 py-1 font-medium">Designation</td>
                    ${computedPartners.map(p => `<td class="border border-slate-300 px-2.5 py-1 text-right">${escapeHtml(p.role)}</td>`).join('')}
                    <td class="border border-slate-300 px-2.5 py-1 text-right text-slate-500">—</td>
                </tr>
                <tr>
                    <td class="border border-slate-300 px-2.5 py-1 font-medium">Profit-sharing ratio</td>
                    ${computedPartners.map(p => `<td class="border border-slate-300 px-2.5 py-1 text-right">${p.ratio}%</td>`).join('')}
                    <td class="border border-slate-300 px-2.5 py-1 text-right font-semibold">${ratioSum}%</td>
                </tr>
                <tr class="bg-slate-50 font-bold">
                    <td class="border border-slate-300 px-2.5 py-1">Closing capital (BDT)</td>
                    ${computedPartners.map(p => `<td class="border border-slate-300 px-2.5 py-1 text-right font-mono">${fmt(p.close)}</td>`).join('')}
                    <td class="border border-slate-300 px-2.5 py-1 text-right font-mono">${fmt(totClose)}</td>
                </tr>
            </tbody>
        `;

        const defaultResponsibilities = [
            [
                'Establishing business strategies, operational objectives, and organizational priorities.',
                'Identifying new commercial opportunities and developing institutional relationships.',
                'Supervising major procurement, trading, and distribution activities.',
                'Evaluating business expansion opportunities and strategic investments.',
                'Monitoring overall business performance and representing the Firm in negotiations.'
            ],
            [
                'Supervising accounting records, financial reporting, and financial documentation.',
                'Monitoring cash flows, working capital, and financial obligations.',
                'Coordinating budgeting, expenditure management, and financial planning.',
                'Monitoring receivables, payables, and 45-day credit management procedures.',
                'Supporting annual financial statement preparation and internal controls.'
            ]
        ];
        const mrRolesEl = document.getElementById('mr-partner-roles-grid');
        if (mrRolesEl) mrRolesEl.innerHTML = computedPartners.map((p, idx) => {
            const bullets = defaultResponsibilities[idx] || [
                'Participating in management governance and strategic partnership decisions.',
                'Overseeing assigned operational units, compliance, and resource administration.'
            ];
            return `
                <div class="bg-slate-50 border border-slate-200 rounded p-2.5">
                    <p class="font-bold text-slate-900 text-[11px]">${escapeHtml(p.name)} — ${escapeHtml(p.role)}</p>
                    <ul class="list-disc pl-4 mt-1 space-y-0.5 text-[10px] text-slate-700">
                        ${bullets.map(b => `<li>${b}</li>`).join('')}
                    </ul>
                </div>
            `;
        }).join('');

        const mrCapDistEl = document.getElementById('mr-cap-dist-table');
        if (mrCapDistEl) mrCapDistEl.innerHTML = `
            <thead class="bg-slate-100">
                <tr>
                    <th class="border border-slate-300 px-2.5 py-1 text-left">Particulars</th>
                    ${computedPartners.map(p => `<th class="border border-slate-300 px-2.5 py-1 text-right">${escapeHtml(p.name)} (${p.ratio}%)</th>`).join('')}
                    <th class="border border-slate-300 px-2.5 py-1 text-right">Total (BDT)</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="border border-slate-300 px-2.5 py-1">Net Profit Allocated (BDT)</td>
                    ${computedPartners.map(p => `<td class="border border-slate-300 px-2.5 py-1 text-right font-mono">${fmt(p.prof)}</td>`).join('')}
                    <td class="border border-slate-300 px-2.5 py-1 text-right font-mono font-semibold">${fmt(totProf)}</td>
                </tr>
                <tr class="bg-slate-50 font-bold">
                    <td class="border border-slate-300 px-2.5 py-1">Closing Capital Balance (BDT)</td>
                    ${computedPartners.map(p => `<td class="border border-slate-300 px-2.5 py-1 text-right font-mono">${fmt(p.close)}</td>`).join('')}
                    <td class="border border-slate-300 px-2.5 py-1 text-right font-mono">${fmt(totClose)}</td>
                </tr>
            </tbody>
        `;

        const mrSignoffEl = document.getElementById('mr-signoff-grid');
        if (mrSignoffEl) mrSignoffEl.innerHTML = computedPartners.map(p => `
            <div>
                <div class="border-b border-slate-700 w-40 mb-1.5 pt-4"></div>
                <p class="font-bold text-slate-900 text-[11px]">${escapeHtml(p.name)}</p>
                <p class="text-[10px] text-slate-600">${escapeHtml(p.role)}</p>
            </div>
        `).join('');

        const plAllocEl = document.getElementById('pl-alloc-table');
        if (plAllocEl) plAllocEl.innerHTML = `
            ${computedPartners.map(p => `
                <tr class="border-b border-slate-300">
                    <td class="py-1">${escapeHtml(p.name)} (${p.ratio}%)</td>
                    <td class="py-1 text-right font-mono font-semibold">${fmt(p.prof)}</td>
                </tr>
            `).join('')}
            <tr class="font-bold">
                <td class="pt-1.5">Total Profit Allocated to Partners' Equity</td>
                <td class="pt-1.5 text-right font-mono">${fmt(totProf)}</td>
            </tr>
        `;

        const buildEquityTableHtml = (bordered) => {
            const cellCls = bordered ? 'border border-slate-300 px-2.5 py-1.5' : 'py-2 px-2';
            const headCls = bordered ? 'border border-slate-300 px-2.5 py-1.5 bg-slate-100' : 'py-2.5 px-2 bg-slate-50';
            return `
                <thead>
                    <tr class="${bordered ? '' : 'border-y-2 border-slate-900'}">
                        <th class="${headCls} text-left font-bold">Particulars</th>
                        ${computedPartners.map(p => `<th class="${headCls} text-right font-bold">${escapeHtml(p.name)}<br><span class="font-normal text-[10px]">(BDT)</span></th>`).join('')}
                        <th class="${headCls} text-right font-bold">Total Equity<br><span class="font-normal text-[10px]">(BDT)</span></th>
                    </tr>
                </thead>
                <tbody class="${bordered ? '' : 'divide-y divide-slate-200'}">
                    <tr>
                        <td class="${cellCls} font-medium">Opening capital balance</td>
                        ${computedPartners.map(p => `<td class="${cellCls} text-right font-mono">${fmt(p.open)}</td>`).join('')}
                        <td class="${cellCls} text-right font-mono font-semibold">${fmt(totOpen)}</td>
                    </tr>
                    <tr>
                        <td class="${cellCls}">Additional capital introduced</td>
                        ${computedPartners.map(p => `<td class="${cellCls} text-right font-mono">${fmt(p.add)}</td>`).join('')}
                        <td class="${cellCls} text-right font-mono">${fmt(totAdd)}</td>
                    </tr>
                    <tr>
                        <td class="${cellCls}">Profit allocated (${computedPartners.map(p => p.ratio + '%').join(' / ')})</td>
                        ${computedPartners.map(p => `<td class="${cellCls} text-right font-mono">${fmt(p.prof)}</td>`).join('')}
                        <td class="${cellCls} text-right font-mono">${fmt(totProf)}</td>
                    </tr>
                    <tr>
                        <td class="${cellCls}">Less: Partners' drawings</td>
                        ${computedPartners.map(p => `<td class="${cellCls} text-right font-mono">${fmt(p.draw, true)}</td>`).join('')}
                        <td class="${cellCls} text-right font-mono">${fmt(totDraw, true)}</td>
                    </tr>
                    <tr class="font-bold text-slate-900 bg-slate-100 ${bordered ? '' : 'border-y-2 border-slate-900'}">
                        <td class="${cellCls} uppercase">Closing Capital Balance</td>
                        ${computedPartners.map(p => `<td class="${cellCls} text-right font-mono">${fmt(p.close)}</td>`).join('')}
                        <td class="${cellCls} text-right font-mono">${fmt(totClose)}</td>
                    </tr>
                </tbody>
            `;
        };

        const eqPreviewEl = document.getElementById('eq-preview-table');
        if (eqPreviewEl) eqPreviewEl.innerHTML = buildEquityTableHtml(false);

        const sigPreviewEl = document.getElementById('signatures-preview-grid');
        if (sigPreviewEl) sigPreviewEl.innerHTML = computedPartners.map(p => `
            <div class="bg-white border border-slate-200 rounded p-3">
                <p class="font-bold text-slate-900 text-xs">${escapeHtml(p.name)}</p>
                <p class="text-[11px] text-slate-600">${escapeHtml(p.role)}</p>
                <p class="text-[10.5px] text-slate-500 mb-6">${escapeHtml(firmName)}</p>
                <div class="border-b border-slate-800 w-full mb-1.5"></div>
                <p class="text-[10px] text-slate-600">Signature</p>
                <p class="text-[10px] text-slate-500 mt-0.5">Date: ${escapeHtml(approvalDate)} | Place: Dhaka, Bangladesh</p>
            </div>
        `).join('');

        // Dynamic Statement of Financial Position (Balance Sheet)
        const ncaCurr = bsItems.nca.reduce((s, i) => s + i.curr, 0);
        const ncaPrev = bsItems.nca.reduce((s, i) => s + i.prev, 0);

        const caCurr = bsItems.ca.reduce((s, i) => s + i.curr, 0);
        const caPrev = bsItems.ca.reduce((s, i) => s + i.prev, 0);

        const totAssetsCurr = ncaCurr + caCurr;
        const totAssetsPrev = ncaPrev + caPrev;

        const cashItem = bsItems.ca.find(i => i.isCash || i.label.toLowerCase().includes('cash'));
        const cashCurr = cashItem ? cashItem.curr : 0;

        const eqCurr = totClose;
        const eqPrev = totOpen;

        const nclCurr = bsItems.ncl.reduce((s, i) => s + i.curr, 0);
        const nclPrev = bsItems.ncl.reduce((s, i) => s + i.prev, 0);

        const clCurr = bsItems.cl.reduce((s, i) => s + i.curr, 0);
        const clPrev = bsItems.cl.reduce((s, i) => s + i.prev, 0);

        const totLiabCurr = nclCurr + clCurr;
        const totLiabPrev = nclPrev + clPrev;

        const totEqLiabCurr = eqCurr + totLiabCurr;
        const totEqLiabPrev = eqPrev + totLiabPrev;

        const bsTbody = document.getElementById('bs-preview-tbody');
        if (bsTbody) {
            bsTbody.innerHTML = `
                <tr class="bg-slate-100/70"><td colspan="4" class="py-2 px-2 font-bold uppercase tracking-wider text-slate-800">Assets</td></tr>
                <tr><td colspan="4" class="py-1.5 px-2 font-semibold text-slate-700">Non-current assets</td></tr>
                ${bsItems.nca.map(i => `
                    <tr>
                        <td class="py-1.5 pl-6 pr-2">${escapeHtml(i.label)}</td>
                        <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(i.note)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.curr)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.prev)}</td>
                    </tr>
                `).join('')}
                <tr class="font-semibold bg-slate-50/50">
                    <td class="py-1.5 pl-4 pr-2">Total non-current assets</td>
                    <td></td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(ncaCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(ncaPrev)}</td>
                </tr>

                <tr><td colspan="4" class="py-1.5 px-2 font-semibold text-slate-700 pt-3">Current assets</td></tr>
                ${bsItems.ca.map(i => `
                    <tr>
                        <td class="py-1.5 pl-6 pr-2">${escapeHtml(i.label)}</td>
                        <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(i.note)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.curr)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.prev)}</td>
                    </tr>
                `).join('')}
                <tr class="font-semibold bg-slate-50/50">
                    <td class="py-1.5 pl-4 pr-2">Total current assets</td>
                    <td></td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(caCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(caPrev)}</td>
                </tr>

                <tr class="font-bold text-slate-900 bg-slate-100 border-y-2 border-slate-900">
                    <td class="py-2.5 px-2 uppercase">Total Assets</td>
                    <td></td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(totAssetsCurr)}</td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(totAssetsPrev)}</td>
                </tr>

                <tr class="bg-slate-100/70"><td colspan="4" class="py-2 px-2 font-bold uppercase tracking-wider text-slate-800 pt-4">Equity and Liabilities</td></tr>
                <tr><td colspan="4" class="py-1.5 px-2 font-semibold text-slate-700">Partners' Equity (IAS 32)</td></tr>
                <tr>
                    <td class="py-1.5 pl-6 pr-2">Partners' capital and current accounts</td>
                    <td class="py-1.5 px-2 text-center text-slate-500">Note 12</td>
                    <td class="py-1.5 px-2 text-right font-mono font-semibold">${fmt(eqCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono font-semibold">${fmt(eqPrev)}</td>
                </tr>

                <tr><td colspan="4" class="py-1.5 px-2 font-semibold text-slate-700 pt-2">Non-current liabilities</td></tr>
                ${bsItems.ncl.map(i => `
                    <tr>
                        <td class="py-1.5 pl-6 pr-2">${escapeHtml(i.label)}</td>
                        <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(i.note)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.curr)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.prev)}</td>
                    </tr>
                `).join('')}
                <tr class="font-semibold bg-slate-50/50">
                    <td class="py-1.5 pl-4 pr-2">Total non-current liabilities</td>
                    <td></td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(nclCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(nclPrev)}</td>
                </tr>

                <tr><td colspan="4" class="py-1.5 px-2 font-semibold text-slate-700 pt-2">Current liabilities</td></tr>
                ${bsItems.cl.map(i => `
                    <tr>
                        <td class="py-1.5 pl-6 pr-2">${escapeHtml(i.label)}</td>
                        <td class="py-1.5 px-2 text-center text-slate-500">${escapeHtml(i.note)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.curr)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.prev)}</td>
                    </tr>
                `).join('')}
                <tr class="font-semibold bg-slate-50/50">
                    <td class="py-1.5 pl-4 pr-2">Total current liabilities</td>
                    <td></td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(clCurr)}</td>
                    <td class="py-1.5 px-2 text-right font-mono border-t border-slate-400">${fmt(clPrev)}</td>
                </tr>

                <tr class="font-bold text-slate-900 bg-slate-100 border-y-2 border-slate-900">
                    <td class="py-2.5 px-2 uppercase">Total Equity and Liabilities</td>
                    <td></td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(totEqLiabCurr)}</td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(totEqLiabPrev)}</td>
                </tr>
            `;
        }

        const diffCurr = totAssetsCurr - totEqLiabCurr;
        const diffPrev = totAssetsPrev - totEqLiabPrev;
        const bsAlert = document.getElementById('bs-balance-alert');
        if (bsAlert) {
            if (diffCurr === 0 && diffPrev === 0) {
                bsAlert.className = 'p-3 rounded-lg text-xs font-medium border bg-emerald-50 border-emerald-200 text-emerald-800';
                bsAlert.innerHTML = '✓ <strong>Balanced (IAS 1):</strong> Assets equal Equity + Liabilities for both Current and Prior periods.';
            } else {
                bsAlert.className = 'p-3 rounded-lg text-xs font-medium border bg-amber-50 border-amber-300 text-amber-900';
                bsAlert.innerHTML = `⚠ <strong>Out of Balance:</strong> Current Variance: BDT ${fmt(diffCurr)} | Prior Variance: BDT ${fmt(diffPrev)}. Adjust asset or liability items to reconcile.`;
            }
        }

        // Dynamic Statement of Cash Flows (IAS 7)
        const cfOpsGen = val('cfOpsGen');
        const netOps = cfOpsGen - finCurr - taxCurr;

        const netInv = cfInvesting.reduce((s, i) => s + i.amount, 0);
        const otherFin = cfFinancing.reduce((s, i) => s + i.amount, 0);
        const netFin = totAdd - totDraw + otherFin;

        const netCashChange = netOps + netInv + netFin;
        const cfOpenCash = val('cfOpenCash');
        const cfCloseCash = cfOpenCash + netCashChange;

        const cfTbody = document.getElementById('cf-preview-tbody');
        if (cfTbody) {
            cfTbody.innerHTML = `
                <tr class="bg-slate-100/70"><td colspan="2" class="py-2 px-2 font-bold uppercase text-slate-800">A. Cash Flows from Operating Activities</td></tr>
                <tr>
                    <td class="py-1.5 pl-5 pr-2">Cash generated from operations</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(cfOpsGen)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 pl-5 pr-2">Interest paid</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(finCurr, true)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 pl-5 pr-2">Income taxes paid</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(taxCurr, true)}</td>
                </tr>
                <tr class="font-bold bg-slate-50">
                    <td class="py-2 px-2">Net Cash from Operating Activities</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(netOps)}</td>
                </tr>

                <tr class="bg-slate-100/70"><td colspan="2" class="py-2 px-2 font-bold uppercase text-slate-800 pt-3">B. Cash Flows from Investing Activities</td></tr>
                ${cfInvesting.map(i => `
                    <tr>
                        <td class="py-1.5 pl-5 pr-2">${escapeHtml(i.label)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.amount)}</td>
                    </tr>
                `).join('')}
                <tr class="font-bold bg-slate-50">
                    <td class="py-2 px-2">Net Cash from / (Used in) Investing Activities</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(netInv)}</td>
                </tr>

                <tr class="bg-slate-100/70"><td colspan="2" class="py-2 px-2 font-bold uppercase text-slate-800 pt-3">C. Cash Flows from Financing Activities</td></tr>
                <tr>
                    <td class="py-1.5 pl-5 pr-2">Capital introduced by partners</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(totAdd)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 pl-5 pr-2">Partners' withdrawals (drawings)</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(totDraw, true)}</td>
                </tr>
                ${cfFinancing.map(i => `
                    <tr>
                        <td class="py-1.5 pl-5 pr-2">${escapeHtml(i.label)}</td>
                        <td class="py-1.5 px-2 text-right font-mono">${fmt(i.amount)}</td>
                    </tr>
                `).join('')}
                <tr class="font-bold bg-slate-50">
                    <td class="py-2 px-2">Net Cash from / (Used in) Financing Activities</td>
                    <td class="py-2 px-2 text-right font-mono border-t border-slate-400">${fmt(netFin)}</td>
                </tr>

                <tr class="font-bold text-slate-900">
                    <td class="py-2 px-2 pt-3">Net Increase / (Decrease) in Cash and Cash Equivalents (A + B + C)</td>
                    <td class="py-2 px-2 pt-3 text-right font-mono border-t border-slate-400">${fmt(netCashChange)}</td>
                </tr>
                <tr>
                    <td class="py-1.5 px-2">Opening cash and cash equivalents</td>
                    <td class="py-1.5 px-2 text-right font-mono">${fmt(cfOpenCash)}</td>
                </tr>
                <tr class="font-bold text-slate-900 bg-slate-100 border-y-2 border-slate-900">
                    <td class="py-2.5 px-2 uppercase">Closing Cash and Cash Equivalents</td>
                    <td class="py-2.5 px-2 text-right font-mono">${fmt(cfCloseCash)}</td>
                </tr>
            `;
        }

        const cfRecNote = document.getElementById('cf-reconcile-note');
        if (cfRecNote) {
            if (cfCloseCash === cashCurr) {
                cfRecNote.className = 'mt-2 text-[11px] text-emerald-700 font-medium';
                cfRecNote.textContent = '✓ Closing cash reconciles exactly with Cash and Cash Equivalents in the Statement of Financial Position.';
            } else {
                cfRecNote.className = 'mt-2 text-[11px] text-amber-700 font-medium';
                cfRecNote.textContent = `⚠ Note: Closing cash (BDT ${fmt(cfCloseCash)}) differs from Balance Sheet Cash (BDT ${fmt(cashCurr)}) by BDT ${fmt(cashCurr - cfCloseCash)}.`;
            }
        }

        // Helper for Two-Column Disclosure Note Tables
        const renderTwoColRows = (group, key, rows) => rows.map(r => {
            if (r.isHeader) {
                return `
                    <tr class="bg-slate-50 font-semibold">
                        <td class="border border-slate-300 px-1.5 py-0.5" colspan="3" contenteditable="true" onblur="updateStructuredNoteRow('${group}','${key}',${r.id},'label',this.textContent)">${escapeHtml(r.label)}</td>
                        <td class="border border-slate-300 px-1 py-0.5 text-center no-print"><button type="button" onclick="removeStructuredNoteRow('${group}','${key}',${r.id})" class="text-rose-500 hover:text-rose-700 font-bold">×</button></td>
                    </tr>`;
            }
            return `
                <tr class="${r.isTotal ? 'font-bold bg-slate-100' : ''}">
                    <td class="border border-slate-300 px-1.5 py-0.5" contenteditable="true" onblur="updateStructuredNoteRow('${group}','${key}',${r.id},'label',this.textContent)">${escapeHtml(r.label)}</td>
                    <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono" contenteditable="true" onblur="updateStructuredNoteRow('${group}','${key}',${r.id},'curr',this.textContent)">${fmt(r.curr)}</td>
                    <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono" contenteditable="true" onblur="updateStructuredNoteRow('${group}','${key}',${r.id},'prev',this.textContent)">${fmt(r.prev)}</td>
                    <td class="border border-slate-300 px-1 py-0.5 text-center no-print"><button type="button" onclick="removeStructuredNoteRow('${group}','${key}',${r.id})" class="text-rose-500 hover:text-rose-700 font-bold">×</button></td>
                </tr>`;
        }).join('');

        const noteHeaderBar = (group, key, title) => `
            <div class="flex items-center justify-between">
                <h4 class="font-bold text-slate-900">${escapeHtml(title)}</h4>
                <button type="button" onclick="toggleStructuredNote('${group}','${key}',false)" class="no-print text-[9px] text-rose-600 hover:underline">Delete this note</button>
            </div>`;

        const addRowBtn = (group, key) => `
            <button type="button" onclick="addStructuredNoteRow('${group}','${key}')" class="no-print mt-0.5 text-[9px] font-semibold text-emerald-700 hover:text-emerald-900">+ Add row</button>`;

        const n8 = structuredNotes8to14.n8;
        const n9 = structuredNotes8to14.n9;
        const n10 = structuredNotes8to14.n10;
        const n11 = structuredNotes8to14.n11;
        const n12 = structuredNotes8to14.n12;
        const n13 = structuredNotes8to14.n13;
        const n14 = structuredNotes8to14.n14;

        const ratioText = computedPartners.map(p => p.ratio + '%').join(' : ');
        const n14Body = n14.text || `Profit for the year of BDT ${fmt(tciCurr)} has been allocated among the partners in accordance with the profit-sharing ratio stipulated in the partnership deed of ${ratioText} respectively. The allocation basis is reviewed annually and no change was made during the reporting period.`;

        const notes814El = document.getElementById('notes-8-14-container');
        if (notes814El) {
            notes814El.innerHTML = `
                ${n8.visible ? `
                <div>
                    ${noteHeaderBar('8_14', 'n8', n8.title)}
                    <table class="w-full text-[9.5px] border-collapse border border-slate-300 mt-0.5">
                        <thead class="bg-slate-100">
                            <tr>
                                <th class="border border-slate-300 px-1.5 py-0.5 text-left">Category</th>
                                <th class="border border-slate-300 px-1.5 py-0.5 text-right w-24">Cost BDT</th>
                                <th class="border border-slate-300 px-1.5 py-0.5 text-right w-28">Accumulated Depreciation BDT</th>
                                <th class="border border-slate-300 px-1.5 py-0.5 text-right w-24">Carrying Amount BDT</th>
                                <th class="border border-slate-300 px-1 py-0.5 w-5 no-print">×</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${n8.rows.map(r => `
                                <tr class="${r.isTotal ? 'font-bold bg-slate-100' : ''}">
                                    <td class="border border-slate-300 px-1.5 py-0.5" contenteditable="true" onblur="updateStructuredNoteRow('8_14','n8',${r.id},'cat',this.textContent)">${escapeHtml(r.cat)}</td>
                                    <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono" contenteditable="true" onblur="updateStructuredNoteRow('8_14','n8',${r.id},'cost',this.textContent)">${fmt(r.cost)}</td>
                                    <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono" contenteditable="true" onblur="updateStructuredNoteRow('8_14','n8',${r.id},'dep',this.textContent)">${fmt(r.dep)}</td>
                                    <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(r.cost + r.dep)}</td>
                                    <td class="border border-slate-300 px-1 py-0.5 text-center no-print"><button type="button" onclick="removeStructuredNoteRow('8_14','n8',${r.id})" class="text-rose-500 hover:text-rose-700 font-bold">×</button></td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    ${addRowBtn('8_14', 'n8')}
                    <p class="text-slate-700 mt-0.5" contenteditable="true" onblur="updateStructuredNoteText('8_14','n8','footer',this.textContent)">${escapeHtml(n8.footer)}</p>
                </div>` : ''}

                <div class="grid grid-cols-1 gap-1.5">
                    ${n9.visible ? `
                    <div>
                        ${noteHeaderBar('8_14', 'n9', n9.title)}
                        <table class="w-full text-[9.5px] border-collapse border border-slate-300 mt-0.5">
                            <thead class="bg-slate-100">
                                <tr>
                                    <th class="border border-slate-300 px-1.5 py-0.5 text-left">Particulars</th>
                                    <th class="border border-slate-300 px-1.5 py-0.5 text-right w-24">2026 BDT</th>
                                    <th class="border border-slate-300 px-1.5 py-0.5 text-right w-24">2025 BDT</th>
                                    <th class="border border-slate-300 px-1 py-0.5 w-5 no-print">×</th>
                                </tr>
                            </thead>
                            <tbody>${renderTwoColRows('8_14', 'n9', n9.rows)}</tbody>
                        </table>
                        ${addRowBtn('8_14', 'n9')}
                    </div>` : ''}

                    <div class="grid grid-cols-2 gap-2">
                        ${n10.visible ? `
                        <div>
                            ${noteHeaderBar('8_14', 'n10', n10.title)}
                            <table class="w-full text-[9.5px] border-collapse border border-slate-300 mt-0.5">
                                <thead class="bg-slate-100">
                                    <tr>
                                        <th class="border border-slate-300 px-1.5 py-0.5 text-left">Particulars</th>
                                        <th class="border border-slate-300 px-1.5 py-0.5 text-right w-20">2026 BDT</th>
                                        <th class="border border-slate-300 px-1.5 py-0.5 text-right w-20">2025 BDT</th>
                                        <th class="border border-slate-300 px-1 py-0.5 w-4 no-print">×</th>
                                    </tr>
                                </thead>
                                <tbody>${renderTwoColRows('8_14', 'n10', n10.rows)}</tbody>
                            </table>
                            ${addRowBtn('8_14', 'n10')}
                        </div>` : ''}

                        ${n11.visible ? `
                        <div>
                            ${noteHeaderBar('8_14', 'n11', n11.title)}
                            <table class="w-full text-[9.5px] border-collapse border border-slate-300 mt-0.5">
                                <thead class="bg-slate-100">
                                    <tr>
                                        <th class="border border-slate-300 px-1.5 py-0.5 text-left">Particulars</th>
                                        <th class="border border-slate-300 px-1.5 py-0.5 text-right w-20">2026 BDT</th>
                                        <th class="border border-slate-300 px-1.5 py-0.5 text-right w-20">2025 BDT</th>
                                        <th class="border border-slate-300 px-1 py-0.5 w-4 no-print">×</th>
                                    </tr>
                                </thead>
                                <tbody>${renderTwoColRows('8_14', 'n11', n11.rows)}</tbody>
                            </table>
                            ${addRowBtn('8_14', 'n11')}
                        </div>` : ''}
                    </div>
                </div>

                ${n12.visible ? `
                <div>
                    ${noteHeaderBar('8_14', 'n12', n12.title)}
                    <table class="w-full text-[9.5px] border-collapse border border-slate-300 mt-0.5">
                        <thead class="bg-slate-100">
                            <tr>
                                <th class="border border-slate-300 px-1.5 py-0.5 text-left">Particulars</th>
                                ${computedPartners.map(p => `<th class="border border-slate-300 px-1.5 py-0.5 text-right">${escapeHtml(p.name)} (BDT)</th>`).join('')}
                                <th class="border border-slate-300 px-1.5 py-0.5 text-right">Total (BDT)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="border border-slate-300 px-1.5 py-0.5">Opening capital balance</td>
                                ${computedPartners.map(p => `<td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(p.open)}</td>`).join('')}
                                <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(totOpen)}</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1.5 py-0.5">Additional capital introduced</td>
                                ${computedPartners.map(p => `<td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(p.add)}</td>`).join('')}
                                <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(totAdd)}</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1.5 py-0.5">Profit allocated for the year</td>
                                ${computedPartners.map(p => `<td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(p.prof)}</td>`).join('')}
                                <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(totProf)}</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1.5 py-0.5">Less: Partners' drawings</td>
                                ${computedPartners.map(p => `<td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(p.draw, true)}</td>`).join('')}
                                <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(totDraw, true)}</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1.5 py-0.5">Other equity adjustments</td>
                                ${computedPartners.map(() => `<td class="border border-slate-300 px-1.5 py-0.5 text-right">—</td>`).join('')}
                                <td class="border border-slate-300 px-1.5 py-0.5 text-right">—</td>
                            </tr>
                            <tr class="font-bold bg-slate-100">
                                <td class="border border-slate-300 px-1.5 py-0.5">Closing capital balance</td>
                                ${computedPartners.map(p => `<td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(p.close)}</td>`).join('')}
                                <td class="border border-slate-300 px-1.5 py-0.5 text-right font-mono">${fmt(totClose)}</td>
                            </tr>
                        </tbody>
                    </table>
                    <p class="text-[9px] italic text-slate-600 mt-0.5">${escapeHtml(n12.footer)}</p>
                </div>` : ''}

                ${n13.visible ? `
                <div>
                    ${noteHeaderBar('8_14', 'n13', n13.title)}
                    <p class="text-slate-700" contenteditable="true" onblur="updateStructuredNoteText('8_14','n13','text',this.textContent)">${escapeHtml(n13.text)}</p>
                </div>` : ''}

                ${n14.visible ? `
                <div>
                    ${noteHeaderBar('8_14', 'n14', n14.title)}
                    <p class="text-slate-700" contenteditable="true" onblur="updateStructuredNoteText('8_14','n14','text',this.textContent)">${escapeHtml(n14Body)}</p>
                </div>` : ''}
            `;
        }

        const s15 = structuredNotes15to25;
        const notes1525El = document.getElementById('notes-15-25-container');
        if (notes1525El) {
            notes1525El.innerHTML = `
                <div class="grid grid-cols-2 gap-2">
                    ${s15.n15.visible ? `
                    <div>
                        ${noteHeaderBar('15_25', 'n15', s15.n15.title)}
                        <table class="w-full text-[9px] border-collapse border border-slate-300 mt-0.5">
                            <thead class="bg-slate-100">
                                <tr>
                                    <th class="border border-slate-300 px-1 py-0.5 text-left">Particulars</th>
                                    <th class="border border-slate-300 px-1 py-0.5 text-right w-16">2026 BDT</th>
                                    <th class="border border-slate-300 px-1 py-0.5 text-right w-16">2025 BDT</th>
                                    <th class="border border-slate-300 px-0.5 py-0.5 w-4 no-print">×</th>
                                </tr>
                            </thead>
                            <tbody>${renderTwoColRows('15_25', 'n15', s15.n15.rows)}</tbody>
                        </table>
                        ${addRowBtn('15_25', 'n15')}
                        <p class="text-[8.5px] text-slate-600 mt-0.5">${escapeHtml(s15.n15.footer)}</p>
                    </div>` : ''}

                    <div class="space-y-1.5">
                        ${s15.n16.visible ? `
                        <div>
                            ${noteHeaderBar('15_25', 'n16', s15.n16.title)}
                            <table class="w-full text-[9px] border-collapse border border-slate-300 mt-0.5">
                                <thead class="bg-slate-100">
                                    <tr>
                                        <th class="border border-slate-300 px-1 py-0.5 text-left">Particulars</th>
                                        <th class="border border-slate-300 px-1 py-0.5 text-right w-16">2026 BDT</th>
                                        <th class="border border-slate-300 px-1 py-0.5 text-right w-16">2025 BDT</th>
                                        <th class="border border-slate-300 px-0.5 py-0.5 w-4 no-print">×</th>
                                    </tr>
                                </thead>
                                <tbody>${renderTwoColRows('15_25', 'n16', s15.n16.rows)}</tbody>
                            </table>
                            ${addRowBtn('15_25', 'n16')}
                        </div>` : ''}

                        ${s15.n18.visible ? `
                        <div>
                            ${noteHeaderBar('15_25', 'n18', s15.n18.title)}
                            <table class="w-full text-[9px] border-collapse border border-slate-300 mt-0.5">
                                <thead class="bg-slate-100">
                                    <tr>
                                        <th class="border border-slate-300 px-1 py-0.5 text-left">Particulars</th>
                                        <th class="border border-slate-300 px-1 py-0.5 text-right w-16">2026 BDT</th>
                                        <th class="border border-slate-300 px-1 py-0.5 text-right w-16">2025 BDT</th>
                                        <th class="border border-slate-300 px-0.5 py-0.5 w-4 no-print">×</th>
                                    </tr>
                                </thead>
                                <tbody>${renderTwoColRows('15_25', 'n18', s15.n18.rows)}</tbody>
                            </table>
                            ${addRowBtn('15_25', 'n18')}
                        </div>` : ''}
                    </div>
                </div>

                ${s15.n17.visible ? `
                <div>
                    ${noteHeaderBar('15_25', 'n17', s15.n17.title)}
                    <p class="text-slate-700" contenteditable="true" onblur="updateStructuredNoteText('15_25','n17','text',this.textContent)">${escapeHtml(s15.n17.text)}</p>
                </div>` : ''}

                ${s15.n19.visible ? `
                <div class="border border-slate-300 rounded p-1.5 bg-slate-50/60">
                    ${noteHeaderBar('15_25', 'n19', s15.n19.title)}
                    <p class="text-[8.5px] text-slate-700 mb-1">
                        S.R.O. No. 324-Ain/Aykor-05/2026 (31 Aug 2026, NBR u/s 76(1) & 163(6) of Income Tax Act 2023) slab-based turnover tax: up to Tk. 20M (<strong>0%</strong>); &gt; Tk. 20M to Tk. 40M (<strong>0.50%</strong>); &gt; Tk. 40M (<strong>1%</strong>). Corporate income tax is the higher of Minimum Tax (A) and Regular Tax @ ${regTaxRatePct.toFixed(1)}% on Taxable Profit (B):
                    </p>
                    <table class="w-full text-[8.5px] border-collapse border border-slate-300 bg-white">
                        <thead class="bg-slate-800 text-white">
                            <tr>
                                <th class="border border-slate-300 px-1 py-0.5 text-center w-6">SL</th>
                                <th class="border border-slate-300 px-1 py-0.5 text-left">Corporate Tax Assessment Item</th>
                                <th class="border border-slate-300 px-1 py-0.5 text-left w-28">Statutory Reference / Rate</th>
                                <th class="border border-slate-300 px-1 py-0.5 text-left w-32">Calculation Base</th>
                                <th class="border border-slate-300 px-1 py-0.5 text-right w-20">Amount (BDT)</th>
                                <th class="border border-slate-300 px-0.5 py-0.5 w-4 no-print">×</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="border border-slate-300 px-1 py-0.5 text-center">1</td>
                                <td class="border border-slate-300 px-1 py-0.5">Gross Receipts / Turnover from Contracts with Customers</td>
                                <td class="border border-slate-300 px-1 py-0.5">S.R.O. 324/2026</td>
                                <td class="border border-slate-300 px-1 py-0.5">${escapeHtml(taxInvoiceRef)}</td>
                                <td class="border border-slate-300 px-1 py-0.5 text-right font-mono">${fmtDec(grossReceipts)}</td>
                                <td class="border border-slate-300 px-0.5 py-0.5 text-center text-slate-400 no-print">—</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1 py-0.5 text-center">2</td>
                                <td class="border border-slate-300 px-1 py-0.5">Minimum Income Tax @ ${slabRatePct.toFixed(2)}% of Gross Receipts (A)</td>
                                <td class="border border-slate-300 px-1 py-0.5">Section 163(6) / S.R.O. 324</td>
                                <td class="border border-slate-300 px-1 py-0.5">${fmtDec(grossReceipts)} @ ${slabRatePct.toFixed(2)}%</td>
                                <td class="border border-slate-300 px-1 py-0.5 text-right font-mono">${fmtDec(minTaxA)}</td>
                                <td class="border border-slate-300 px-0.5 py-0.5 text-center text-slate-400 no-print">—</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1 py-0.5 text-center">3</td>
                                <td class="border border-slate-300 px-1 py-0.5">Taxable Business Profit for the Period</td>
                                <td class="border border-slate-300 px-1 py-0.5">Income Statement</td>
                                <td class="border border-slate-300 px-1 py-0.5">${escapeHtml(taxPbtNote)}</td>
                                <td class="border border-slate-300 px-1 py-0.5 text-right font-mono">${fmtDec(taxableProfit)}</td>
                                <td class="border border-slate-300 px-0.5 py-0.5 text-center text-slate-400 no-print">—</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1 py-0.5 text-center">4</td>
                                <td class="border border-slate-300 px-1 py-0.5">Tax Applied on Taxable Profit @ ${regTaxRatePct.toFixed(1)}% (B)</td>
                                <td class="border border-slate-300 px-1 py-0.5">Regular Tax Rate @ ${regTaxRatePct.toFixed(2)}%</td>
                                <td class="border border-slate-300 px-1 py-0.5">${fmtDec(taxableProfit)} @ ${regTaxRatePct.toFixed(2)}%</td>
                                <td class="border border-slate-300 px-1 py-0.5 text-right font-mono">${fmtDec(regTaxB)}</td>
                                <td class="border border-slate-300 px-0.5 py-0.5 text-center text-slate-400 no-print">—</td>
                            </tr>
                            <tr class="font-bold bg-slate-100">
                                <td class="border border-slate-300 px-1 py-0.5 text-center">—</td>
                                <td class="border border-slate-300 px-1 py-0.5">APPLICABLE CORPORATE INCOME TAX (HIGHER OF A AND B)</td>
                                <td class="border border-slate-300 px-1 py-0.5">Higher of (A) & (B)</td>
                                <td class="border border-slate-300 px-1 py-0.5">Higher of (A) & (B)</td>
                                <td class="border border-slate-300 px-1 py-0.5 text-right font-mono">${fmtDec(applicableTax)}</td>
                                <td class="border border-slate-300 px-0.5 py-0.5 text-center text-slate-400 no-print">—</td>
                            </tr>
                            <tr>
                                <td class="border border-slate-300 px-1 py-0.5 text-center">5</td>
                                <td class="border border-slate-300 px-1 py-0.5">Less: Advance Tax Paid (AIT) / TDS Deducted at Source</td>
                                <td class="border border-slate-300 px-1 py-0.5">Section 163 / AIT Challan</td>
                                <td class="border border-slate-300 px-1 py-0.5">${escapeHtml(taxAitRef)}</td>
                                <td class="border border-slate-300 px-1 py-0.5 text-right font-mono">${fmtDec(aitPaid, true)}</td>
                                <td class="border border-slate-300 px-0.5 py-0.5 text-center text-slate-400 no-print">—</td>
                            </tr>
                            ${s15.n19.extraRows.map(r => `
                                <tr>
                                    <td class="border border-slate-300 px-1 py-0.5 text-center" contenteditable="true" onblur="updateStructuredNoteRow('15_25','n19',${r.id},'sl',this.textContent)">${escapeHtml(r.sl)}</td>
                                    <td class="border border-slate-300 px-1 py-0.5" contenteditable="true" onblur="updateStructuredNoteRow('15_25','n19',${r.id},'item',this.textContent)">${escapeHtml(r.item)}</td>
                                    <td class="border border-slate-300 px-1 py-0.5" contenteditable="true" onblur="updateStructuredNoteRow('15_25','n19',${r.id},'ref',this.textContent)">${escapeHtml(r.ref)}</td>
                                    <td class="border border-slate-300 px-1 py-0.5" contenteditable="true" onblur="updateStructuredNoteRow('15_25','n19',${r.id},'base',this.textContent)">${escapeHtml(r.base)}</td>
                                    <td class="border border-slate-300 px-1 py-0.5 text-right font-mono" contenteditable="true" onblur="updateStructuredNoteRow('15_25','n19',${r.id},'amt',this.textContent)">${fmtDec(r.amt)}</td>
                                    <td class="border border-slate-300 px-0.5 py-0.5 text-center no-print"><button type="button" onclick="removeStructuredNoteRow('15_25','n19',${r.id})" class="text-rose-500 hover:text-rose-700 font-bold">×</button></td>
                                </tr>
                            `).join('')}
                            <tr class="font-bold bg-slate-900 text-white">
                                <td class="border border-slate-400 px-1 py-0.5 text-center">—</td>
                                <td class="border border-slate-400 px-1 py-0.5">${escapeHtml(taxClosingLabel)}</td>
                                <td class="border border-slate-400 px-1 py-0.5 font-normal">Net Payable to NBR</td>
                                <td class="border border-slate-400 px-1 py-0.5 font-normal">Balance Due (TDS covers liability)</td>
                                <td class="border border-slate-400 px-1 py-0.5 text-right font-mono">${fmtDec(netCurrentTaxPayable)}</td>
                                <td class="border border-slate-400 px-0.5 py-0.5 text-center no-print">—</td>
                            </tr>
                        </tbody>
                    </table>
                    ${addRowBtn('15_25', 'n19')}
                    <p class="text-[8.5px] text-slate-700 mt-0.5"><strong>Note on Statutory Assessment:</strong> ${escapeHtml(taxStatutoryNote)}</p>
                </div>` : ''}

                ${['n20', 'n21', 'n22', 'n23', 'n24', 'n25'].map(k => s15[k].visible ? `
                    <div>
                        ${noteHeaderBar('15_25', k, s15[k].title)}
                        <p class="text-slate-700" contenteditable="true" onblur="updateStructuredNoteText('15_25','${k}','text',this.textContent)">${escapeHtml(s15[k].text)}</p>
                    </div>
                ` : '').join('')}

                ${disclosureNotes.map(n => `
                    <div>
                        <div class="flex items-center justify-between">
                            <h4 class="font-bold text-slate-900">${escapeHtml(n.title)}</h4>
                            <button type="button" onclick="removeDisclosureNote(${n.id})" class="no-print text-[9px] text-rose-600 hover:underline">Delete this note</button>
                        </div>
                        <p class="text-slate-700">${escapeHtml(n.text)}</p>
                    </div>
                `).join('')}
            `;
        }

        // Populate Management Report, Executive Summary & Ratios
        const revDiff = revCurr - revPrev;
        const revGrowthPct = revPrev > 0 ? ((revDiff / revPrev) * 100).toFixed(1) : '0.0';
        const gpm = revCurr > 0 ? ((gpCurr / revCurr) * 100).toFixed(1) : '0.0';
        const npm = revCurr > 0 ? ((npCurr / revCurr) * 100).toFixed(1) : '0.0';
        const currRatio = clCurr > 0 ? (caCurr / clCurr).toFixed(2) : '0.00';
        const roe = eqCurr > 0 ? ((npCurr / eqCurr) * 100).toFixed(1) : '0.0';

        setText('mr-rev', fmt(revCurr));
        setText('mr-rev-prev', fmt(revPrev));
        setText('mr-rev-diff', fmt(revDiff));
        setText('mr-rev-growth', revGrowthPct + '%');
        setText('mr-tbl-rev', fmt(revCurr));
        setText('mr-tbl-rev-prev', fmt(revPrev));
        setText('mr-tbl-cos', fmt(cosCurr));
        setText('mr-tbl-cos-prev', fmt(cosPrev));
        setText('mr-gp', fmt(gpCurr));
        setText('mr-gp-prev', fmt(gpPrev));
        setText('mr-gpm', gpm + '%');
        setText('mr-np', fmt(npCurr));
        setText('mr-np-prev', fmt(npPrev));
        setText('mr-npm', npm + '%');

        setText('mr-assets', fmt(totAssetsCurr));
        setText('mr-liab', fmt(totLiabCurr));
        setText('mr-eq', fmt(eqCurr));
        setText('mr-wc', fmt(caCurr - clCurr));
        setText('mr-add-cap', fmt(totAdd));
        setText('mr-draw', fmt(totDraw));
        setText('mr-np-alloc', fmt(tciCurr));
        setText('mr-curr-rev-base', fmt(revCurr));
        setText('mr-target-rev', fmt(Math.round(revCurr * 1.15)));

        const eqPct = totAssetsCurr > 0 ? ((eqCurr / totAssetsCurr) * 100).toFixed(1) : '0.0';
        const liabPct = totAssetsCurr > 0 ? ((totLiabCurr / totAssetsCurr) * 100).toFixed(1) : '0.0';

        setText('es-rev', fmt(revCurr));
        setText('es-gp', fmt(gpCurr));
        setText('es-gpm-txt', gpm + '%');
        setText('es-gpm-bdt', gpm);
        setText('es-np', fmt(npCurr));
        setText('es-npm-txt', npm + '%');
        setText('es-assets', fmt(totAssetsCurr));
        setText('es-liab', fmt(totLiabCurr));
        setText('es-eq', fmt(eqCurr));
        setText('es-eq-pct', eqPct + '%');
        setText('es-liab-pct', liabPct + '%');
        setText('es-cfo', fmt(netOps));
        setText('es-cash', fmt(cfCloseCash));

        setText('es-hl-rev', 'BDT ' + fmt(revCurr));
        setText('es-hl-gp', 'BDT ' + fmt(gpCurr));
        setText('es-hl-np', 'BDT ' + fmt(npCurr));
        setText('es-hl-cfo', 'BDT ' + fmt(netOps));
        setText('es-hl-assets', 'BDT ' + fmt(totAssetsCurr));
        setText('es-hl-liab', 'BDT ' + fmt(totLiabCurr));
        setText('es-hl-eq', 'BDT ' + fmt(eqCurr));
        setText('es-hl-cash', 'BDT ' + fmt(cfCloseCash));

        setText('ratio-gpm', gpm + '%');
        setText('formula-gpm', `BDT ${fmt(gpCurr)} ÷ BDT ${fmt(revCurr)} × 100 = ${gpm}%`);
        setText('ratio-npm', npm + '%');
        setText('formula-npm', `BDT ${fmt(npCurr)} ÷ BDT ${fmt(revCurr)} × 100 = ${npm}%`);
        setText('ratio-curr', currRatio + 'x');
        setText('formula-curr', `BDT ${fmt(caCurr)} ÷ BDT ${fmt(clCurr)} = ${currRatio}x`);
        setText('ratio-roe', roe + '%');
        setText('formula-roe', `BDT ${fmt(npCurr)} ÷ BDT ${fmt(eqCurr)} × 100 = ${roe}%`);

        // Auto-Populate & Calculate NBR Income Tax Return Form IT-116 (2023)
        const includeIt116El = document.getElementById('includeIt116');
        const it116AutoSyncEl = document.getElementById('it116AutoSync');
        const includeIt116 = includeIt116El ? includeIt116El.checked : true;
        const it116AutoSync = it116AutoSyncEl ? it116AutoSyncEl.checked : true;
        const it116Wrapper = document.getElementById('it116-return-wrapper');
        const tocIt116 = document.getElementById('toc-it116-section');
        if (it116Wrapper) it116Wrapper.style.display = includeIt116 ? 'block' : 'none';
        if (tocIt116) tocIt116.style.display = includeIt116 ? 'block' : 'none';

        if (includeIt116) {
            window.syncTinFromSidebar(strVal('it116Tin'));
            ['p1', 'p2', 's1', 's2', 's3', 'p6'].forEach(prefix => {
                const el = document.getElementById(`it116-${prefix}-name`);
                if (el) el.value = firmName;
            });

            document.getElementById('it116-p1-bin').value = strVal('it116Bin');
            document.getElementById('it116-p1-circle').value = strVal('it116Circle');
            document.getElementById('it116-p1-zone').value = strVal('it116Zone');
            document.getElementById('it116-p1-taxyear').value = strVal('it116TaxYear');

            const catVal = strVal('it116Category');
            ['firm', 'aop', 'fund', 'trust', 'huf'].forEach(c => {
                const cb = document.getElementById(`it116-cat-${c}`);
                if (cb) cb.checked = (c === catVal);
            });

            const resVal = strVal('it116Resident');
            const resYes = document.getElementById('it116-res-yes');
            const resNo = document.getElementById('it116-res-no');
            if (resYes) resYes.checked = (resVal === 'resident');
            if (resNo) resNo.checked = (resVal === 'nonresident');

            document.getElementById('it116-p1-addr1').value = address;
            document.getElementById('it116-p1-phone').value = strVal('it116Phone');
            document.getElementById('it116-p1-email').value = strVal('it116Email');
            document.getElementById('it116-p1-regno').value = regNo;
            document.getElementById('it116-p1-regdate').value = strVal('it116RegDate');

            const p1PartnersEl = document.getElementById('it116-p1-partners-lines');
            if (p1PartnersEl) {
                p1PartnersEl.innerHTML = computedPartners.map((p, idx) => `
                    <div class="flex items-center">
                        <span class="text-[11.5px] font-semibold mr-2">(${idx + 1})</span>
                        <input type="text" value="${escapeHtml(p.name)} — ${escapeHtml(p.role)} (Profit Share: ${p.ratio}% | Closing Capital: BDT ${fmt(p.close)})" class="it116-line-input !ml-0">
                    </div>
                `).join('');
            }

            document.getElementById('it116-p2-enddate').value = reportDate;
            document.getElementById('it116-s3-bizname').value = firmName;
            document.getElementById('it116-s3-biztype').value = businessNature;
            document.getElementById('it116-s3-addr').value = address;
            document.getElementById('it116-p6-docs').value = strVal('it116DocsList');
            document.getElementById('it116-p6-date').value = approvalDate;

            const leadPartner = computedPartners[0] || { name: 'Managing Partner', role: 'Partner' };
            document.getElementById('it116-p6-declname').value = leadPartner.name;
            setText('it116-p6-sig-name', leadPartner.name);
            setText('it116-p6-sig-role', `${leadPartner.role} — ${firmName}`);

            if (it116AutoSync) {
                const sellingExp = pnlExpenses[0]?.curr || 0;
                const adminAndOtherExp = totOpExpCurr - sellingExp;
                const netOpProfitBeforeOtherInc = gpCurr - adminAndOtherExp - sellingExp - finCurr;

                document.getElementById('it116-s3-inc1').value = fmt(revCurr);
                document.getElementById('it116-s3-inc2').value = '0';
                document.getElementById('it116-s3-inc3').value = fmt(revCurr);
                document.getElementById('it116-s3-inc4').value = fmt(cosCurr);
                document.getElementById('it116-s3-inc5').value = fmt(gpCurr);
                document.getElementById('it116-s3-inc6').value = fmt(adminAndOtherExp);
                document.getElementById('it116-s3-inc7').value = fmt(sellingExp);
                document.getElementById('it116-s3-inc8').value = fmt(finCurr);
                document.getElementById('it116-s3-inc9').value = fmt(netOpProfitBeforeOtherInc);
                document.getElementById('it116-s3-inc10').value = fmt(othIncCurr);
                document.getElementById('it116-s3-inc11').value = fmt(pbtCurr);

                const ppeItem = bsItems.nca.find(i => i.label.toLowerCase().includes('property'));
                const intangItem = bsItems.nca.find(i => i.label.toLowerCase().includes('intangible'));
                const fixedAst = ppeItem ? ppeItem.curr : ncaCurr;
                const intangAst = intangItem ? intangItem.curr : (ncaCurr - fixedAst);
                const otherNca = Math.max(0, ncaCurr - fixedAst - intangAst);

                const recvItem = bsItems.ca.find(i => i.label.toLowerCase().includes('receivable'));
                const recvAst = recvItem ? recvItem.curr : 0;

                const cashHandRow = structuredNotes8to14.n11.rows.find(r => r.label.toLowerCase().includes('hand'));
                const cashInHand = cashHandRow ? Math.min(cashCurr, cashHandRow.curr) : 0;
                const bankBal = Math.max(0, cashCurr - cashInHand);
                const otherCurrAst = Math.max(0, caCurr - recvAst - cashCurr);

                document.getElementById('it116-s3-ast1').value = fmt(fixedAst);
                document.getElementById('it116-s3-ast2').value = fmt(intangAst);
                document.getElementById('it116-s3-ast3').value = fmt(otherNca);
                document.getElementById('it116-s3-ast4').value = fmt(otherCurrAst);
                document.getElementById('it116-s3-ast5').value = fmt(recvAst);
                document.getElementById('it116-s3-ast6').value = fmt(cashInHand);
                document.getElementById('it116-s3-ast7').value = fmt(bankBal);
                document.getElementById('it116-s3-ast8').value = fmt(totAssetsCurr);

                document.getElementById('it116-s3-cap9').value = fmt(totOpen + totAdd);
                document.getElementById('it116-s3-cap10').value = fmt(totProf);
                document.getElementById('it116-s3-cap11').value = fmt(totDraw);
                document.getElementById('it116-s3-cap12').value = fmt(totClose);
                document.getElementById('it116-s3-liab13').value = fmt(nclCurr);
                document.getElementById('it116-s3-liab14').value = fmt(clCurr);
                document.getElementById('it116-s3-liab15').value = fmt(totEqLiabCurr);

                const rentInc = val('it116RentInc');
                const agriInc = val('it116AgriInc');
                const bizInc = pbtCurr;
                const capGain = val('it116CapGain');
                const finInc = val('it116FinInc');
                const othSrcInc = 0;
                const foreignInc = val('it116ForeignInc');
                const totalIncome = rentInc + agriInc + bizInc + capGain + finInc + othSrcInc + foreignInc;

                document.getElementById('it116-s1-net').value = fmt(rentInc);
                document.getElementById('it116-s2-net').value = fmt(agriInc);

                document.getElementById('it116-inc-1').value = fmt(rentInc);
                document.getElementById('it116-inc-2').value = fmt(agriInc);
                document.getElementById('it116-inc-3').value = fmt(bizInc);
                document.getElementById('it116-inc-4').value = fmt(capGain);
                document.getElementById('it116-inc-5').value = fmt(finInc);
                document.getElementById('it116-inc-6').value = fmt(othSrcInc);
                document.getElementById('it116-inc-7').value = fmt(foreignInc);
                document.getElementById('it116-inc-8').value = fmt(totalIncome);

                const regularTaxOnTotalInc = taxCurr;
                const minTaxOnTds = val('it116Tds');
                const turnoverTaxCalc = Math.round(revCurr * (slabRatePct / 100));
                const taxPayableHigher = Math.max(regularTaxOnTotalInc, minTaxOnTds, turnoverTaxCalc);
                const envSurcharge = val('it116EnvSurcharge');
                const penaltyAmt = val('it116Penalty');
                const totalTaxPayable = taxPayableHigher + envSurcharge + penaltyAmt;

                document.getElementById('it116-tax-9').value = fmt(regularTaxOnTotalInc);
                document.getElementById('it116-tax-10').value = fmt(minTaxOnTds);
                document.getElementById('it116-tax-11').value = fmt(turnoverTaxCalc);
                document.getElementById('it116-tax-12').value = fmt(taxPayableHigher);
                document.getElementById('it116-tax-13').value = fmt(envSurcharge);
                document.getElementById('it116-tax-14').value = fmt(penaltyAmt);
                document.getElementById('it116-tax-15').value = fmt(totalTaxPayable);

                const tdsPaid = val('it116Tds');
                const aitPaidReturn = val('it116Ait');
                const balanceTaxWithReturn = Math.max(0, totalTaxPayable - tdsPaid - aitPaidReturn);
                const totalTaxPaid = tdsPaid + aitPaidReturn + balanceTaxWithReturn;
                const excessPayment = Math.max(0, totalTaxPaid - totalTaxPayable);

                document.getElementById('it116-pay-16').value = fmt(tdsPaid);
                document.getElementById('it116-pay-17').value = fmt(aitPaidReturn);
                document.getElementById('it116-pay-18').value = excessPayment > 0 ? strVal('it116TaxYear') : 'N/A';
                document.getElementById('it116-pay-19').value = fmt(balanceTaxWithReturn);
                document.getElementById('it116-pay-20').value = fmt(totalTaxPaid);
                document.getElementById('it116-pay-21').value = fmt(excessPayment);
                document.getElementById('it116-pay-22').value = fmt(val('it116ExemptInc'));

                // Sync IT-11 Cola (2023) Acknowledgement Receipt
                document.getElementById('it11cola-taxyear').value = strVal('it116TaxYear');
                document.getElementById('it11cola-name').value = firmName;
                document.getElementById('it11cola-nid').value = `Reg. No. ${regNo}`;
                document.getElementById('it11cola-circle').value = strVal('it116Circle');
                document.getElementById('it11cola-zone').value = strVal('it116Zone');
                document.getElementById('it11cola-income').value = fmt(totalIncome);
                document.getElementById('it11cola-taxpaid').value = fmt(totalTaxPaid);
            }
        }
    };

    initIt116TinBoxes();
    renderPartnersEditor();
    renderPnlExpensesEditor();
    renderCfEditors();
    renderBsEditors();
    renderNotesEditor();
    window.updateReport();
}
