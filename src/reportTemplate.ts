export const REPORT_HTML_TEMPLATE = `
    <!-- Top Navigation Bar -->
    <header id="top-navbar" class="no-print bg-slate-900 text-white px-4 py-3 flex flex-wrap items-center justify-between shadow-md z-20 shrink-0">
        <div class="flex items-center space-x-3">
            <div class="bg-emerald-600 p-2 rounded-lg">
                <svg class="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                </svg>
            </div>
            <div>
                <h1 class="font-bold text-sm sm:text-base tracking-tight">IFRS / IAS Partnership Report Generator</h1>
                <p class="text-xs text-slate-400 hidden sm:block">Compliant with IAS 1, IAS 7, IAS 8 & Bangladesh FRC Reporting Standards</p>
            </div>
        </div>

        <div class="flex flex-wrap items-center gap-2 mt-2 sm:mt-0">
            <!-- Mobile View Toggle -->
            <button id="toggle-view-btn" type="button" onclick="toggleMobileView()" class="lg:hidden bg-slate-700 hover:bg-slate-600 text-xs font-medium px-3 py-2 rounded-lg transition cursor-pointer">
                Switch to Preview
            </button>

            <button type="button" onclick="resetDemoData()" class="bg-slate-800 hover:bg-slate-700 border border-slate-700 text-slate-200 text-xs font-medium px-3 py-2 rounded-lg transition cursor-pointer">
                Reset Demo Data
            </button>

            <!-- Export / Save Options Group: PDF, Word (.docx), Excel (.xlsx) -->
            <div class="inline-flex flex-wrap items-center gap-1.5 rounded-lg shadow-sm bg-slate-800 p-1 border border-slate-700">
                <button id="print-pdf-btn" type="button" onclick="printReportToPdf()" title="Print or Save A4 PDF via Browser Native Print API" class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition cursor-pointer">
                    <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
                    </svg>
                    <span>Print / Save PDF (A4)</span>
                </button>

                <button id="export-word-btn" type="button" onclick="exportReportToWord()" title="Download complete A4 Report & IT-116 Return as Microsoft Word (.docx)" class="bg-blue-700 hover:bg-blue-600 text-white text-xs font-semibold px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition cursor-pointer">
                    <svg class="w-3.5 h-3.5 text-blue-200" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                    </svg>
                    <span>Download as Word</span>
                </button>

                <button id="export-excel-btn" type="button" onclick="exportReportToExcel()" title="Download all Financial Statements, Notes, Tax Schedule & IT-116 Return as Microsoft Excel (.xlsx)" class="bg-teal-700 hover:bg-teal-600 text-white text-xs font-semibold px-3 py-1.5 rounded-md flex items-center space-x-1.5 transition cursor-pointer">
                    <svg class="w-3.5 h-3.5 text-teal-200" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 0h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"></path>
                    </svg>
                    <span>Download as Excel</span>
                </button>
            </div>
        </div>
    </header>

    <!-- Pre-Print / Export Visual Notification Toast -->
    <div id="print-toast-notification" role="status" aria-live="polite" class="no-print hidden fixed top-16 right-4 z-50 max-w-md bg-slate-900 text-white border border-emerald-500/60 rounded-xl shadow-2xl px-4 py-3 flex items-start space-x-3 transition-all duration-200">
        <div class="bg-emerald-600/20 border border-emerald-500/40 text-emerald-400 p-1.5 rounded-lg shrink-0 mt-0.5">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
            </svg>
        </div>
        <div class="text-xs leading-relaxed">
            <p id="print-toast-title" class="font-semibold text-emerald-400 mb-0.5">Preparing A4 PDF Print Dialog…</p>
            <p id="print-toast-body" class="text-slate-200">Ensure printer settings are set to A4 portrait and 'Background graphics' are enabled for best results.</p>
        </div>
    </div>

    <!-- Main Workspace -->
    <div id="main-workspace" class="flex-1 flex overflow-hidden">

        <!-- Left Sidebar: Interactive Form Inputs -->
        <aside id="editor-sidebar" class="no-print w-full lg:w-[440px] xl:w-[480px] bg-white border-r border-slate-200 flex flex-col h-full shrink-0 z-10">
            <!-- Tabs -->
            <div class="flex border-b border-slate-200 bg-slate-50 text-xs font-semibold overflow-x-auto shrink-0">
                <button onclick="switchTab('firm')" id="tab-firm" class="tab-btn flex-1 py-3 px-2 border-b-2 border-emerald-600 text-emerald-700 bg-white whitespace-nowrap">1. Firm & Partners</button>
                <button onclick="switchTab('pnl')" id="tab-pnl" class="tab-btn flex-1 py-3 px-2 border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">2. Income & Cash</button>
                <button onclick="switchTab('bs')" id="tab-bs" class="tab-btn flex-1 py-3 px-2 border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">3. Balance Sheet</button>
                <button onclick="switchTab('mgmt')" id="tab-mgmt" class="tab-btn flex-1 py-3 px-2 border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">4. Narrative & Audit</button>
                <button onclick="switchTab('tax')" id="tab-tax" class="tab-btn flex-1 py-3 px-2 border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">5. Tax (IAS 12)</button>
                <button onclick="switchTab('it116')" id="tab-it116" class="tab-btn flex-1 py-3 px-2 border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">6. IT-116 Return</button>
            </div>

            <!-- Scrollable Form Container -->
            <div class="flex-1 overflow-y-auto p-5 space-y-6" id="form-container">
                
                <!-- TAB 1: FIRM & PARTNERS -->
                <div id="panel-firm" class="space-y-5">
                    <div class="bg-emerald-50 border border-emerald-200 rounded-lg p-3 text-xs text-emerald-800">
                        <strong>IFRS & BFRS Note:</strong> Under IAS 1 (applicable for periods starting before 1 Jan 2027), comparative figures and explicit compliance statements are mandatory.
                    </div>

                    <div class="space-y-3">
                        <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Entity Details</h3>
                        <div>
                            <label class="block text-xs font-medium text-slate-700 mb-1">Partnership Firm Name</label>
                            <input type="text" id="firmName" value="Padma & Meghna Associates" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none">
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-xs font-medium text-slate-700 mb-1">Financial Year Ended</label>
                                <input type="text" id="reportDate" value="31 December 2026" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg">
                            </div>
                            <div>
                                <label class="block text-xs font-medium text-slate-700 mb-1">Prior Year Ended</label>
                                <input type="text" id="prevDate" value="31 December 2025" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg">
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-xs font-medium text-slate-700 mb-1">Registration No. (RJSC)</label>
                                <input type="text" id="regNo" value="P-48291/2021" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg">
                            </div>
                            <div>
                                <label class="block text-xs font-medium text-slate-700 mb-1">Date of Establishment</label>
                                <input type="text" id="estDate" value="14 March 2021" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg">
                            </div>
                        </div>
                        <div>
                            <label class="block text-xs font-medium text-slate-700 mb-1">Registered Office Address</label>
                            <input type="text" id="address" value="Level 7, Gulshan Avenue, Dhaka-1212, Bangladesh" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg">
                        </div>
                        <div>
                            <label class="block text-xs font-medium text-slate-700 mb-1">Nature of Business & Principal Activities</label>
                            <textarea id="businessNature" rows="2" oninput="updateReport()" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg">Supply chain consulting, agro-commodity trading, and commercial distribution services across Bangladesh.</textarea>
                        </div>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Partners Configuration -->
                    <div class="space-y-4">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Partners' Capital & Profit Ratio</h3>
                            <div class="flex items-center space-x-2">
                                <span id="ratio-badge" class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">Total Ratio: 100%</span>
                                <button type="button" onclick="addPartner()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold px-2.5 py-1 rounded-md shadow-sm transition">+ Add Partner</button>
                            </div>
                        </div>

                        <!-- Dynamic Partners List -->
                        <div id="partners-editor-list" class="space-y-3"></div>
                    </div>
                </div>

                <!-- TAB 2: INCOME STATEMENT & CASH FLOWS -->
                <div id="panel-pnl" class="hidden space-y-5">
                    <div class="space-y-3">
                        <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Trading & Gross Profit (BDT)</h3>
                        
                        <div class="grid grid-cols-12 gap-1.5 text-[11px] font-semibold text-slate-500 px-1">
                            <span class="col-span-5">Account Head</span>
                            <span class="col-span-2 text-center">Note</span>
                            <span class="col-span-2 text-right">2026</span>
                            <span class="col-span-3 text-right">2025</span>
                        </div>

                        <div class="grid grid-cols-12 gap-1.5 items-center">
                            <input type="text" id="revLabel" value="Revenue" oninput="updateReport()" class="col-span-5 px-2 py-1.5 text-xs font-medium text-slate-700 border border-slate-300 rounded">
                            <input type="text" id="revNote" value="Note 6" oninput="updateReport()" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded">
                            <input type="number" id="revCurr" value="18500000" oninput="updateReport()" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded">
                            <input type="number" id="revPrev" value="15200000" oninput="updateReport()" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50">
                        </div>

                        <div class="grid grid-cols-12 gap-1.5 items-center">
                            <input type="text" id="cosLabel" value="Cost of sales" oninput="updateReport()" class="col-span-5 px-2 py-1.5 text-xs font-medium text-slate-700 border border-slate-300 rounded">
                            <input type="text" id="cosNote" value="Note 7" oninput="updateReport()" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded">
                            <input type="number" id="cosCurr" value="11800000" oninput="updateReport()" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded">
                            <input type="number" id="cosPrev" value="9850000" oninput="updateReport()" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50">
                        </div>

                        <div class="grid grid-cols-12 gap-1.5 items-center">
                            <input type="text" id="othIncLabel" value="Other income" oninput="updateReport()" class="col-span-5 px-2 py-1.5 text-xs font-medium text-slate-700 border border-slate-300 rounded">
                            <input type="text" id="othIncNote" value="Note 4" oninput="updateReport()" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded">
                            <input type="number" id="othIncCurr" value="240000" oninput="updateReport()" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded">
                            <input type="number" id="othIncPrev" value="180000" oninput="updateReport()" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50">
                        </div>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Dynamic Operating Expenses -->
                    <div class="space-y-2.5">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Operating Expenses (Add / Edit / Delete)</h3>
                            <button type="button" onclick="addPnlExpense()" class="bg-slate-800 hover:bg-slate-700 text-white text-[11px] font-medium px-2 py-1 rounded transition">+ Add Expense Head</button>
                        </div>
                        <div id="pnl-expenses-list" class="space-y-2"></div>
                    </div>

                    <hr class="border-slate-200">

                    <div class="space-y-2.5">
                        <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Finance, Tax & OCI</h3>
                        <div class="grid grid-cols-12 gap-1.5 items-center">
                            <input type="text" id="finLabel" value="Finance costs" oninput="updateReport()" class="col-span-5 px-2 py-1.5 text-xs font-medium text-slate-700 border border-slate-300 rounded">
                            <input type="text" id="finNote" value="Note 18" oninput="updateReport()" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded">
                            <input type="number" id="finCostCurr" value="320000" oninput="updateReport()" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded">
                            <input type="number" id="finCostPrev" value="350000" oninput="updateReport()" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50">
                        </div>

                        <div class="grid grid-cols-12 gap-1.5 items-center">
                            <input type="text" id="taxLabel" value="Income tax expense" oninput="updateReport()" class="col-span-5 px-2 py-1.5 text-xs font-medium text-slate-700 border border-slate-300 rounded">
                            <input type="text" id="taxNote" value="Note 19" oninput="updateReport()" class="col-span-2 px-1 py-1.5 text-[11px] text-center text-slate-500 border border-slate-300 rounded">
                            <input type="number" id="taxCurr" value="580000" oninput="updateReport()" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded">
                            <input type="number" id="taxPrev" value="430000" oninput="updateReport()" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50">
                        </div>
                        <p id="tax-sync-hint" class="text-[10px] text-emerald-700 hidden">✓ Current year tax is auto-synced from Tab 5 (IAS 12 & S.R.O. 324/2026).</p>

                        <div class="grid grid-cols-12 gap-1.5 items-center">
                            <input type="text" id="ociLabel" value="Other comprehensive income" oninput="updateReport()" class="col-span-5 px-2 py-1.5 text-xs font-medium text-slate-700 border border-slate-300 rounded">
                            <span class="col-span-2 text-center text-[11px] text-slate-400">—</span>
                            <input type="number" id="ociCurr" value="0" oninput="updateReport()" class="col-span-2 px-1.5 py-1.5 text-xs text-right border border-slate-300 rounded">
                            <input type="number" id="ociPrev" value="0" oninput="updateReport()" class="col-span-3 px-2 py-1.5 text-xs text-right border border-slate-300 rounded bg-slate-50">
                        </div>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Dynamic Cash Flow Inputs (IAS 7) -->
                    <div class="space-y-4">
                        <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Cash Flow Statement (IAS 7) - Current Year</h3>
                        
                        <div class="grid grid-cols-2 gap-3 bg-slate-50 p-3 rounded-lg border border-slate-200">
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">Cash Generated from Ops</label>
                                <input type="number" id="cfOpsGen" value="2650000" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">Opening Cash (1 Jan)</label>
                                <input type="number" id="cfOpenCash" value="1450000" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white">
                            </div>
                        </div>

                        <!-- Investing Cash Flow Dynamic List -->
                        <div class="space-y-2">
                            <div class="flex items-center justify-between">
                                <span class="text-xs font-semibold text-slate-700">Investing Activities (Use - for Outflow)</span>
                                <button type="button" onclick="addCfItem('investing')" class="bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-medium px-2 py-1 rounded">+ Add Investing Line</button>
                            </div>
                            <div id="cf-investing-list" class="space-y-1.5"></div>
                        </div>

                        <!-- Financing Cash Flow Dynamic List -->
                        <div class="space-y-2">
                            <div class="flex items-center justify-between">
                                <span class="text-xs font-semibold text-slate-700">Other Financing Lines (Use - for Outflow)</span>
                                <button type="button" onclick="addCfItem('financing')" class="bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-medium px-2 py-1 rounded">+ Add Financing Line</button>
                            </div>
                            <div id="cf-financing-list" class="space-y-1.5"></div>
                        </div>

                        <p class="text-[11px] text-slate-500 italic">Note: Partner capital introduced, partner drawings, interest paid, and taxes paid automatically sync from Partners & Income Statement.</p>
                    </div>
                </div>

                <!-- TAB 3: STATEMENT OF FINANCIAL POSITION (BALANCE SHEET) -->
                <div id="panel-bs" class="hidden space-y-5">
                    <div id="bs-balance-alert" class="p-3 rounded-lg text-xs font-medium border">
                        Checking Balance Sheet equation...
                    </div>

                    <!-- Non-Current Assets -->
                    <div class="space-y-2">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Non-Current Assets</h3>
                            <button type="button" onclick="addBsItem('nca')" class="bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-medium px-2 py-1 rounded">+ Add Head</button>
                        </div>
                        <div id="bs-nca-list" class="space-y-1.5"></div>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Current Assets -->
                    <div class="space-y-2">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Current Assets</h3>
                            <button type="button" onclick="addBsItem('ca')" class="bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-medium px-2 py-1 rounded">+ Add Head</button>
                        </div>
                        <p class="text-[10px] text-slate-500">Tip: Keep "Cash and cash equivalents" marked as Cash so Cash Flow reconciliation tracks accurately.</p>
                        <div id="bs-ca-list" class="space-y-1.5"></div>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Non-Current Liabilities -->
                    <div class="space-y-2">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Non-Current Liabilities</h3>
                            <button type="button" onclick="addBsItem('ncl')" class="bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-medium px-2 py-1 rounded">+ Add Head</button>
                        </div>
                        <div id="bs-ncl-list" class="space-y-1.5"></div>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Current Liabilities -->
                    <div class="space-y-2">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Current Liabilities</h3>
                            <button type="button" onclick="addBsItem('cl')" class="bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-medium px-2 py-1 rounded">+ Add Head</button>
                        </div>
                        <div id="bs-cl-list" class="space-y-1.5"></div>
                    </div>
                </div>

                <!-- TAB 4: MANAGEMENT NARRATIVE, NOTES & AUDITOR -->
                <div id="panel-mgmt" class="hidden space-y-4">
                    <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Management Report Sections (Editable)</h3>
                    
                    <div>
                        <label class="block text-xs font-medium text-slate-700 mb-1">1.3 Business Operations & Key Developments</label>
                        <textarea id="txtOps" rows="4" oninput="updateReport()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg">During FY 2026, the Firm continued its principal activities across Supply Chain Consulting, Agro-Commodity Trading, and Commercial Distribution Services. Key developments included expanding regional distribution across Chattogram and Sylhet, onboarding 18 new institutional clients, and upgrading cold-storage logistics capacity.</textarea>
                    </div>

                    <div>
                        <label class="block text-xs font-medium text-slate-700 mb-1">1.7 Risk Management & Business Challenges</label>
                        <textarea id="txtRisks" rows="4" oninput="updateReport()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg">Management actively monitors Foreign Exchange Risk on imported inputs, Credit & Receivables Risk (enforcing a standard 45-day maximum credit period), Supply Chain & Inflation Risk, Operational & Logistics Risk, and Business Continuity Risk.</textarea>
                    </div>

                    <div>
                        <label class="block text-xs font-medium text-slate-700 mb-1">1.8 Future Business Plans & Prospects (FY 2027)</label>
                        <textarea id="txtFuture" rows="4" oninput="updateReport()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg">For FY 2027, the partners plan to invest approximately BDT 2,500,000 in ERP automation and sustainable packaging lines while targeting a 15% revenue growth (indicative target of BDT 21,275,000) through institutional expansion and operational efficiency.</textarea>
                    </div>

                    <div>
                        <label class="block text-xs font-medium text-slate-700 mb-1">1.9 Management Responsibility Statement</label>
                        <textarea id="txtResp" rows="3" oninput="updateReport()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg">The management of Padma & Meghna Associates acknowledges its responsibility for maintaining complete accounting records, selecting consistent accounting policies, establishing internal financial controls, assessing going concern, and preparing financial statements in accordance with IFRS, IAS, and Bangladesh regulatory requirements.</textarea>
                    </div>

                    <hr class="border-slate-200">

                    <!-- Dynamic Notes to the Financial Statements -->
                    <div class="space-y-2.5">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Notes to Financial Statements (Add / Edit / Delete)</h3>
                            <button type="button" onclick="addDisclosureNote()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-semibold px-2.5 py-1 rounded">+ Add Note</button>
                        </div>
                        <div id="notes-editor-list" class="space-y-2"></div>
                    </div>

                    <hr class="border-slate-200">

                    <div class="space-y-3">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Independent Auditor's Report (ISA 700)</h3>
                            <label class="inline-flex items-center cursor-pointer">
                                <input type="checkbox" id="includeAudit" checked onchange="updateReport()" class="rounded text-emerald-600 focus:ring-emerald-500">
                                <span class="ml-2 text-xs font-medium text-slate-700">Include Report</span>
                            </label>
                        </div>

                        <div id="audit-fields" class="space-y-2.5">
                            <div>
                                <label class="block text-xs text-slate-600 mb-1">Audit Firm Name (Chartered Accountants)</label>
                                <input type="text" id="auditorFirm" value="Rahman, Chowdhury & Co., Chartered Accountants" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div class="grid grid-cols-2 gap-2">
                                <div>
                                    <label class="block text-xs text-slate-600 mb-1">Engagement Partner & ICAB Enrolment</label>
                                    <input type="text" id="auditorPartner" value="Mahmudul Hasan, FCA (Enrol: 1492)" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                                </div>
                                <div>
                                    <label class="block text-xs text-slate-600 mb-1">Approval / Audit Date</label>
                                    <input type="text" id="approvalDate" value="28 February 2027" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- TAB 5: TAXATION (IAS 12 & S.R.O. 324/2026) -->
                <div id="panel-tax" class="hidden space-y-4">
                    <div class="bg-emerald-50 border border-emerald-200 rounded-lg p-3 text-xs text-emerald-900 space-y-1">
                        <div class="flex items-center justify-between">
                            <strong>S.R.O. No. 324-Ain/Aykor-05/2026 (NBR)</strong>
                            <button type="button" onclick="loadBwdbTaxSample()" class="bg-emerald-700 hover:bg-emerald-800 text-white text-[10px] font-semibold px-2 py-1 rounded shadow-sm transition">Load BWDB Bill 1952 Case</button>
                        </div>
                        <p class="text-[11px] text-emerald-800">Issued 31 Aug 2026 u/s 76(1) & 163(6) of Income Tax Act 2023: 0% up to Tk. 20M; 0.50% (> Tk. 20M–40M); 1% (> Tk. 40M).</p>
                    </div>

                    <div class="flex items-center justify-between bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                        <label class="inline-flex items-center cursor-pointer">
                            <input type="checkbox" id="includeTaxSchedule" checked onchange="updateReport()" class="rounded text-emerald-600">
                            <span class="ml-2 text-xs font-semibold text-slate-700">Include IAS 12 Tax Page in Report</span>
                        </label>
                        <label class="inline-flex items-center cursor-pointer">
                            <input type="checkbox" id="syncTaxToPnl" onchange="updateReport()" class="rounded text-emerald-600">
                            <span class="ml-2 text-xs font-semibold text-emerald-700">Auto-Sync to P&L</span>
                        </label>
                    </div>

                    <div class="grid grid-cols-2 gap-2.5">
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Tax Assessment Year / Period</label>
                            <input type="text" id="taxPeriodTitle" value="FY 2025-2026 (As at 30 June 2026)" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Turnover Tax Slab (S.R.O. 324)</label>
                            <select id="taxTurnoverSlab" onchange="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded bg-white">
                                <option value="0">Up to Tk. 20M (0.00%)</option>
                                <option value="0.5" selected>Tk. 20M to Tk. 40M (0.50%)</option>
                                <option value="1.0">Exceeding Tk. 40M (1.00%)</option>
                            </select>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-2.5">
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Gross Receipts / Turnover (BDT)</label>
                            <input type="number" step="0.01" id="taxGrossReceipts" value="940434.00" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Contract / Invoice Reference</label>
                            <input type="text" id="taxInvoiceRef" value="Contract Invoices (BWDB Bill 1952)" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                    </div>

                    <div class="grid grid-cols-3 gap-2">
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Taxable Profit / PBT</label>
                            <input type="number" step="0.01" id="taxableProfit" value="139143.50" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Regular Rate (%)</label>
                            <input type="number" step="0.01" id="regularTaxRate" value="27.50" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">PBT Note Ref</label>
                            <input type="text" id="taxPbtNote" value="PBT (Note 11.00)" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-2.5">
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Advance Tax / AIT / TDS Paid (BDT)</label>
                            <input type="number" step="0.01" id="taxAitPaid" value="47022.00" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">AIT / Challan Reference</label>
                            <input type="text" id="taxAitRef" value="BWDB Bill 1952 Token 00324450" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-2.5">
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Deferred Tax Expense/(Income) (IAS 12)</label>
                            <input type="number" step="0.01" id="deferredTaxAmt" value="0.00" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">Net Tax Closing Date Label</label>
                            <input type="text" id="taxClosingLabel" value="TOTAL CURRENT TAX LIABILITY AS AT 30 JUNE 2026" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                        </div>
                    </div>

                    <div>
                        <label class="block text-[11px] font-medium text-slate-600 mb-1">Note on Statutory Assessment (Editable)</label>
                        <textarea id="taxStatutoryNote" rows="4" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">Regular tax @ 27.5% on taxable profit is BDT 38,264.46 (higher than 0.50% minimum tax of BDT 4,702.17). Advance Income Tax (AIT / TDS) of BDT 47,022.00 was deducted at source by BWDB under Section 89 of the Income Tax Act 2023. Under Section 163(2)(b), tax deducted at source serves as the minimum tax, leaving a Net Current Tax Liability of BDT 0.00 at June 30, 2026.</textarea>
                    </div>
                </div>

                <!-- TAB 6: NBR INCOME TAX RETURN IT-116 (2023) -->
                <div id="panel-it116" class="hidden space-y-4">
                    <div class="bg-blue-50 border border-blue-200 rounded-lg p-3 text-xs text-blue-900 space-y-1">
                        <div class="flex items-center justify-between">
                            <strong>NBR Form IT-116 (2023) & IT-11 Cola</strong>
                            <span class="px-2 py-0.5 rounded bg-blue-100 text-blue-800 text-[10px] font-semibold">Section E (8 Pages)</span>
                        </div>
                        <p class="text-[11px] text-blue-800">Automatically pulls Firm Name, Address, Registration, Partners, Schedule 3 (Business Income & Assets/Liabilities), Total Income, Tax Computation, TDS/AIT, and the IT-11 Cola (2023) Acknowledgement Receipt.</p>
                    </div>

                    <div class="flex items-center justify-between bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                        <label class="inline-flex items-center cursor-pointer">
                            <input type="checkbox" id="includeIt116" checked onchange="updateReport()" class="rounded text-emerald-600">
                            <span class="ml-2 text-xs font-semibold text-slate-700">Include IT-116 Return Pages in Report</span>
                        </label>
                        <label class="inline-flex items-center cursor-pointer">
                            <input type="checkbox" id="it116AutoSync" checked onchange="updateReport()" class="rounded text-emerald-600">
                            <span class="ml-2 text-xs font-semibold text-emerald-700">Auto-Calculate from Report</span>
                        </label>
                    </div>

                    <div class="space-y-2.5">
                        <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Taxpayer Registration & Jurisdiction</h3>
                        <div class="grid grid-cols-2 gap-2">
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">12-Digit Taxpayer TIN</label>
                                <input type="text" maxlength="12" id="it116Tin" value="482910472819" oninput="syncTinFromSidebar(this.value)" class="w-full px-2.5 py-1.5 text-xs font-mono tracking-widest border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">BIN (VAT Reg. No.)</label>
                                <input type="text" id="it116Bin" value="003849201-0102" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                        </div>
                        <div class="grid grid-cols-3 gap-2">
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">4(a) Tax Circle</label>
                                <input type="text" id="it116Circle" value="Circle-142 (Companies)" oninput="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">4(b) Tax Zone</label>
                                <input type="text" id="it116Zone" value="Tax Zone-07, Dhaka" oninput="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">5. Assessment Year</label>
                                <input type="text" id="it116TaxYear" value="2027-2028" oninput="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                        </div>
                        <div class="grid grid-cols-3 gap-2">
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">Category</label>
                                <select id="it116Category" onchange="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded bg-white">
                                    <option value="firm" selected>Firm</option>
                                    <option value="aop">Association of Persons</option>
                                    <option value="fund">Fund</option>
                                    <option value="trust">Trust</option>
                                    <option value="huf">Hindu Undivided Family</option>
                                </select>
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">Residential Status</label>
                                <select id="it116Resident" onchange="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded bg-white">
                                    <option value="resident" selected>Resident</option>
                                    <option value="nonresident">Non-resident</option>
                                </select>
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">Reg. Date (YYYY-MM-DD)</label>
                                <input type="date" id="it116RegDate" value="2021-03-14" oninput="updateReport()" class="w-full px-2 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-2">
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">Telephone</label>
                                <input type="text" id="it116Phone" value="+880 2-9884120" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">E-mail</label>
                                <input type="email" id="it116Email" value="accounts@padmameghna.com.bd" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                        </div>
                    </div>

                    <hr class="border-slate-200">

                    <div class="space-y-2.5">
                        <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Other Income Heads & Tax Payments (BDT)</h3>
                        <div class="grid grid-cols-2 gap-2">
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">1. Income from Rent (Sch. 1)</label>
                                <input type="number" id="it116RentInc" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">2. Income from Agriculture (Sch. 2)</label>
                                <input type="number" id="it116AgriInc" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">4. Capital Gain</label>
                                <input type="number" id="it116CapGain" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">5. Financial Assets Income</label>
                                <input type="number" id="it116FinInc" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">7. Foreign Sourced Income</label>
                                <input type="number" id="it116ForeignInc" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">13. Environment Surcharge</label>
                                <input type="number" id="it116EnvSurcharge" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">14. Delay Interest / Penalty</label>
                                <input type="number" id="it116Penalty" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">16. TDS / Tax Deducted at Source</label>
                                <input type="number" id="it116Tds" value="350000" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">17. Advance Tax Paid (AIT)</label>
                                <input type="number" id="it116Ait" value="230000" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                            <div>
                                <label class="block text-[11px] font-medium text-slate-600 mb-1">22. Tax Exempted Income</label>
                                <input type="number" id="it116ExemptInc" value="0" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">
                            </div>
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-600 mb-1">List of Attached Documents (Page 6 of Return)</label>
                            <textarea id="it116DocsList" rows="3" oninput="updateReport()" class="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded">1. Audited Statement of Financial Position, Statement of Profit or Loss, Partners' Equity & Cash Flows for FY ended 31 December 2026.&#10;2. Notes to the Financial Statements (Notes 1 to 25) & Fixed Asset Depreciation Schedule.&#10;3. Bank Statements, AIT Challans & TDS Certificates u/s 89 of the Income Tax Act, 2023.</textarea>
                        </div>
                    </div>
                </div>

            </div>
        </aside>

        <!-- Right Pane: Live Document Preview -->
        <main id="preview-container" class="flex-1 overflow-y-auto p-4 sm:p-8 bg-slate-300/70 hidden lg:block">
            
            <!-- ==================== PAGE 1: COVER PAGE ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="pt-10">
                    <div class="inline-block px-3.5 py-1.5 bg-slate-900 text-white text-xs font-semibold tracking-widest uppercase mb-8">
                        Annual Financial Report
                    </div>
                    <h1 class="report-serif text-4xl font-bold text-slate-900 tracking-tight mb-4" id="doc-cover-firm">
                        Padma & Meghna Associates
                    </h1>
                    <p class="text-xl text-slate-600 font-medium">
                        For the Financial Year Ended <span id="doc-cover-date" class="text-slate-900 font-semibold">31 December 2026</span>
                    </p>
                    <div class="w-28 h-1.5 bg-emerald-700 mt-8"></div>
                </div>

                <div class="my-12 bg-slate-50 border-l-4 border-emerald-700 p-7 space-y-2.5">
                    <p class="text-xs font-bold uppercase tracking-wider text-emerald-800">Applicable Reporting Framework</p>
                    <p class="text-sm text-slate-700 leading-relaxed">
                        Prepared in accordance with <strong>International Financial Reporting Standards (IFRS)</strong> and <strong>International Accounting Standards (IAS)</strong> as issued by the IASB and adopted in Bangladesh by the Financial Reporting Council (FRC).
                    </p>
                </div>

                <div class="text-xs text-slate-600 space-y-2 border-t border-slate-300 pt-6 pb-4">
                    <p><strong class="text-slate-800">Registration Number:</strong> <span id="doc-cover-reg">P-48291/2021</span></p>
                    <p><strong class="text-slate-800">Registered Office:</strong> <span id="doc-cover-addr">Level 7, Gulshan Avenue, Dhaka-1212, Bangladesh</span></p>
                    <p><strong class="text-slate-800">Reporting Currency:</strong> Bangladeshi Taka (BDT)</p>
                </div>
            </section>

            <!-- ==================== PAGE 2: TABLE OF CONTENTS ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div>
                    <div class="border-b-2 border-slate-900 pb-2.5 mb-8">
                        <p class="text-[11px] font-bold uppercase tracking-wider text-emerald-700 doc-firm-name">Padma & Meghna Associates</p>
                        <h2 class="report-serif text-2xl font-bold text-slate-900 uppercase tracking-wide">
                            Table of Contents
                        </h2>
                    </div>

                    <div class="space-y-6 text-sm">
                        <div>
                            <p class="text-xs font-bold uppercase tracking-wider text-emerald-800 mb-2.5">Section A — Management Reporting (Supplementary Information)</p>
                            <div class="space-y-2 pl-4 border-l-2 border-slate-300">
                                <div class="flex justify-between"><span>1. Management Report (Subsections 1.1 – 1.9)</span><span class="text-slate-500 font-medium">Section A.1</span></div>
                                <div class="flex justify-between"><span>2. Executive Summary of Financial Performance (Subsections 2.1 – 2.8)</span><span class="text-slate-500 font-medium">Section A.2</span></div>
                            </div>
                        </div>

                        <div id="toc-audit-section">
                            <p class="text-xs font-bold uppercase tracking-wider text-emerald-800 mb-2.5">Section B — Independent Assurance</p>
                            <div class="space-y-2 pl-4 border-l-2 border-slate-300">
                                <div class="flex justify-between"><span>3. Independent Auditor's Report (ISA 700)</span><span class="text-slate-500 font-medium">Section B</span></div>
                            </div>
                        </div>

                        <div>
                            <p class="text-xs font-bold uppercase tracking-wider text-emerald-800 mb-2.5">Section C — Financial Statements (IAS 1 Complete Set)</p>
                            <div class="space-y-2 pl-4 border-l-2 border-slate-300">
                                <div class="flex justify-between"><span>4. Statement of Financial Position (Balance Sheet)</span><span class="text-slate-500 font-medium">Section C.1</span></div>
                                <div class="flex justify-between"><span>5. Statement of Profit or Loss and Other Comprehensive Income</span><span class="text-slate-500 font-medium">Section C.2</span></div>
                                <div class="flex justify-between"><span>6. Statement of Changes in Partners' Equity</span><span class="text-slate-500 font-medium">Section C.3</span></div>
                                <div class="flex justify-between"><span>7. Statement of Cash Flows (IAS 7)</span><span class="text-slate-500 font-medium">Section C.4</span></div>
                                <div class="flex justify-between"><span>8. Notes to the Financial Statements (Notes 1 to 25)</span><span class="text-slate-500 font-medium">Section C.5</span></div>
                                <div id="toc-tax-row" class="flex justify-between"><span>9. Taxation Computation Schedule (IAS 12 & S.R.O. 324/2026)</span><span class="text-slate-500 font-medium">Section C.6</span></div>
                            </div>
                        </div>

                        <div>
                            <p class="text-xs font-bold uppercase tracking-wider text-emerald-800 mb-2.5">Section D — Closing and Authorization</p>
                            <div class="space-y-2 pl-4 border-l-2 border-slate-300">
                                <div class="flex justify-between"><span>10. Management's Conclusion</span><span class="text-slate-500 font-medium">Section D.1</span></div>
                                <div class="flex justify-between"><span>11. Approval, Authorization & Partners' Signatures</span><span class="text-slate-500 font-medium">Section D.2</span></div>
                            </div>
                        </div>

                        <div id="toc-it116-section">
                            <p class="text-xs font-bold uppercase tracking-wider text-emerald-800 mb-2.5">Section E — Statutory Income Tax Return (NBR Form IT-116 & IT-11 Cola, 2023)</p>
                            <div class="space-y-2 pl-4 border-l-2 border-slate-300">
                                <div class="flex justify-between"><span>12. Income Tax Return Form IT-116 (Part I, Total Income, Tax Computation & Schedules 1–3)</span><span class="text-slate-500 font-medium">Section E (Pages 1–7)</span></div>
                                <div class="flex justify-between"><span>13. Acknowledgement Receipt / Certificate of Return of Income (IT-11 Cola, 2023)</span><span class="text-slate-500 font-medium">Section E (Page 8)</span></div>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-3 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Annual Financial Report — FY Ended <span class="doc-report-date">31 December 2026</span></span>
                </div>
            </section>

            <!-- ==================== PAGE 3: SECTION A.1 MANAGEMENT REPORT ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-2.5 text-[10.5px] leading-snug">
                    <div class="border-b-2 border-slate-900 pb-1.5">
                        <p class="text-[10px] font-bold uppercase tracking-wider text-emerald-700">Section A — Management Reporting</p>
                        <h2 class="report-serif text-lg font-bold text-slate-900 uppercase">1. Management Report</h2>
                    </div>

                    <div>
                        <h3 class="font-bold text-slate-900 text-xs mb-0.5">1.1 Introduction to the Partnership Firm</h3>
                        <p class="text-slate-700">
                            The management of <strong class="doc-firm-name">Padma & Meghna Associates</strong> ("the Firm") presents its Annual Management Report for the financial year ended <strong class="doc-report-date">31 December 2026</strong>. Established on <span class="doc-est-date" id="doc-est-date">14 March 2021</span> (Registration No. <span class="doc-reg-no">P-48291/2021</span>), its principal office is located at <span class="doc-addr">Level 7, Gulshan Avenue, Dhaka-1212, Bangladesh</span>. Principal business activities: <span class="doc-biz-nature" id="doc-biz-nature">Supply chain consulting, agro-commodity trading, and commercial distribution services across Bangladesh.</span>
                        </p>
                    </div>

                    <div>
                        <h3 class="font-bold text-slate-900 text-xs mb-1">1.2 Partnership Structure, Management & Partner Responsibilities</h3>
                        <table class="w-full text-[10px] border-collapse border border-slate-300 mb-1.5" id="mr-structure-table">
                            <!-- Dynamically populated -->
                        </table>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-2" id="mr-partner-roles-grid">
                            <!-- Dynamically populated -->
                        </div>
                    </div>

                    <div>
                        <h3 class="font-bold text-slate-900 text-xs mb-0.5">1.3 Business Operations and Activities</h3>
                        <p class="text-slate-700" id="doc-txt-ops"></p>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-2 gap-2.5">
                        <div>
                            <h3 class="font-bold text-slate-900 text-xs mb-1">1.4 Financial Performance Overview</h3>
                            <table class="w-full text-[10px] border-collapse border border-slate-300">
                                <thead class="bg-slate-100">
                                    <tr>
                                        <th class="border border-slate-300 px-2 py-0.5 text-left">Particulars</th>
                                        <th class="border border-slate-300 px-2 py-0.5 text-right">Current (BDT)</th>
                                        <th class="border border-slate-300 px-2 py-0.5 text-right">Prior (BDT)</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr><td class="border border-slate-300 px-2 py-0.5">Total Revenue</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-tbl-rev">0</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-tbl-rev-prev">0</td></tr>
                                    <tr><td class="border border-slate-300 px-2 py-0.5">Cost of Sales</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-tbl-cos">0</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-tbl-cos-prev">0</td></tr>
                                    <tr class="font-bold bg-slate-50"><td class="border border-slate-300 px-2 py-0.5">Gross Profit (<span id="mr-gpm">0.0%</span>)</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-gp">0</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-gp-prev">0</td></tr>
                                    <tr class="font-bold bg-slate-50"><td class="border border-slate-300 px-2 py-0.5">Net Profit (<span id="mr-npm">0.0%</span>)</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono text-emerald-700" id="mr-np">0</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-np-prev">0</td></tr>
                                </tbody>
                            </table>
                            <p class="text-[9.5px] text-slate-600 mt-0.5">Revenue grew by <strong>BDT <span id="mr-rev-diff">0</span></strong> (<strong><span id="mr-rev-growth">0.0%</span></strong> YoY) from BDT <span id="mr-rev-prev">0</span> to BDT <span id="mr-rev">0</span>.</p>
                        </div>

                        <div>
                            <h3 class="font-bold text-slate-900 text-xs mb-1">1.5 Financial Position & Liquidity</h3>
                            <table class="w-full text-[10px] border-collapse border border-slate-300">
                                <thead class="bg-slate-100">
                                    <tr>
                                        <th class="border border-slate-300 px-2 py-0.5 text-left">Position Metric</th>
                                        <th class="border border-slate-300 px-2 py-0.5 text-right">Amount (BDT)</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr><td class="border border-slate-300 px-2 py-0.5">Total Assets</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono font-semibold" id="mr-assets">0</td></tr>
                                    <tr><td class="border border-slate-300 px-2 py-0.5">Total Liabilities</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-liab">0</td></tr>
                                    <tr class="font-bold bg-slate-50"><td class="border border-slate-300 px-2 py-0.5">Total Partners' Equity</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono" id="mr-eq">0</td></tr>
                                    <tr class="font-bold bg-slate-50"><td class="border border-slate-300 px-2 py-0.5">Net Working Capital (CA − CL)</td><td class="border border-slate-300 px-2 py-0.5 text-right font-mono text-emerald-700" id="mr-wc">0</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <div>
                        <h3 class="font-bold text-slate-900 text-xs mb-1">1.6 Capital and Profit Distribution</h3>
                        <p class="text-slate-700 mb-1">During the year, partners introduced additional capital of <strong>BDT <span id="mr-add-cap">0</span></strong>, withdrew drawings of <strong>BDT <span id="mr-draw">0</span></strong>, and allocated comprehensive income of <strong>BDT <span id="mr-np-alloc">0</span></strong>:</p>
                        <table class="w-full text-[10px] border-collapse border border-slate-300" id="mr-cap-dist-table">
                            <!-- Dynamically populated -->
                        </table>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-2 gap-2.5">
                        <div>
                            <h3 class="font-bold text-slate-900 text-xs mb-0.5">1.7 Risk Management & Business Challenges</h3>
                            <p class="text-slate-700" id="doc-txt-risks"></p>
                        </div>
                        <div>
                            <h3 class="font-bold text-slate-900 text-xs mb-0.5">1.8 Future Business Plans & Prospects</h3>
                            <p class="text-slate-700" id="doc-txt-future"></p>
                            <p class="text-[9.5px] text-emerald-800 font-medium mt-0.5">Indicative FY 2027 Revenue Target (+15% on BDT <span id="mr-curr-rev-base">0</span>): <strong>BDT <span id="mr-target-rev">0</span></strong></p>
                        </div>
                    </div>

                    <div>
                        <h3 class="font-bold text-slate-900 text-xs mb-0.5">1.9 Management Responsibility Statement</h3>
                        <p class="text-slate-700" id="doc-txt-resp"></p>
                    </div>
                </div>

                <div class="pt-2 border-t border-slate-300">
                    <div class="flex flex-wrap justify-between items-end gap-4">
                        <div class="flex gap-6" id="mr-signoff-grid"></div>
                        <div class="text-right text-[10px] text-slate-600">
                            <p><strong>Approved Date:</strong> <span class="doc-approval-date">28 February 2027</span></p>
                            <p><strong>Place:</strong> Dhaka, Bangladesh</p>
                        </div>
                    </div>
                </div>
            </section>

            <!-- ==================== PAGE 4: SECTION A.2 EXECUTIVE SUMMARY ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-3.5 text-[11.5px] leading-relaxed">
                    <div class="border-b-2 border-slate-900 pb-2">
                        <p class="text-[10px] font-bold uppercase tracking-wider text-emerald-700">Section A — Management Reporting</p>
                        <h2 class="report-serif text-xl font-bold text-slate-900 uppercase">2. Executive Summary of Financial Performance</h2>
                    </div>

                    <p class="text-slate-700">
                        The financial statements of <strong class="doc-firm-name">Padma & Meghna Associates</strong> have been prepared for the financial year ended <strong class="doc-report-date">31 December 2026</strong>, with reference to the applicable International Financial Reporting Standards (IFRS), International Accounting Standards (IAS), and relevant financial reporting and regulatory requirements in Bangladesh.
                    </p>

                    <!-- 2.1 & 2.2 -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                        <div class="bg-slate-50 border border-slate-200 rounded p-3 space-y-1">
                            <h3 class="font-bold text-slate-900 text-xs uppercase">2.1 Overview of Financial Performance</h3>
                            <p class="text-slate-700 text-[11px]">
                                During the financial year ended <span class="doc-report-date">31 December 2026</span>, the Firm generated total revenue of <strong>BDT <span id="es-rev">0</span></strong> and achieved a gross profit of <strong>BDT <span id="es-gp">0</span></strong> (gross margin of <strong><span id="es-gpm-txt">0.0%</span></strong>, retaining BDT <span id="es-gpm-bdt">0.00</span> per BDT 100 of revenue). After operating, administrative, finance, and tax charges, net profit reached <strong>BDT <span id="es-np">0</span></strong> (net margin of <strong><span id="es-npm-txt">0.0%</span></strong>).
                            </p>
                        </div>
                        <div class="bg-slate-50 border border-slate-200 rounded p-3 space-y-1">
                            <h3 class="font-bold text-slate-900 text-xs uppercase">2.2 Summary of Financial Position</h3>
                            <p class="text-slate-700 text-[11px]">
                                As at <span class="doc-report-date">31 December 2026</span>, total assets stood at <strong>BDT <span id="es-assets">0</span></strong> against total liabilities of <strong>BDT <span id="es-liab">0</span></strong>, resulting in total partners' equity of <strong>BDT <span id="es-eq">0</span></strong>. Partners' equity finances <strong><span id="es-eq-pct">0.0%</span></strong> of total assets while liabilities account for <strong><span id="es-liab-pct">0.0%</span></strong>.
                            </p>
                        </div>
                    </div>

                    <!-- 2.3 Key Financial Highlights Table -->
                    <div>
                        <h3 class="font-bold text-slate-900 text-xs uppercase tracking-wide mb-1.5">2.3 Key Financial Highlights</h3>
                        <table class="w-full text-[11px] border-collapse border border-slate-300">
                            <thead class="bg-slate-900 text-white">
                                <tr>
                                    <th class="border border-slate-400 px-3 py-1.5 text-left">Financial Indicator</th>
                                    <th class="border border-slate-400 px-3 py-1.5 text-right">Reported Amount (BDT)</th>
                                    <th class="border border-slate-400 px-3 py-1.5 text-left">Financial Indicator</th>
                                    <th class="border border-slate-400 px-3 py-1.5 text-right">Reported Amount (BDT)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Total Revenue</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono font-semibold" id="es-hl-rev">BDT 0</td>
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Total Assets</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono font-semibold" id="es-hl-assets">BDT 0</td>
                                </tr>
                                <tr class="bg-slate-50">
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Gross Profit</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono" id="es-hl-gp">BDT 0</td>
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Total Liabilities</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono" id="es-hl-liab">BDT 0</td>
                                </tr>
                                <tr>
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Net Profit</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono font-semibold text-emerald-700" id="es-hl-np">BDT 0</td>
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Total Partners' Equity</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono font-semibold" id="es-hl-eq">BDT 0</td>
                                </tr>
                                <tr class="bg-slate-50">
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Net Operating Cash Flow</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono" id="es-hl-cfo">BDT 0</td>
                                    <td class="border border-slate-300 px-3 py-1 font-medium">Closing Cash & Equivalents</td>
                                    <td class="border border-slate-300 px-3 py-1 text-right font-mono font-semibold" id="es-hl-cash">BDT 0</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>

                    <!-- 2.4 Key Financial Ratios and Performance Analysis -->
                    <div>
                        <h3 class="font-bold text-slate-900 text-xs uppercase tracking-wide mb-2">2.4 Key Financial Ratios and Performance Analysis</h3>
                        <div class="grid grid-cols-2 gap-2.5">
                            <div class="border border-slate-300 rounded p-2.5 bg-slate-50">
                                <div class="flex justify-between items-baseline">
                                    <span class="font-bold text-slate-900 text-[11px]">A. Gross Profit Margin</span>
                                    <span class="text-base font-bold text-slate-900" id="ratio-gpm">0.0%</span>
                                </div>
                                <p class="text-[10px] text-slate-600 mt-1" id="formula-gpm">Gross Profit ÷ Total Revenue × 100</p>
                            </div>
                            <div class="border border-slate-300 rounded p-2.5 bg-slate-50">
                                <div class="flex justify-between items-baseline">
                                    <span class="font-bold text-slate-900 text-[11px]">B. Net Profit Margin</span>
                                    <span class="text-base font-bold text-emerald-700" id="ratio-npm">0.0%</span>
                                </div>
                                <p class="text-[10px] text-slate-600 mt-1" id="formula-npm">Net Profit ÷ Total Revenue × 100</p>
                            </div>
                            <div class="border border-slate-300 rounded p-2.5 bg-slate-50">
                                <div class="flex justify-between items-baseline">
                                    <span class="font-bold text-slate-900 text-[11px]">C. Current Ratio</span>
                                    <span class="text-base font-bold text-slate-900" id="ratio-curr">0.00x</span>
                                </div>
                                <p class="text-[10px] text-slate-600 mt-1" id="formula-curr">Current Assets ÷ Current Liabilities</p>
                            </div>
                            <div class="border border-slate-300 rounded p-2.5 bg-slate-50">
                                <div class="flex justify-between items-baseline">
                                    <span class="font-bold text-slate-900 text-[11px]">D. Return on Equity (ROE)</span>
                                    <span class="text-base font-bold text-slate-900" id="ratio-roe">0.0%</span>
                                </div>
                                <p class="text-[10px] text-slate-600 mt-1" id="formula-roe">Net Profit ÷ Partners' Equity × 100</p>
                            </div>
                        </div>
                    </div>

                    <!-- 2.5 & 2.6 -->
                    <div class="space-y-1.5">
                        <h3 class="font-bold text-slate-900 text-xs uppercase tracking-wide border-b border-slate-200 pb-0.5">2.5 Cash Flow Performance & 2.6 Overall Assessment</h3>
                        <p class="text-slate-700">
                            The Statement of Cash Flows demonstrates net operating cash inflows of <strong>BDT <span id="es-cfo">0</span></strong>, supporting capital expenditure and partner distributions while closing with cash and cash equivalents of <strong>BDT <span id="es-cash">0</span></strong>. Management considers that <span class="doc-firm-name">Padma & Meghna Associates</span> maintained a positive financial position while recognizing the need for continuous monitoring of operating expenses, working capital, and obligations.
                        </p>
                    </div>

                    <!-- 2.7 Strategic Priorities -->
                    <div class="space-y-1">
                        <h3 class="font-bold text-slate-900 text-xs uppercase tracking-wide border-b border-slate-200 pb-0.5">2.7 Management's Strategic Priorities and Future Outlook</h3>
                        <div class="grid grid-cols-2 gap-x-4 gap-y-0.5 text-[10.5px] text-slate-700">
                            <p><strong>1. Revenue Growth:</strong> Expanding institutional & regional markets.</p>
                            <p><strong>2. Cost Optimization:</strong> Monitoring administrative & direct expenditure.</p>
                            <p><strong>3. Liquidity Management:</strong> Preserving cash buffers for obligations.</p>
                            <p><strong>4. Capital Management:</strong> Balanced reinvestment & partner drawings.</p>
                            <p><strong>5. Internal Controls:</strong> Strengthening financial documentation & reviews.</p>
                            <p><strong>6. Risk & Transparency:</strong> Proactive risk tracking & IFRS compliance.</p>
                        </div>
                    </div>

                    <!-- 2.8 Management's Overall Conclusion -->
                    <div class="bg-slate-50 border-l-4 border-emerald-700 p-2.5 text-[11px] text-slate-700">
                        <p class="font-bold text-slate-900 mb-0.5">2.8 Management's Overall Conclusion</p>
                        <p>
                            Considering the financial results and financial position as at <span class="doc-report-date">31 December 2026</span>, management considers that <strong class="doc-firm-name">Padma & Meghna Associates</strong> demonstrated positive financial performance during the reporting period. All amounts are presented in Bangladeshi Taka (BDT), unless otherwise stated.
                        </p>
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-2.5 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Executive Summary — FY Ended <span class="doc-report-date">31 December 2026</span></span>
                </div>
            </section>

            <!-- ==================== PAGE 5: SECTION B INDEPENDENT AUDITOR'S REPORT ==================== -->
            <section id="doc-audit-block" class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-4 text-[13px] leading-relaxed">
                    <div class="border-b-2 border-slate-900 pb-2">
                        <p class="text-xs font-bold uppercase tracking-wider text-emerald-700">Section B — Independent Assurance (ISA 700)</p>
                        <h2 class="report-serif text-xl font-bold text-slate-900 uppercase">3. Independent Auditor's Report</h2>
                    </div>

                    <p class="font-semibold text-slate-900">To the Partners of <span class="doc-firm-name">Padma & Meghna Associates</span></p>

                    <div>
                        <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wide mb-1">Opinion</h4>
                        <p class="text-slate-700">
                            We have audited the financial statements of <strong class="doc-firm-name">Padma & Meghna Associates</strong> ("the Firm"), which comprise the Statement of Financial Position as at <span class="doc-report-date">31 December 2026</span>, and the Statement of Profit or Loss and Other Comprehensive Income, Statement of Changes in Partners' Equity, and Statement of Cash Flows for the year then ended, and notes to the financial statements, including material accounting policy information.
                        </p>
                        <p class="text-slate-700 mt-2">
                            In our opinion, the accompanying financial statements present fairly, in all material respects, the financial position of the Firm as at <span class="doc-report-date">31 December 2026</span>, and its financial performance and its cash flows for the year then ended in accordance with International Financial Reporting Standards (IFRS).
                        </p>
                    </div>

                    <div>
                        <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wide mb-1">Basis for Opinion</h4>
                        <p class="text-slate-700">
                            We conducted our audit in accordance with International Standards on Auditing (ISAs). Our responsibilities under those standards are further described in the <em>Auditor's Responsibilities for the Audit of the Financial Statements</em> section of our report. We are independent of the Firm in accordance with the ethical requirements that are relevant to our audit of the financial statements in Bangladesh, and we have fulfilled our other ethical responsibilities in accordance with these requirements.
                        </p>
                    </div>

                    <div>
                        <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wide mb-1">Other Information (Management Report & Executive Summary)</h4>
                        <p class="text-slate-700">
                            Management is responsible for the other information contained in Section A (Management Report and Executive Summary). Our opinion on the financial statements does not cover the other information and we do not express any form of assurance conclusion thereon.
                        </p>
                    </div>

                    <div>
                        <h4 class="font-bold text-slate-900 text-xs uppercase tracking-wide mb-1">Responsibilities of Management and Those Charged with Governance</h4>
                        <p class="text-slate-700">
                            Management is responsible for the preparation and fair presentation of the financial statements in accordance with IFRS, and for such internal control as management determines is necessary to enable the preparation of financial statements that are free from material misstatement, whether due to fraud or error. In preparing the financial statements, management is responsible for assessing the Firm's ability to continue as a going concern.
                        </p>
                    </div>
                </div>

                <div class="pt-4 border-t border-slate-300 flex justify-between items-end text-xs">
                    <div>
                        <p class="font-bold text-slate-900" id="doc-auditor-firm">Rahman, Chowdhury & Co., Chartered Accountants</p>
                        <p class="text-slate-600" id="doc-auditor-partner">Mahmudul Hasan, FCA (Enrol: 1492)</p>
                        <p class="text-slate-500 mt-0.5">Dhaka, Bangladesh</p>
                    </div>
                    <div class="text-right">
                        <p class="text-slate-600">Date: <strong class="doc-approval-date">28 February 2027</strong></p>
                    </div>
                </div>
            </section>

            <!-- ==================== PAGE 6: SECTION C.1 STATEMENT OF FINANCIAL POSITION ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div>
                    <div class="text-center mb-5">
                        <p class="text-[11px] font-bold uppercase tracking-widest text-emerald-700 mb-0.5">Section C — Financial Statements (IAS 1)</p>
                        <h2 class="report-serif text-lg font-bold uppercase text-slate-900 doc-firm-name">Padma & Meghna Associates</h2>
                        <h3 class="text-sm font-bold uppercase text-slate-800">Statement of Financial Position</h3>
                        <p class="text-xs text-slate-600">As at <span class="doc-report-date">31 December 2026</span></p>
                    </div>

                    <table class="w-full text-xs border-collapse">
                        <thead>
                            <tr class="border-y-2 border-slate-900 bg-slate-50">
                                <th class="py-2 px-2 text-left font-bold">Particulars</th>
                                <th class="py-2 px-2 text-center font-bold w-16">Notes</th>
                                <th class="py-2 px-2 text-right font-bold w-32"><span class="doc-report-date">31 Dec 2026</span><br><span class="font-normal text-[10px]">(BDT)</span></th>
                                <th class="py-2 px-2 text-right font-bold w-32"><span class="doc-prev-date">31 Dec 2025</span><br><span class="font-normal text-[10px]">(BDT)</span></th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-200" id="bs-preview-tbody">
                            <!-- Dynamically populated -->
                        </tbody>
                    </table>
                    <p class="text-[11px] text-slate-500 mt-3 italic">The annexed notes form an integral part of these financial statements.</p>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-3 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Statement of Financial Position — As at <span class="doc-report-date">31 December 2026</span></span>
                </div>
            </section>

            <!-- ==================== PAGE 7: SECTION C.2 & C.3 PROFIT OR LOSS & PARTNERS' EQUITY ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-5">
                    <div>
                        <div class="text-center mb-3">
                            <h2 class="report-serif text-base font-bold uppercase text-slate-900 doc-firm-name">Padma & Meghna Associates</h2>
                            <h3 class="text-xs font-bold uppercase text-slate-800">Statement of Profit or Loss and Other Comprehensive Income</h3>
                            <p class="text-[11px] text-slate-600">For the financial year ended <span class="doc-report-date">31 December 2026</span> (IAS 1)</p>
                        </div>

                        <table class="w-full text-[11px] border-collapse">
                            <thead>
                                <tr class="border-y-2 border-slate-900 bg-slate-50">
                                    <th class="py-1.5 px-2 text-left font-bold">Particulars</th>
                                    <th class="py-1.5 px-2 text-center font-bold w-14">Notes</th>
                                    <th class="py-1.5 px-2 text-right font-bold w-28">Current Year (BDT)</th>
                                    <th class="py-1.5 px-2 text-right font-bold w-28">Previous Year (BDT)</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-200" id="pl-preview-tbody">
                                <!-- Dynamically populated -->
                            </tbody>
                        </table>

                        <!-- Profit Allocation Sub-schedule -->
                        <div class="mt-2.5 border border-slate-300 rounded p-2.5 bg-slate-50">
                            <h4 class="text-[11px] font-bold uppercase text-slate-800 mb-1">Allocation of Profit Among Partners (Note 14)</h4>
                            <table class="w-full text-[11px]" id="pl-alloc-table">
                                <!-- Dynamically populated -->
                            </table>
                        </div>
                    </div>

                    <!-- Statement of Changes in Partners' Equity -->
                    <div class="pt-2 border-t border-slate-300">
                        <div class="text-center mb-2.5">
                            <h3 class="text-xs font-bold uppercase text-slate-800">Statement of Changes in Partners' Equity</h3>
                            <p class="text-[11px] text-slate-600">For the financial year ended <span class="doc-report-date">31 December 2026</span> (IAS 1 & IAS 32)</p>
                        </div>
                        <table class="w-full text-[11px] border-collapse" id="eq-preview-table">
                            <!-- Dynamically populated -->
                        </table>
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-2.5 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Profit or Loss & Changes in Partners' Equity</span>
                </div>
            </section>

            <!-- ==================== PAGE 8: SECTION C.4 STATEMENT OF CASH FLOWS ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div>
                    <div class="text-center mb-6">
                        <p class="text-[11px] font-bold uppercase tracking-widest text-emerald-700 mb-0.5">Section C — Financial Statements (IAS 7)</p>
                        <h2 class="report-serif text-lg font-bold uppercase text-slate-900 doc-firm-name">Padma & Meghna Associates</h2>
                        <h3 class="text-sm font-bold uppercase text-slate-800">Statement of Cash Flows</h3>
                        <p class="text-xs text-slate-600">For the financial year ended <span class="doc-report-date">31 December 2026</span></p>
                    </div>

                    <table class="w-full text-xs border-collapse">
                        <thead>
                            <tr class="border-y-2 border-slate-900 bg-slate-50">
                                <th class="py-2.5 px-2 text-left font-bold">Particulars</th>
                                <th class="py-2.5 px-2 text-right font-bold w-36">Current Year (BDT)</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-200" id="cf-preview-tbody">
                            <!-- Dynamically populated -->
                        </tbody>
                    </table>
                    <div id="cf-reconcile-note" class="mt-3 text-[11px] text-emerald-700 font-medium">
                        ✓ Reconciles directly with Cash and Cash Equivalents in the Statement of Financial Position.
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-3 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Statement of Cash Flows — FY Ended <span class="doc-report-date">31 December 2026</span></span>
                </div>
            </section>

            <!-- ==================== PAGE 9: SECTION C.5 NOTES TO THE FINANCIAL STATEMENTS (NOTES 1 TO 7) ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-3 text-[11px] leading-snug">
                    <div class="border-b-2 border-slate-900 pb-1.5">
                        <p class="text-[10px] font-bold uppercase tracking-wider text-emerald-700">Section C — Financial Statements</p>
                        <h2 class="report-serif text-base font-bold text-slate-900 uppercase">8. Notes to the Financial Statements (Notes 1 to 7)</h2>
                        <p class="text-[10px] text-slate-600">For the financial year ended <span class="doc-report-date">31 December 2026</span></p>
                    </div>

                    <div class="space-y-2.5 text-slate-700">
                        <div>
                            <h4 class="font-bold text-slate-900">Note 1 — General Information about the Partnership Firm</h4>
                            <p><strong class="doc-firm-name">Padma & Meghna Associates</strong> is a registered partnership firm domiciled in Bangladesh (Reg. No. <span class="doc-reg-no">P-48291/2021</span>), established on <span class="doc-est-date">14 March 2021</span>, with its principal place of business at <span class="doc-addr">Level 7, Gulshan Avenue, Dhaka-1212, Bangladesh</span>.</p>
                        </div>

                        <div>
                            <h4 class="font-bold text-slate-900">Note 2 — Basis of Preparation & Going Concern (IAS 1)</h4>
                            <p>These financial statements have been prepared on a historical cost basis under the accrual basis of accounting and on a going concern assumption. Management has assessed the Firm's ability to continue as a going concern and identified no material uncertainties. All figures are presented in Bangladeshi Taka (BDT).</p>
                        </div>

                        <div>
                            <h4 class="font-bold text-slate-900">Note 3 — Statement of Compliance with IFRS (IAS 1.16)</h4>
                            <p>The financial statements have been prepared in full compliance with International Financial Reporting Standards (IFRS) and International Accounting Standards (IAS) as issued by the International Accounting Standards Board (IASB) and adopted in Bangladesh by the Financial Reporting Council (FRC).</p>
                        </div>

                        <div>
                            <h4 class="font-bold text-slate-900">Note 4 — Material Accounting Policy Information (IAS 1 / IAS 8)</h4>
                            <p><strong>Property, Plant & Equipment (IAS 16):</strong> Stated at cost less accumulated depreciation and impairment losses. Depreciation is charged on a straight-line basis over estimated useful lives. <strong>Intangible Assets (IAS 38):</strong> Software and licences amortised over 5 years. <strong>Inventories (IAS 2):</strong> Valued at the lower of cost (weighted average) and net realisable value. <strong>Financial Instruments (IFRS 9):</strong> Trade receivables are measured at amortised cost less expected credit losses (ECL). <strong>Partners' Capital (IAS 32):</strong> Classified as equity as the Firm has no unconditional obligation to deliver cash.</p>
                        </div>

                        <div>
                            <h4 class="font-bold text-slate-900">Note 5 — Significant Accounting Judgements and Estimates (IAS 1.122 / 1.125)</h4>
                            <p>Key estimates involve useful lives of depreciable assets, expected credit loss allowances on trade receivables, and net realisable value write-downs for perishable or slow-moving inventory.</p>
                        </div>

                        <div>
                            <h4 class="font-bold text-slate-900">Note 6 — Revenue from Contracts with Customers (IFRS 15)</h4>
                            <p>Revenue is recognised at a point in time when control of commercial goods or distribution services transfers to the customer, measured at the transaction price net of trade discounts and VAT.</p>
                        </div>

                        <div>
                            <h4 class="font-bold text-slate-900">Note 7 — Cost of Sales (IAS 2)</h4>
                            <p>Cost of sales comprises direct procurement costs of agro-commodities, inbound freight, cold-storage handling, and packaging materials directly attributable to goods sold during the year.</p>
                        </div>
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-2.5 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>8. Notes to the Financial Statements (Basis & Accounting Policies)</span>
                </div>
            </section>

            <!-- ==================== PAGE 10: SECTION C.5 NOTES CONTINUED (NOTES 8 TO 14) ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-2 text-[10px] leading-tight">
                    <div class="border-b-2 border-slate-900 pb-1 flex justify-between items-end">
                        <div>
                            <h2 class="report-serif text-sm font-bold text-slate-900 uppercase">Notes to the Financial Statements (Continued — Notes 8 to 14)</h2>
                        </div>
                        <span class="text-[10px] font-semibold text-slate-600">FY Ended <span class="doc-report-date">31 December 2026</span></span>
                    </div>

                    <div id="notes-8-14-container" class="space-y-2">
                        <!-- Dynamically rendered Notes 8 to 14 with interactive tables -->
                    </div>
                </div>

                <div class="text-[10.5px] text-slate-500 border-t border-slate-200 pt-2 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>8. Notes to the Financial Statements (Assets & Partners' Capital)</span>
                </div>
            </section>

            <!-- ==================== PAGE 11: SECTION C.5 NOTES CONTINUED (NOTES 15 TO 25) ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-1.5 text-[9.5px] leading-tight">
                    <div class="border-b-2 border-slate-900 pb-1 flex justify-between items-end">
                        <div>
                            <h2 class="report-serif text-sm font-bold text-slate-900 uppercase">8. Notes to the Financial Statements (Continued — Notes 15 to 25)</h2>
                        </div>
                        <span class="text-[10px] font-semibold text-slate-600">FY Ended <span class="doc-report-date">31 December 2026</span></span>
                    </div>

                    <div id="notes-15-25-container" class="space-y-1.5">
                        <!-- Dynamically rendered Notes 15 to 25 + custom notes -->
                    </div>
                </div>

                <div class="text-[10.5px] text-slate-500 border-t border-slate-200 pt-2 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>8. Notes to the Financial Statements (Liabilities, Tax & Disclosures)</span>
                </div>
            </section>

            <!-- ==================== PAGE 12: SECTION C.6 TAXATION SCHEDULE (IAS 12 & S.R.O. 324/2026) ==================== -->
            <section id="doc-tax-page" class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-3 text-[11px] leading-snug">
                    <div class="border-b-2 border-slate-900 pb-1.5">
                        <p class="text-[10px] font-bold uppercase tracking-wider text-emerald-700">Section C — Financial Statements (IAS 12 & Income Tax Act 2023)</p>
                        <h2 class="report-serif text-base font-bold text-slate-900 uppercase">Note 19: Current and Deferred Taxation (<span id="doc-tax-period">FY 2025-2026</span>)</h2>
                    </div>

                    <div class="bg-slate-50 border border-slate-300 rounded p-2.5 text-[10.5px] text-slate-700 space-y-1">
                        <p>
                            <strong>Statutory Framework (S.R.O. No. 324-Ain/Aykor-05/2026):</strong> Issued on 31 August 2026 by the National Board of Revenue (NBR) under Section 76(1) read with Section 163(6) of the Income Tax Act 2023, replacing the flat 1% turnover tax rate with a slab-based minimum turnover tax structure:
                        </p>
                        <ul class="list-disc pl-5 space-y-0.5">
                            <li>Gross receipts up to Tk. 20 million (2 crore): <strong>0%</strong> turnover tax rate.</li>
                            <li>Gross receipts exceeding Tk. 20 million up to Tk. 40 million (4 crore): <strong>0.50%</strong> turnover tax rate.</li>
                            <li>Gross receipts exceeding Tk. 40 million (4 crore): <strong>1.00%</strong> turnover tax rate.</li>
                        </ul>
                        <p id="doc-tax-slab-desc" class="pt-0.5">
                            The entity's contract turnover falls into the second slab (&gt; Tk. 20M to Tk. 40M), subjecting it to a <strong>0.50%</strong> minimum turnover tax. Corporate income tax is determined as the higher of Minimum Tax (A) and Regular Tax @ 27.50% on Taxable Profit (B).
                        </p>
                    </div>

                    <table class="w-full text-[10.5px] border-collapse border border-slate-400">
                        <thead class="bg-slate-900 text-white">
                            <tr>
                                <th class="border border-slate-400 px-2 py-1.5 text-center w-8">SL</th>
                                <th class="border border-slate-400 px-2 py-1.5 text-left">Corporate Tax Assessment Item</th>
                                <th class="border border-slate-400 px-2 py-1.5 text-left w-36">Statutory Reference / Rate</th>
                                <th class="border border-slate-400 px-2 py-1.5 text-left w-40">Calculation Base</th>
                                <th class="border border-slate-400 px-2 py-1.5 text-right w-28">Amount (BDT)</th>
                            </tr>
                        </thead>
                        <tbody id="tax-schedule-tbody" class="divide-y divide-slate-300">
                            <!-- Dynamically populated -->
                        </tbody>
                    </table>

                    <div class="bg-emerald-50/70 border-l-4 border-emerald-700 p-2.5 text-[10.5px] text-slate-800">
                        <p class="font-bold text-emerald-900 mb-0.5">Note on Statutory Assessment & IAS 12 Reconciliation:</p>
                        <p id="doc-tax-statutory-note" class="leading-relaxed"></p>
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-2.5 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Current & Deferred Taxation (IAS 12 & S.R.O. 324/2026)</span>
                </div>
            </section>

            <!-- ==================== PAGE 13: SECTION D CLOSING & AUTHORIZATION ==================== -->
            <section class="a4-sheet flex flex-col justify-between text-slate-900">
                <div class="space-y-4 text-[12px] leading-relaxed">
                    <div class="border-b-2 border-slate-900 pb-2">
                        <p class="text-xs font-bold uppercase tracking-wider text-emerald-700">Section D — Closing and Authorization</p>
                        <h2 class="report-serif text-xl font-bold text-slate-900 uppercase">6. Management's Conclusion</h2>
                    </div>

                    <p class="text-slate-700">
                        Management of <strong class="doc-firm-name">Padma & Meghna Associates</strong> has reviewed the financial statements and accompanying notes for the financial year ended <strong class="doc-report-date">31 December 2026</strong>, including the Statement of Financial Position, Statement of Profit or Loss and Other Comprehensive Income, Statement of Cash Flows, and Statement of Changes in Partners' Equity.
                    </p>
                    <p class="text-slate-700">
                        Based on this review, management considers that the financial statements provide a structured presentation of the firm's financial performance, financial position, and cash flow activities for the reporting period. The statements reflect the firm's assets, liabilities, partners' capital contributions, revenue, operating expenses, financial results, and other relevant financial transactions as recorded in the underlying accounting records.
                    </p>
                    <p class="text-slate-700">
                        Management acknowledges that the preparation and fair presentation of financial statements are the responsibility of the partners. This responsibility includes maintaining complete and accurate accounting records, selecting and consistently applying appropriate accounting policies, making reasonable and supportable accounting estimates and judgments, and ensuring that financial transactions are properly recognized, classified, measured, and disclosed.
                    </p>
                    <p class="text-slate-700">
                        Management further recognizes its responsibility for establishing and maintaining appropriate internal financial controls and accounting procedures to safeguard the firm's assets, prevent and detect material errors or irregularities, ensure the reliability of financial information, and promote compliance with applicable laws, regulations, and financial reporting requirements.
                    </p>
                    <p class="text-slate-700">
                        The financial statements have been prepared with reference to the applicable requirements of International Financial Reporting Standards (IFRS) and International Accounting Standards (IAS), as relevant to the firm's circumstances and reporting obligations. Management has also considered the going concern assumption in preparing these financial statements and believes that the firm has an appropriate basis for continuing its business operations, subject to the firm's actual financial circumstances and future business conditions.
                    </p>
                    <p class="text-slate-700">
                        Looking forward, <strong class="doc-firm-name">Padma & Meghna Associates</strong> intends to continue strengthening its financial discipline, improving operational efficiency, enhancing internal control procedures, and identifying sustainable business opportunities. Management believes that the financial statements, together with the accompanying notes, should be read as a complete set of financial information for understanding the firm's financial performance and position for the year ended <span class="doc-report-date">31 December 2026</span>.
                    </p>

                    <div class="bg-slate-50 border border-slate-300 rounded-lg p-5 mt-4">
                        <h3 class="font-bold text-slate-900 uppercase text-xs tracking-wider mb-2">7. Approval and Authorization of Financial Statements</h3>
                        <p class="text-[11.5px] text-slate-700 mb-2">
                            The partners of <strong class="doc-firm-name">Padma & Meghna Associates</strong> acknowledge their responsibility for reviewing and approving the financial statements for the financial year ended <strong class="doc-report-date">31 December 2026</strong>. Following consideration of the financial statements and the accompanying explanatory notes, the partners confirm that the financial statements have been presented in accordance with the firm's accounting records and the applicable financial reporting framework.
                        </p>
                        <p class="text-[11.5px] text-slate-700 mb-6">
                            The partners hereby approve and authorize the issue of the financial statements of <strong class="doc-firm-name">Padma & Meghna Associates</strong> for the financial year ended <strong class="doc-report-date">31 December 2026</strong>, together with all accompanying notes and disclosures, on <strong class="doc-approval-date">28 February 2027</strong> in Dhaka, Bangladesh.
                        </p>

                        <p class="text-xs font-bold text-slate-900 uppercase tracking-wide mb-6">Authorized and Approved by</p>

                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6 pt-2" id="signatures-preview-grid">
                            <!-- Dynamically populated for all partners -->
                        </div>
                    </div>
                </div>

                <div class="text-[11px] text-slate-500 border-t border-slate-200 pt-3 flex justify-between">
                    <span class="doc-firm-name">Padma & Meghna Associates</span>
                    <span>Approval & Partners' Signatures — FY Ended <span class="doc-report-date">31 December 2026</span></span>
                </div>
            </section>

            <!-- ==================== SECTION E: NBR INCOME TAX RETURN FORM IT-116 (2023) ==================== -->
            <div id="it116-return-wrapper">

                <!-- IT-116 PAGE 1: TAXPAYER PROFILE -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="flex justify-between items-start mb-4">
                            <div class="it116-office-box">
                                <div class="row"><div class="col-1 text-center font-bold">For Official Use</div><div class="col-2"></div></div>
                                <div class="row"><div class="col-1">Return Register Serial No.</div><div class="col-2"></div></div>
                                <div class="row"><div class="col-1">Date of Submission of Return</div><div class="col-2"></div></div>
                            </div>

                            <div class="text-center flex-1 px-3">
                                <div class="font-bold text-[15px]">National Board of Revenue</div>
                                <div class="font-bold text-[12px]">www.nbr.gov.bd</div>
                                <div class="font-bold text-[15px] mt-2.5">Income Tax Return</div>
                                <div class="text-[11px]">[Applicable for taxpayers other than company<br>taxpayers as defined in clause (31) of section<br>2 of the Income Tax Act, 2023]</div>
                            </div>

                            <div class="text-right font-bold whitespace-nowrap text-[13px]">IT-116 (2023)</div>
                        </div>

                        <div class="it116-main-box space-y-2.5">
                            <div class="flex items-center">
                                <span class="font-bold whitespace-nowrap">1. Name of Taxpayer:</span>
                                <input type="text" id="it116-p1-name" class="it116-line-input font-semibold">
                            </div>

                            <div class="flex items-center">
                                <span class="font-bold whitespace-nowrap">2. TIN:</span>
                                <div class="it116-tin-container" data-tin-group="p1"></div>
                            </div>

                            <div class="flex items-center">
                                <span class="font-bold whitespace-nowrap">3. Business Identification Number (BIN) (s):</span>
                                <input type="text" id="it116-p1-bin" class="it116-line-input">
                            </div>

                            <div class="flex justify-between gap-4">
                                <div class="w-1/2 flex items-center">
                                    <span class="font-bold whitespace-nowrap">4. (a) Circle:</span>
                                    <input type="text" id="it116-p1-circle" class="it116-line-input">
                                </div>
                                <div class="w-1/2 flex items-center">
                                    <span class="font-bold whitespace-nowrap">(b) Tax Zone:</span>
                                    <input type="text" id="it116-p1-zone" class="it116-line-input">
                                </div>
                            </div>

                            <div class="flex items-center">
                                <span class="font-bold whitespace-nowrap">5. Tax Year:</span>
                                <input type="text" id="it116-p1-taxyear" class="it116-line-input">
                            </div>

                            <div class="flex flex-wrap items-center gap-x-3 gap-y-1">
                                <span class="font-bold whitespace-nowrap">6. Category of Taxpayer (Tick (√) mark where applicable):</span>
                                <label class="inline-flex items-center">Firm <input type="checkbox" id="it116-cat-firm" class="ml-1 accent-black"></label>
                                <label class="inline-flex items-center">Association of Persons <input type="checkbox" id="it116-cat-aop" class="ml-1 accent-black"></label>
                                <label class="inline-flex items-center">Fund <input type="checkbox" id="it116-cat-fund" class="ml-1 accent-black"></label>
                                <label class="inline-flex items-center">Trust <input type="checkbox" id="it116-cat-trust" class="ml-1 accent-black"></label>
                                <label class="inline-flex items-center">Hindu Undivided Family <input type="checkbox" id="it116-cat-huf" class="ml-1 accent-black"></label>
                            </div>

                            <div class="flex items-center gap-x-6">
                                <span class="font-bold whitespace-nowrap">7. Residential Status:</span>
                                <label class="inline-flex items-center">Resident <input type="checkbox" id="it116-res-yes" class="ml-1.5 accent-black"></label>
                                <label class="inline-flex items-center">Non-resident <input type="checkbox" id="it116-res-no" class="ml-1.5 accent-black"></label>
                            </div>

                            <div class="space-y-1.5 pt-1">
                                <div class="flex items-center">
                                    <span class="font-bold whitespace-nowrap">8. Contact Address:</span>
                                    <input type="text" id="it116-p1-addr1" class="it116-line-input">
                                </div>
                                <div class="flex"><input type="text" id="it116-p1-addr2" class="it116-line-input !ml-0"></div>
                            </div>

                            <div class="flex justify-between gap-4">
                                <div class="w-1/2 flex items-center">
                                    <span class="whitespace-nowrap">Telephone:</span>
                                    <input type="text" id="it116-p1-phone" class="it116-line-input">
                                </div>
                                <div class="w-1/2 flex items-center">
                                    <span class="whitespace-nowrap">E-mail:</span>
                                    <input type="email" id="it116-p1-email" class="it116-line-input">
                                </div>
                            </div>

                            <div class="flex justify-between gap-4 pt-1">
                                <div class="w-1/2 flex items-center">
                                    <span class="font-bold whitespace-nowrap">9. Registration No:</span>
                                    <input type="text" id="it116-p1-regno" class="it116-line-input">
                                </div>
                                <div class="w-1/2 flex items-center">
                                    <span class="font-bold whitespace-nowrap">10. Registration Date:</span>
                                    <input type="date" id="it116-p1-regdate" class="it116-line-input">
                                </div>
                            </div>

                            <div class="pt-1.5 space-y-1.5">
                                <div class="font-bold">11. In case of Firm/Association of Persons, Name and TIN of partners/members (use separate sheet if necessary):</div>
                                <div id="it116-p1-partners-lines" class="space-y-1.5"></div>
                            </div>
                        </div>
                    </div>
                    <div class="text-center font-bold pt-2">1</div>
                </section>

                <!-- IT-116 PAGE 2: STATEMENT OF TOTAL INCOME, TAX COMPUTATION & PAYMENT -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="flex justify-end items-center mb-3">
                            <span class="mr-2">Statement of Income and Tax for the income year ending on date:</span>
                            <input type="text" id="it116-p2-enddate" class="it116-line-input !flex-grow-0 w-44 text-center font-semibold">
                        </div>

                        <div class="flex justify-between items-center gap-4 mb-3">
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Name of Taxpayer:</span>
                                <input type="text" id="it116-p2-name" class="it116-line-input font-semibold">
                            </div>
                            <div class="flex items-center w-1/2 justify-end">
                                <span class="font-bold mr-1">TIN:</span>
                                <div class="it116-tin-container" data-tin-group="p2"></div>
                            </div>
                        </div>

                        <div class="flex justify-between font-bold text-[13px] mb-1">
                            <span>Statement of Total Income</span>
                            <span>Amount in Taka</span>
                        </div>
                        <table class="it116-table">
                            <tr><td class="w-[6%]">1.</td><td class="w-[69%]">Income from Rent (as per Schedule 1 of this return)</td><td class="w-[25%]"><input type="text" id="it116-inc-1" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>2.</td><td>Income from Agriculture (as per Schedule 2 of this return)</td><td><input type="text" id="it116-inc-2" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>3.</td><td>Income from Business (as per Schedule 3 of this return)</td><td><input type="text" id="it116-inc-3" class="it116-cell-input text-right font-mono font-semibold"></td></tr>
                            <tr><td>4.</td><td>Capital Gain</td><td><input type="text" id="it116-inc-4" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>5.</td><td>Income from Financial Assets (Bank interest/profit, dividend, securities, etc.)</td><td><input type="text" id="it116-inc-5" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>6.</td><td>Income from Other Sources (royalty, license fee, souvenir, fee, cash subsidy provided by government, etc.)</td><td><input type="text" id="it116-inc-6" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>7.</td><td>Foreign sourced taxable income</td><td><input type="text" id="it116-inc-7" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>8.</td><td>Total Income (sum of serial 1 to 7)</td><td><input type="text" id="it116-inc-8" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                        </table>

                        <div class="flex justify-between font-bold text-[13px] mt-3 mb-1">
                            <span>Computation of Tax</span>
                            <span>Amount in Taka</span>
                        </div>
                        <table class="it116-table">
                            <tr><td class="w-[6%]">9.</td><td class="w-[69%]">Income tax imposable on total taxable income</td><td class="w-[25%]"><input type="text" id="it116-tax-9" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>10.</td><td>Minimum Tax</td><td><input type="text" id="it116-tax-10" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>11.</td><td>Turnover Tax</td><td><input type="text" id="it116-tax-11" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>12.</td><td>Tax Payable (whichever is higher among serial 9, 10 and 11)</td><td><input type="text" id="it116-tax-12" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                            <tr><td>13.</td><td>Environment Surcharge</td><td><input type="text" id="it116-tax-13" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>14.</td><td>Delay interest, penalty or any other amount payable under the Income Tax Act (if any)</td><td><input type="text" id="it116-tax-14" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>15.</td><td>Total Tax Payable (12+13+14)</td><td><input type="text" id="it116-tax-15" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                        </table>

                        <div class="flex justify-between font-bold text-[13px] mt-3 mb-1">
                            <span>Statement of Tax Payment</span>
                            <span>Amount in Taka</span>
                        </div>
                        <table class="it116-table">
                            <tr><td class="w-[6%]">16.</td><td class="w-[69%]">Tax deducted/collected at source (attach proof)</td><td class="w-[25%]"><input type="text" id="it116-pay-16" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>17.</td><td>Advance tax paid (attach proof)</td><td><input type="text" id="it116-pay-17" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>18.</td><td>Time limit for refundable tax (if any) <span class="text-[10px]">(Mention the relevant tax year/years for refund)</span></td><td><input type="text" id="it116-pay-18" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>19.</td><td>Balance tax paid along with this return (attach proof)</td><td><input type="text" id="it116-pay-19" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>20.</td><td>Tax paid (sum of serial 16 to 19)</td><td><input type="text" id="it116-pay-20" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                            <tr><td>21.</td><td>Excess payment (if any)</td><td><input type="text" id="it116-pay-21" class="it116-cell-input text-right font-mono"></td></tr>
                        </table>

                        <table class="it116-table mt-2">
                            <tr><td class="w-[6%]">22.</td><td class="w-[69%]">Tax exempted/tax free income (attach details)</td><td class="w-[25%]"><input type="text" id="it116-pay-22" class="it116-cell-input text-right font-mono"></td></tr>
                        </table>
                    </div>
                    <div class="text-center font-bold pt-2">2</div>
                </section>

                <!-- IT-116 PAGE 3: SCHEDULE 1 (INCOME FROM RENT) -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="text-center font-bold text-[15px] mb-1">Schedule 1</div>
                        <div class="text-center mb-4">If there is income from rent, the following schedule must be filled.</div>

                        <div class="flex justify-between items-center gap-4 mb-4">
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Name of Taxpayer:</span>
                                <input type="text" id="it116-s1-name" class="it116-line-input font-semibold">
                            </div>
                            <div class="flex items-center w-1/2 justify-end">
                                <span class="font-bold mr-1">TIN:</span>
                                <div class="it116-tin-container" data-tin-group="s1"></div>
                            </div>
                        </div>

                        <table class="it116-table">
                            <tr>
                                <th class="w-[25%]">Location, Description and Share of Ownership of Property</th>
                                <th class="w-[45%]">Computation of Total Rent Value</th>
                                <th class="w-[15%]">Amount in Taka</th>
                                <th class="w-[15%]">Amount in Taka</th>
                            </tr>
                            <tr>
                                <td rowspan="16"><textarea class="it116-cell-input h-80 resize-none" placeholder="Nil / Not applicable"></textarea></td>
                                <td>1. Amount of rent received or annual value, whichever is higher</td>
                                <td><input type="text" class="it116-cell-input text-right"></td>
                                <td rowspan="4"><input type="text" class="it116-cell-input text-right"></td>
                            </tr>
                            <tr><td>2. Advance rent received</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>3. Monetary value of any amount or benefit received (in addition to 1 and 2)</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>4. Adjusted advance amount</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>5. Vacancy allowance</td><td><input type="text" class="it116-cell-input text-right"></td><td rowspan="2"><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>6. Total rent value (1+2+3)-4-5</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>7. Permissible deductions:</td><td></td><td rowspan="7"><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td class="pl-5">(a) Repair collection, etc.</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td class="pl-5">(b) Municipal tax or local tax</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td class="pl-5">(c) Land revenue</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td class="pl-5">(d) Interest on loan/ mortgage/ capital charge</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td class="pl-5">(e) Insurance premium paid</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td class="pl-5">(f) Others (if any)</td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>8. Total permissible deduction</td><td><input type="text" class="it116-cell-input text-right"></td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr><td>9. Net income (deducting serial 8 from serial 6)</td><td><input type="text" class="it116-cell-input text-right"></td><td><input type="text" class="it116-cell-input text-right"></td></tr>
                            <tr class="font-bold"><td>10. Taxpayer's share (where applicable)</td><td><input type="text" class="it116-cell-input text-right"></td><td><input type="text" id="it116-s1-net" class="it116-cell-input text-right font-mono"></td></tr>
                        </table>
                    </div>
                    <div class="text-center font-bold pt-2">3</div>
                </section>

                <!-- IT-116 PAGE 4: SCHEDULE 2 (INCOME FROM AGRICULTURE) -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="text-center font-bold text-[15px] mb-1">Schedule 2</div>
                        <div class="text-center mb-4">If there is income from agriculture, the following schedule must be filled.</div>

                        <div class="flex justify-between items-center gap-4 mb-4">
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Name of Taxpayer:</span>
                                <input type="text" id="it116-s2-name" class="it116-line-input font-semibold">
                            </div>
                            <div class="flex items-center w-1/2 justify-end">
                                <span class="font-bold mr-1">TIN:</span>
                                <div class="it116-tin-container" data-tin-group="s2"></div>
                            </div>
                        </div>

                        <div class="flex items-center mb-3">
                            <span class="font-bold whitespace-nowrap">Type of agricultural work:</span>
                            <input type="text" value="Not applicable (Commercial & Agro-Commodity Trading reported under Schedule 3)" class="it116-line-input">
                        </div>

                        <table class="it116-table">
                            <tr>
                                <th class="w-[8%]">Serial<br>No.</th>
                                <th class="w-[67%]">Summary of Income</th>
                                <th class="w-[25%]">Amount in Taka</th>
                            </tr>
                            <tr><td>1.</td><td>Sales/ Turnover/ Receipts</td><td><input type="text" value="0" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>2.</td><td>Gross Profit</td><td><input type="text" value="0" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>3.</td><td>General expenses, investment expenses, land development tax, rent, interest on loan, insurance premium and other expenses</td><td><input type="text" value="0" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold"><td>4.</td><td>Net income (deducting serial 3 from serial 2)</td><td><input type="text" id="it116-s2-net" value="0" class="it116-cell-input text-right font-mono"></td></tr>
                        </table>
                    </div>
                    <div class="text-center font-bold pt-2">4</div>
                </section>

                <!-- IT-116 PAGE 5: SCHEDULE 3 (INCOME FROM BUSINESS & STATEMENT OF ASSETS/LIABILITIES) -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="text-center font-bold text-[15px] mb-0.5">Schedule 3</div>
                        <div class="text-center text-[12px] mb-2.5">If there is income from business or profession, the following schedule must be filled.</div>

                        <div class="flex justify-between items-center gap-4 mb-2">
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Name of Taxpayer:</span>
                                <input type="text" id="it116-s3-name" class="it116-line-input font-semibold">
                            </div>
                            <div class="flex items-center w-1/2 justify-end">
                                <span class="font-bold mr-1">TIN:</span>
                                <div class="it116-tin-container" data-tin-group="s3"></div>
                            </div>
                        </div>

                        <div class="flex justify-between gap-4 mb-2">
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Name of Business/Profession:</span>
                                <input type="text" id="it116-s3-bizname" class="it116-line-input">
                            </div>
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Type of Business/Profession:</span>
                                <input type="text" id="it116-s3-biztype" class="it116-line-input">
                            </div>
                        </div>

                        <div class="flex items-center mb-2.5">
                            <span class="font-bold whitespace-nowrap">Address:</span>
                            <input type="text" id="it116-s3-addr" class="it116-line-input">
                        </div>

                        <table class="it116-table !mb-2.5">
                            <tr>
                                <th class="w-[10%]">Serial No.</th>
                                <th class="w-[65%]">Summary of Income</th>
                                <th class="w-[25%]">Amount in Taka</th>
                            </tr>
                            <tr><td>1.</td><td>Sales/ Turnover/ Receipts</td><td><input type="text" id="it116-s3-inc1" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>2.</td><td>VAT (if any)</td><td><input type="text" id="it116-s3-inc2" value="0" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>3.</td><td>Net Sales/ Turnover/ Receipts</td><td><input type="text" id="it116-s3-inc3" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>4.</td><td>Cost of Goods Sold</td><td><input type="text" id="it116-s3-inc4" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>5.</td><td>Gross Profit (3-4)</td><td><input type="text" id="it116-s3-inc5" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                            <tr><td>6.</td><td>General administrative and other expenses</td><td><input type="text" id="it116-s3-inc6" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>7.</td><td>Selling and marketing expenses</td><td><input type="text" id="it116-s3-inc7" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>8.</td><td>Interest and other expenses</td><td><input type="text" id="it116-s3-inc8" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>9.</td><td>Net Profit (5-6-7-8)</td><td><input type="text" id="it116-s3-inc9" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                            <tr><td>10.</td><td>Other income (if any)</td><td><input type="text" id="it116-s3-inc10" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-100"><td>11.</td><td>Pre-tax income/ profit</td><td><input type="text" id="it116-s3-inc11" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                        </table>

                        <table class="it116-table !mb-0">
                            <tr>
                                <th class="w-[10%]">Serial No.</th>
                                <th class="w-[65%]">Statement of Assets</th>
                                <th class="w-[25%]">Amount in Taka</th>
                            </tr>
                            <tr><td>1.</td><td>Fixed assets (Plant, machinery, software, etc.)</td><td><input type="text" id="it116-s3-ast1" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>2.</td><td>Intangible assets</td><td><input type="text" id="it116-s3-ast2" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>3.</td><td>Financial assets</td><td><input type="text" id="it116-s3-ast3" value="0" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>4.</td><td>Current assets (Inventories & others)</td><td><input type="text" id="it116-s3-ast4" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>5.</td><td>Land, business and other receivables</td><td><input type="text" id="it116-s3-ast5" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>6.</td><td>Cash in hand</td><td><input type="text" id="it116-s3-ast6" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>7.</td><td>Bank balance</td><td><input type="text" id="it116-s3-ast7" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>8.</td><td>Total assets (sum of 01 to 07)</td><td><input type="text" id="it116-s3-ast8" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                            <tr><td colspan="2" class="text-center font-bold bg-slate-100">Capital and Liabilities</td><td></td></tr>
                            <tr><td>9.</td><td>Opening capital (including additional capital introduced)</td><td><input type="text" id="it116-s3-cap9" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>10.</td><td>Net profit / post-tax income allocated to equity</td><td><input type="text" id="it116-s3-cap10" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>11.</td><td>Withdrawal from business during the income year</td><td><input type="text" id="it116-s3-cap11" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-50"><td>12.</td><td>Closing capital (09+10-11)</td><td><input type="text" id="it116-s3-cap12" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                            <tr><td>13.</td><td>Institutional liabilities (Bank borrowings)</td><td><input type="text" id="it116-s3-liab13" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr><td>14.</td><td>Non-institutional liabilities (Trade & other payables)</td><td><input type="text" id="it116-s3-liab14" class="it116-cell-input text-right font-mono"></td></tr>
                            <tr class="font-bold bg-slate-100"><td>15.</td><td>Total capital and liabilities (12+13+14)</td><td><input type="text" id="it116-s3-liab15" class="it116-cell-input text-right font-mono font-bold"></td></tr>
                        </table>
                    </div>
                    <div class="text-center font-bold pt-1">5</div>
                </section>

                <!-- IT-116 PAGE 6: LIST OF DOCUMENTS & DECLARATION -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="flex justify-between items-center gap-4 mb-4">
                            <div class="flex items-center w-1/2">
                                <span class="font-bold whitespace-nowrap">Name of Taxpayer:</span>
                                <input type="text" id="it116-p6-name" class="it116-line-input font-semibold">
                            </div>
                            <div class="flex items-center w-1/2 justify-end">
                                <span class="font-bold mr-1">TIN:</span>
                                <div class="it116-tin-container" data-tin-group="p6"></div>
                            </div>
                        </div>

                        <div class="border-2 border-black min-h-[125mm] flex flex-col p-3">
                            <div class="text-center font-bold pb-2 border-b border-black mb-3">
                                List of documents submitted along with this return
                            </div>
                            <textarea id="it116-p6-docs" class="it116-cell-input flex-1 resize-none leading-relaxed"></textarea>
                        </div>

                        <div class="text-center font-bold mt-6 underline text-[14px]">Declaration</div>

                        <div class="mt-3 space-y-2 leading-relaxed">
                            <div class="flex items-end">
                                <span>I,</span>
                                <input type="text" id="it116-p6-declname" class="it116-line-input mx-2 font-semibold">
                                <span>hereby declare that</span>
                            </div>
                            <div>That,</div>
                            <div>(a) The information provided in this return and statements and attached proofs are complete and correct to the best of my knowledge and belief;</div>
                            <div>(b) Signed and authenticated by me as partner/ chief officer/ principal.</div>
                        </div>

                        <div class="flex justify-between mt-10">
                            <div class="space-y-4">
                                <div class="flex items-center">
                                    <span>Place:</span>
                                    <input type="text" value="Dhaka, Bangladesh" class="it116-line-input w-40">
                                </div>
                                <div class="flex items-center">
                                    <span>Date:</span>
                                    <input type="text" id="it116-p6-date" class="it116-line-input w-40">
                                </div>
                            </div>
                            <div class="text-center w-64">
                                <div class="border-b border-black w-full h-6 mb-1"></div>
                                <div>Signature</div>
                                <div class="mt-2 font-bold" id="it116-p6-sig-name">(Name of Signatory)</div>
                                <div class="text-[11px]" id="it116-p6-sig-role">Designation and Seal</div>
                            </div>
                        </div>
                    </div>
                    <div class="text-center font-bold pt-2">6</div>
                </section>

                <!-- IT-116 PAGE 7: INSTRUCTIONS -->
                <section class="a4-sheet it116-page flex flex-col justify-between">
                    <div>
                        <div class="text-center font-bold mb-2.5 text-[14px]">Instructions to be followed while filling out the return form:</div>
                        <div class="border-2 border-black p-4 text-[12px] leading-relaxed space-y-2">
                            <p>1. This income tax return must be signed and authenticated by the individual taxpayer or the specified person as described in the law.</p>
                            <p>2. Attach if necessary:</p>
                            <div class="pl-5 space-y-1.5">
                                <p>(a) In case of bank interest, bank statement, in case of interest on savings certificates, certificate from the bank providing the interest, in case of income from house property, rent agreement, receipt of municipal tax and land revenue, in case of interest on housing loan, certificate/statement from the lending institution, if insurance premium is paid, receipt of premium payment, in case of professional income, statement of income as per the relevant rules, if there is a share of income of a partnership firm, copy of the assessment order of the partnership firm/ income-expenditure account and balance sheet, if there is capital gain, proof, if there is dividend income, certificate of receipt of dividend, if there is income from other sources, details thereof, and if there is investment in savings certificates, L.I.P, DPS, Zakat, purchase of stocks/shares, etc., proof;</p>
                                <p>(b) If there is business income, income-expenditure account statement, production account, trading account, profit and loss account and balance sheet;</p>
                                <p>(c) Depreciation statement containing depreciation claim as per the relevant schedule;</p>
                                <p>(d) Computation of income as per the Income Tax Act, 2023.</p>
                            </div>
                            <p>3. Attach separate statements:</p>
                            <div class="pl-5 space-y-1.5">
                                <p>(a) If there is any income in the name of the taxpayer's wife or husband (if not a taxpayer), minor child and dependent;</p>
                                <p>(b) Details of tax exempted and tax free income as per the relevant schedule and SRO;</p>
                                <p>(c) Income exempted from tax as declared under Part 1 of the Sixth Schedule of the Income Tax Act, 2023;</p>
                                <p>(d) Income exempted from tax as declared under the Income Tax Act, 2023.</p>
                            </div>
                            <p>4. The submitted documents must be signed by the taxpayer or the authorized representative of the taxpayer.</p>
                            <p>5. Signature of the taxpayer or his authorized representative.</p>
                            <p>6. If space is insufficient, separate sheets may be used if necessary.</p>
                        </div>
                    </div>
                    <div class="text-center font-bold pt-2">7</div>
                </section>

                <!-- IT-11 COLA (2023): ACKNOWLEDGEMENT RECEIPT / CERTIFICATE OF RETURN OF INCOME -->
                <section class="a4-sheet it11cola-page flex flex-col justify-between" style="padding: 18mm 22mm;">
                    <div>
                        <!-- Top Row -->
                        <div class="flex justify-between text-[12px] font-bold mb-4">
                            <div>|</div>
                            <div style="color: #e31837;">IT-11 Cola (2023)</div>
                        </div>

                        <!-- Main Header -->
                        <div class="text-center it11cola-serif mb-5">
                            <img src="https://upload.wikimedia.org/wikipedia/commons/8/84/Government_Seal_of_Bangladesh.svg"
                                 onerror="this.onerror=null;this.src='https://placehold.co/65x65/ffffff/0f172a?text=GOV+BD';"
                                 alt="Bangladesh Government Seal"
                                 class="w-[62px] h-auto mx-auto mb-2.5">
                            <div class="text-[14px] font-normal leading-snug">Government of the People's Republic of Bangladesh</div>
                            <div class="text-[14px] font-bold leading-snug">National Board of Revenue</div>
                            <div class="text-[13px] font-bold leading-snug">(Income Tax Office)</div>
                        </div>

                        <!-- Document Title -->
                        <div class="text-center text-[15px] font-bold underline my-6 it11cola-serif">
                            Acknowledgement Receipt/Certificate of Return of Income
                        </div>

                        <!-- Form Section -->
                        <div class="space-y-3.5 text-[13px] font-bold">
                            <div class="flex items-end">
                                <label class="w-[220px] whitespace-nowrap">Assessment Year</label>
                                <span class="mx-2.5">:</span>
                                <input type="text" id="it11cola-taxyear" class="it11cola-input font-normal">
                            </div>

                            <div class="flex items-end">
                                <label class="w-[220px] whitespace-nowrap">Name of the Taxpayer</label>
                                <span class="mx-2.5">:</span>
                                <input type="text" id="it11cola-name" class="it11cola-input font-semibold">
                            </div>

                            <div class="flex items-end">
                                <label class="w-[220px] whitespace-nowrap">NID / Passport N° (if No NID)</label>
                                <span class="mx-2.5">:</span>
                                <input type="text" id="it11cola-nid" class="it11cola-input font-normal" placeholder="N/A (Registered Partnership Firm)">
                            </div>

                            <!-- TIN Input -->
                            <div class="flex items-center pt-1">
                                <label class="w-[220px] whitespace-nowrap">TIN</label>
                                <span class="mx-2.5">:</span>
                                <div class="it116-tin-container !ml-0" data-tin-group="cola"></div>
                            </div>
                        </div>

                        <!-- Circle & Taxes Zone -->
                        <div class="flex justify-between mt-5 mb-6 text-[13px] font-bold">
                            <div class="flex items-end w-[46%]">
                                <label class="whitespace-nowrap mr-2">Circle:</label>
                                <input type="text" id="it11cola-circle" class="it11cola-input font-normal">
                            </div>
                            <div class="flex items-end w-[48%]">
                                <label class="whitespace-nowrap mr-2">Taxes Zone:</label>
                                <input type="text" id="it11cola-zone" class="it11cola-input font-normal">
                            </div>
                        </div>

                        <!-- Financial Amounts -->
                        <div class="mb-6 text-[13px] font-bold space-y-2.5">
                            <div class="flex items-end">
                                <label class="w-[180px] whitespace-nowrap">Total Income Shown</label>
                                <span class="mx-2.5">:</span>
                                <span class="ml-4 mr-1.5">Tk.</span>
                                <input type="text" id="it11cola-income" class="it11cola-input !flex-grow-0 w-52 font-mono font-bold">
                            </div>
                            <div class="flex items-end">
                                <label class="w-[180px] whitespace-nowrap">Total Tax Paid</label>
                                <span class="mx-2.5">:</span>
                                <span class="ml-4 mr-1.5">Tk.</span>
                                <input type="text" id="it11cola-taxpaid" class="it11cola-input !flex-grow-0 w-52 font-mono font-bold">
                            </div>
                        </div>

                        <!-- Register Table -->
                        <div class="mb-8">
                            <table class="it11cola-reg-table">
                                <tr>
                                    <td>Serial No. of Return Register</td>
                                    <td><input type="text" id="it11cola-serial"></td>
                                </tr>
                                <tr>
                                    <td>Volume No. of Return Register</td>
                                    <td><input type="text" id="it11cola-volume"></td>
                                </tr>
                                <tr>
                                    <td>Date of Return Submission</td>
                                    <td><input type="date" id="it11cola-subdate"></td>
                                </tr>
                            </table>
                        </div>
                    </div>

                    <!-- Signatures & Bottom Footer -->
                    <div>
                        <div class="flex justify-between mb-10 text-[12px] font-bold">
                            <div class="w-[240px] text-left">
                                <div class="h-10"></div>
                                Seal of Tax Office
                            </div>
                            <div class="w-[260px] text-left">
                                <div class="h-10"></div>
                                Signature and Seal of the Official Receiving the Return
                            </div>
                        </div>

                        <div class="text-center text-[11px] text-slate-600 pt-2 border-t border-slate-200">
                            Please Visit: <a href="https://etaxnbr.gov.bd/" target="_blank" rel="noopener noreferrer" class="text-[#0066cc] underline">"https://etaxnbr.gov.bd/"</a> website to get Income Tax Certificate in Online
                        </div>
                    </div>
                </section>

            </div>

        </main>
    </div>
`;
