# SS-DYNAMIC-FIX-PHONE-LAPTOP-REPAIR-SERVICE-PROFESSIONAL-EXPERT-REPAIR-AND-SERVICE-MANAGEMENT-SYSTEM
SS DYNAMIC FIX PHONE &amp; LAPTOP REPAIR SERVICE PROFESSIONAL EXPERT REPAIR AND SERVICE MANAGEMENT SYSTEM
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SS DYNAMIC FIX | Expert Repair Management Lab</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel, jsPDF for PDF -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.29/jspdf.plugin.autotable.min.js"></script>
    
    <style>
        body {
            background-color: #0c0c0c;
            background-image: radial-gradient(rgba(239, 68, 68, 0.1) 1px, transparent 0);
            background-size: 24px 24px;
        }
        .glow-red {
            box-shadow: 0 0 15px rgba(239, 68, 68, 0.3);
        }
    </style>
</head>
<body class="text-gray-100 font-sans min-h-screen">

    <!-- HEADER / NAVIGATION -->
    <header class="border-b-2 border-red-600 bg-black sticky top-0 z-50 glow-red">
        <div class="max-w-7xl mx-auto px-4 py-4 flex flex-col sm:flex-row justify-between items-center gap-4">
            <div class="flex items-center gap-3">
                <div class="bg-red-600 p-2.5 rounded-lg text-black font-black text-xl animate-pulse">
                    <i class="fa-solid fa-screwdriver-wrench"></i>
                </div>
                <div>
                    <h1 class="text-2xl font-black tracking-wider text-white">SS DYNAMIC FIX</h1>
                    <p class="text-xs text-red-500 font-bold tracking-widest uppercase">Phone & Laptop Repairing Master Expert Lab</p>
                </div>
            </div>
            <!-- Sync Status Indicator -->
            <div class="flex items-center gap-3 bg-neutral-900 border border-neutral-800 px-4 py-2 rounded-full">
                <span class="relative flex h-3 w-3">
                    <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-green-400 opacity-75"></span>
                    <span class="relative inline-flex rounded-full h-3 w-3 bg-green-500"></span>
                </span>
                <span id="syncStatus" class="text-xs font-semibold text-gray-400">Cloud Data Synced</span>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 py-8 grid grid-cols-1 lg:grid-cols-3 gap-8">
        
        <!-- COLUMN 1 & 2: INPUT ENGINE -->
        <section class="lg:col-span-1 bg-neutral-950 border border-neutral-800 p-6 rounded-xl glow-red h-fit">
            <div class="flex items-center gap-2 mb-6 border-b border-neutral-800 pb-3">
                <i class="fa-solid fa-square-plus text-red-500 text-xl"></i>
                <h2 class="text-lg font-bold uppercase tracking-wider text-red-500">Log New Job Entry</h2>
            </div>

            <form id="repairForm" class="space-y-4" onsubmit="handleFormSubmit(event)">
                <!-- 1. DATE -->
                <div>
                    <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Repair Date</label>
                    <input type="date" id="repairDate" required class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 text-white focus:outline-none focus:border-red-500">
                </div>

                <!-- 2. BRAND & 3. MODEL -->
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Brand</label>
                        <input type="text" id="brand" placeholder="e.g., Apple, ASUS" required class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 text-white focus:outline-none focus:border-red-500">
                    </div>
                    <div>
                        <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Model Name</label>
                        <input type="text" id="modelName" placeholder="e.g., iPhone 13, ROG 5" required class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 text-white focus:outline-none focus:border-red-500">
                    </div>
                </div>

                <!-- 4. CODE MODEL -->
                <div>
                    <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Model Code / Serial</label>
                    <input type="text" id="modelCode" placeholder="e.g., A2633, G513" required class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 text-white focus:outline-none focus:border-red-500">
                </div>

                <!-- 5. DAMAGE -->
                <div>
                    <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Damage / Fault Symptoms</label>
                    <textarea id="damage" rows="2" placeholder="e.g., Shorted capacitor on main VDD rail, no power" required class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 text-white focus:outline-none focus:border-red-500"></textarea>
                </div>

                <!-- 6. HOW TO FIX -->
                <div>
                    <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Solution / How to Fix</label>
                    <textarea id="howToFix" rows="2" placeholder="e.g., Replaced capacitor, micro-soldering required" required class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 text-white focus:outline-none focus:border-red-500"></textarea>
                </div>

                <!-- FINANCIAL CALCULATOR: 7, 8 -->
                <div class="border-t border-neutral-800 pt-4 mt-2">
                    <h3 class="text-xs font-black uppercase text-red-500 tracking-wider mb-3"><i class="fa-solid fa-calculator mr-1"></i> Financial Calculations</h3>
                    <div class="grid grid-cols-2 gap-3 mb-3">
                        <div>
                            <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Cost (Parts/Wholesale)</label>
                            <div class="relative">
                                <span class="absolute left-3 top-2.5 text-gray-500 text-sm">RM</span>
                                <input type="number" id="costRepair" step="0.01" placeholder="0.00" required oninput="calculateMetrics()" class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 pl-10 text-white focus:outline-none focus:border-red-500">
                            </div>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-gray-400 mb-1">Charge to Customer</label>
                            <div class="relative">
                                <span class="absolute left-3 top-2.5 text-gray-500 text-sm">RM</span>
                                <input type="number" id="chargeCustomer" step="0.01" placeholder="0.00" required oninput="calculateMetrics()" class="w-full bg-neutral-900 border border-neutral-700 rounded p-2.5 pl-10 text-white focus:outline-none focus:border-red-500">
                            </div>
                        </div>
                    </div>

                    <!-- Auto Margin Display -->
                    <div class="grid grid-cols-2 gap-3 bg-neutral-900 p-3 rounded border border-neutral-800">
                        <div>
                            <span class="block text-[10px] uppercase font-bold text-gray-500">Net Profit</span>
                            <span id="labelProfit" class="text-base font-black text-green-400">RM 0.00</span>
                        </div>
                        <div>
                            <span class="block text-[10px] uppercase font-bold text-gray-500">Profit Margin</span>
                            <span id="labelMargin" class="text-base font-black text-red-400">0%</span>
                        </div>
                    </div>
                </div>

                <button type="submit" class="w-full bg-red-600 hover:bg-red-700 text-white font-bold uppercase py-3 rounded tracking-wider transition-colors mt-2">
                    <i class="fa-solid fa-floppy-disk mr-2"></i> Save Repair Job
                </button>
            </form>
        </section>

        <!-- COLUMN 3: LIVE DATABASE & EXPORT CENTER -->
        <section class="lg:col-span-2 flex flex-col gap-6">
            
            <!-- EXPORT ENGINE CONTROLS (10) -->
            <div class="bg-neutral-950 border border-neutral-800 p-4 rounded-xl flex flex-wrap items-center justify-between gap-4">
                <div>
                    <h3 class="font-bold text-sm text-white uppercase tracking-wider">Export Master Management Center</h3>
                    <p class="text-xs text-gray-400">Instantly generate files from active sync data</p>
                </div>
                <div class="flex flex-wrap gap-2">
                    <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold uppercase px-3 py-2 rounded flex items-center gap-1.5 transition-colors">
                        <i class="fa-solid fa-file-excel"></i> Sheet
                    </button>
                    <button onclick="exportToDoc()" class="bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold uppercase px-3 py-2 rounded flex items-center gap-1.5 transition-colors">
                        <i class="fa-solid fa-file-word"></i> Word
                    </button>
                    <button onclick="exportToPDF()" class="bg-red-600 hover:bg-red-700 text-white text-xs font-bold uppercase px-3 py-2 rounded flex items-center gap-1.5 transition-colors">
                        <i class="fa-solid fa-file-pdf"></i> PDF
                    </button>
                </div>
            </div>

            <!-- LIVE METRICS BAR -->
            <div class="grid grid-cols-3 gap-4">
                <div class="bg-neutral-950 border border-neutral-800 p-4 rounded-xl">
                    <span class="text-xs font-bold uppercase text-gray-400">Total Tasks</span>
                    <div id="statTotal" class="text-2xl font-black text-white mt-1">0</div>
                </div>
                <div class="bg-neutral-950 border border-neutral-800 p-4 rounded-xl">
                    <span class="text-xs font-bold uppercase text-gray-400">Gross Revenue</span>
                    <div id="statRevenue" class="text-2xl font-black text-red-500 mt-1">RM 0.00</div>
                </div>
                <div class="bg-neutral-950 border border-neutral-800 p-4 rounded-xl">
                    <span class="text-xs font-bold uppercase text-gray-400">Total Profit</span>
                    <div id="statProfit" class="text-2xl font-black text-green-400 mt-1">RM 0.00</div>
                </div>
            </div>

            <!-- THE LIVE RECORD TABLE -->
            <div class="bg-neutral-950 border border-neutral-800 rounded-xl overflow-hidden flex-1 flex flex-col">
                <div class="p-4 border-b border-neutral-800 flex justify-between items-center bg-black">
                    <span class="text-xs font-black uppercase tracking-widest text-red-500 flex items-center gap-2">
                        <i class="fa-solid fa-database"></i> Live Database Store
                    </span>
                    <button onclick="clearDatabase()" class="text-[10px] text-gray-500 hover:text-red-400 uppercase font-bold transition-colors">
                        Clear All Local Logs
                    </button>
                </div>

                <div class="overflow-x-auto flex-1">
                    <table class="w-full text-left text-xs text-gray-300 border-collapse">
                        <thead>
                            <tr class="bg-neutral-900 text-gray-400 uppercase border-b border-neutral-800">
                                <th class="p-3 font-bold">Date</th>
                                <th class="p-3 font-bold">Device / Code</th>
                                <th class="p-3 font-bold">Diagnosis & Repair Action</th>
                                <th class="p-3 font-bold text-right">Cost</th>
                                <th class="p-3 font-bold text-right">Charged</th>
                                <th class="p-3 font-bold text-right">Profit</th>
                            </tr>
                        </thead>
                        <tbody id="databaseBody" class="divide-y divide-neutral-900">
                            <!-- Dynamic Content Injected Here -->
                        </tbody>
                    </table>
                    <div id="emptyState" class="p-8 text-center text-gray-500 flex flex-col items-center justify-center gap-2">
                        <i class="fa-solid fa-box-open text-3xl"></i>
                        <p class="font-medium">No repair records found. Log your first case entry.</p>
                    </div>
                </div>
            </div>

        </section>
    </main>

    <!-- FOOTER -->
    <footer class="max-w-7xl mx-auto px-4 py-8 mt-12 border-t border-neutral-900 text-center text-xs text-gray-600">
        &copy; 2026 SS DYNAMIC FIX &bull; High Power Enterprise System Architecture. All Rights Reserved.
    </footer>

    <!-- LOGIC CORE ENGINE -->
    <script>
        // Set standard layout initialization date to current system local date
        document.getElementById('repairDate').valueAsDate = new Date();

        // Local State Array
        let repairDB = JSON.parse(localStorage.getItem('ss_dynamic_fix_db')) || [];

        // Financial Math Processing
        function calculateMetrics() {
            const cost = parseFloat(document.getElementById('costRepair').value) || 0;
            const charge = parseFloat(document.getElementById('chargeCustomer').value) || 0;
            
            const profit = charge - cost;
            const margin = charge > 0 ? (profit / charge) * 100 : 0;

            document.getElementById('labelProfit').innerText = `RM ${profit.toFixed(2)}`;
            document.getElementById('labelProfit').className = profit >= 0 ? "text-base font-black text-green-400" : "text-base font-black text-red-500";
            document.getElementById('labelMargin').innerText = `${margin.toFixed(0)}%`;
        }

        // Form Submit & Storage Pipeline
        function handleFormSubmit(e) {
            e.preventDefault();
            
            const cost = parseFloat(document.getElementById('costRepair').value) || 0;
            const charge = parseFloat(document.getElementById('chargeCustomer').value) || 0;
            const profit = charge - cost;
            const margin = charge > 0 ? ((profit / charge) * 100).toFixed(1) : 0;

            const record = {
                id: Date.now(),
                date: document.getElementById('repairDate').value,
                brand: document.getElementById('brand').value,
                modelName: document.getElementById('modelName').value,
                modelCode: document.getElementById('modelCode').value,
                damage: document.getElementById('damage').value,
                howToFix: document.getElementById('howToFix').value,
                costRepair: cost,
                chargeCustomer: charge,
                profit: profit,
                margin: margin
            };

            repairDB.unshift(record);
            saveAndSync();
            document.getElementById('repairForm').reset();
            document.getElementById('repairDate').valueAsDate = new Date();
            calculateMetrics();
        }

        // 9. AUTO SYNC PIPELINE
        function saveAndSync() {
            localStorage.setItem('ss_dynamic_fix_db', JSON.stringify(repairDB));
            
            // Visual feedback loop showing Auto Sync Action
            const statusEl = document.getElementById('syncStatus');
            statusEl.innerText = "Syncing Cloud Data...";
            statusEl.className = "text-xs font-semibold text-yellow-500";

            setTimeout(() => {
                statusEl.innerText = "Cloud Data Synced";
                statusEl.className = "text-xs font-semibold text-gray-400";
            }, 800);

            renderDashboard();
        }

        // UI Core Renderer
        function renderDashboard() {
            const tbody = document.getElementById('databaseBody');
            const emptyState = document.getElementById('emptyState');
            
            tbody.innerHTML = '';
            
            if (repairDB.length === 0) {
                emptyState.classList.remove('hidden');
                document.getElementById('statTotal').innerText = '0';
                document.getElementById('statRevenue').innerText = 'RM 0.00';
                document.getElementById('statProfit').innerText = 'RM 0.00';
                return;
            }
            
            emptyState.classList.add('hidden');
            
            let totalRevenue = 0;
            let totalProfit = 0;

            repairDB.forEach(item => {
                totalRevenue += item.chargeCustomer;
                totalProfit += item.profit;

                const tr = document.createElement('tr');
                tr.className = "hover:bg-neutral-900/50 transition-colors border-b border-neutral-900";
                tr.innerHTML = `
                    <td class="p-3 text-gray-400 font-mono whitespace-nowrap">${item.date}</td>
                    <td class="p-3">
                        <span class="block font-black text-white">${item.brand} ${item.modelName}</span>
                        <span class="text-[10px] text-red-400 font-mono font-semibold">${item.modelCode}</span>
                    </td>
                    <td class="p-3 max-w-xs break-words">
                        <div class="text-gray-400"><strong class="text-gray-500">Fault:</strong> ${item.damage}</div>
                        <div class="text-gray-300 mt-0.5"><strong class="text-red-500">Fix:</strong> ${item.howToFix}</div>
                    </td>
                    <td class="p-3 text-right font-mono text-gray-400">RM ${item.costRepair.toFixed(2)}</td>
                    <td class="p-3 text-right font-mono font-bold text-white">RM ${item.chargeCustomer.toFixed(2)}</td>
                    <td class="p-3 text-right font-mono font-black ${item.profit >= 0 ? 'text-green-400' : 'text-red-500'}">
                        RM ${item.profit.toFixed(2)}
                        <span class="block text-[9px] text-gray-500 font-normal">${item.margin}% margin</span>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            // Set Global Dashboard Counters
            document.getElementById('statTotal').innerText = repairDB.length;
            document.getElementById('statRevenue').innerText = `RM ${totalRevenue.toFixed(2)}`;
            document.getElementById('statProfit').innerText = `RM ${totalProfit.toFixed(2)}`;
        }

        // Database reset functional routine
        function clearDatabase() {
            if(confirm("Are you sure you want to completely wipe all dynamic repair database records? This cannot be undone.")) {
                repairDB = [];
                saveAndSync();
            }
        }

        /* ==========================================================================
           10. MASTER EXPORT ARCHITECTURE ENGINE
           ========================================================================== */
        
        // EXPORT TO EXCEL SHEET
        function exportToExcel() {
            if (repairDB.length === 0) return alert("No active data entries to export.");
            
            const wsData = repairDB.map(r => ({
                "Date": r.date,
                "Brand": r.brand,
                "Model Name": r.modelName,
                "Model Code": r.modelCode,
                "Damage Symptom": r.damage,
                "Repair Method": r.howToFix,
                "Cost (RM)": r.costRepair,
                "Charged Customer (RM)": r.chargeCustomer,
                "Profit Margins (RM)": r.profit,
                "Margin Rate (%)": r.margin + "%"
            }));

            const ws = XLSX.utils.json_to_sheet(wsData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Repair Logs Matrix");
            XLSX.writeFile(wb, `SS_DYNAMIC_FIX_MASTER_SHEET.xlsx`);
        }

        // EXPORT TO WORD DOCUMENT (.doc)
        function exportToDoc() {
            if (repairDB.length === 0) return alert("No active data entries to export.");

            let htmlString = `
                <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
                <head><title>Repair Report Ledger</title><style>
                    body { font-family: Arial, sans-serif; }
                    h1 { color: #cc0000; font-size: 20pt; border-bottom: 2px solid #000; padding-bottom: 5px; }
                    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                    th, td { border: 1px solid #ddd; padding: 8px; font-size: 10pt; text-align: left; }
                    th { background-color: #111; color: #fff; font-weight: bold; }
                </style></head>
                <body>
                    <h1>SS DYNAMIC FIX - MASTER REPORT LEDGER</h1>
                    <p>Generated Document Timestamp: ${new Date().toLocaleString()}</p>
                    <table>
                        <tr>
                            <th>Date</th><th>Device Info</th><th>Damage & Diagnosis</th><th>Cost</th><th>Charged</th><th>Net Profit</th>
                        </tr>`;

            repairDB.forEach(r => {
                htmlString += `
                    <tr>
                        <td>${r.date}</td>
                        <td><b>${r.brand} ${r.modelName}</b><br><small>${r.modelCode}</small></td>
                        <td><b>Fault:</b> ${r.damage}<br><b>Fix:</b> ${r.howToFix}</td>
                        <td>RM ${r.costRepair.toFixed(2)}</td>
                        <td>RM ${r.chargeCustomer.toFixed(2)}</td>
                        <td>RM ${r.profit.toFixed(2)} (${r.margin}%)</td>
                    </tr>`;
            });

            htmlString += `</table></body></html>`;

            const blob = new Blob(['\ufeff' + htmlString], { type: 'application/msword' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'SS_DYNAMIC_FIX_WORD_REPORT.doc';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }

        // EXPORT TO PDF
        function exportToPDF() {
            if (repairDB.length === 0) return alert("No active data entries to export.");
            
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4'); // Landscape layout configuration

            doc.setFillColor(12, 12, 12);
            doc.rect(0, 0, 297, 210, 'F');

            doc.setFont("Helvetica", "bold");
            doc.setFontSize(22);
            doc.setTextColor(239, 68, 68); // Brand Red
            doc.text("SS DYNAMIC FIX", 14, 20);

            doc.setFontSize(9);
            doc.setTextColor(150, 150, 150);
            doc.text("PHONE & LAPTOP REPAIRING MASTER EXPERT LAB  |  EXECUTIVE MASTER LEDGER", 14, 26);
            doc.text(`Generated: ${new Date().toLocaleString()}`, 230, 26);

            const headers = [["Date", "Device Identity", "Diagnosis / Mitigation Action", "Cost", "Price Charged", "Net Profit Margin"]];
            
            const dataRows = repairDB.map(r => [
                r.date,
                `${r.brand} ${r.modelName}\n(${r.modelCode})`,
                `Symptom: ${r.damage}\nAction: ${r.howToFix}`,
                `RM ${r.costRepair.toFixed(2)}`,
                `RM ${r.chargeCustomer.toFixed(2)}`,
                `RM ${r.profit.toFixed(2)}\n(${r.margin}%)`
            ]);

            doc.autoTable({
                startY: 32,
                head: headers,
                body: dataRows,
                theme: 'grid',
                headStyles: { fillColor: [185, 28, 28], textColor: [255, 255, 255], fontStyle: 'bold' },
                bodyStyles: { fillColor: [24, 24, 24], textColor: [220, 220, 220], fontSize: 9 },
                alternateRowStyles: { fillColor: [18, 18, 18] },
                gridLineColor: [40, 40, 40],
                styles: { overflow: 'linebreak', cellPadding: 4 }
            });

            doc.save("SS_DYNAMIC_FIX_LAB_REPORT.pdf");
        }

        // Run application on start
        renderDashboard();
    </script>
</body>
