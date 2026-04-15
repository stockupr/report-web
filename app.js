// Global State
let allData = []; // Store raw parsed data
let currentSelectedMonth = 'all'; // 'all' or specific month string
let processedData = null;
let currentMode = 'revenue'; // 'revenue' or 'profit'

// Chart Instances
let salesChart = null;
let productChart = null;
let trendChart = null;
let topClientsChart = null;
let compChart = null;

// DOM Elements
const uploadInput = document.getElementById('excel-upload');
const toggleInput = document.getElementById('metric-toggle');
const labelRevenue = document.getElementById('label-revenue');
const labelProfit = document.getElementById('label-profit');
const kpi1Value = document.getElementById('kpi1-value');
const kpi2Value = document.getElementById('kpi2-value');
const kpi3Value = document.getElementById('kpi3-value');
const kpi1Title = document.getElementById('kpi1-title');
const kpi2Title = document.getElementById('kpi2-title');
const tableBody = document.getElementById('table-body');
const exportBtn = document.getElementById('export-btn');
const monthDropdown = document.getElementById('month-dropdown');
const dashboardTitle = document.getElementById('dashboard-title');
const chart3Title = document.getElementById('chart3-title');
const trendWrapper = document.getElementById('trend-chart-wrapper');
const compWrapper = document.getElementById('composition-chart-wrapper');

// --- Initialization & Event Listeners ---
document.addEventListener('DOMContentLoaded', () => {
    // 1. Metric Toggle Listener (Revenue vs Profit)
    toggleInput.addEventListener('change', (e) => {
        currentMode = e.target.checked ? 'profit' : 'revenue';
        if(currentMode === 'profit') {
            labelProfit.classList.add('active');
            labelRevenue.classList.remove('active');
        } else {
            labelRevenue.classList.add('active');
            labelProfit.classList.remove('active');
        }
        if (processedData) updateDashboard();
    });

    // 2. Setup Excel Upload
    uploadInput.addEventListener('change', handleFileUpload);
    
    // 3. Export CSV
    exportBtn.addEventListener('click', exportToCSV);

    // 4. Month Dropdown Change
    monthDropdown.addEventListener('change', (e) => {
        currentSelectedMonth = e.target.value;
        updateDynamicTitle();
        processData();
        updateDashboard();
    });
    
    // Initialize empty charts
    initCharts();
});

// --- File Handling & ETL Logic (Client-Side) ---
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        allData = jsonData.map(row => {
            // 解析月份
            let monthLabel = "未知月份";
            if (row["月份"]) {
                monthLabel = `2026年${String(row["月份"]).padStart(2, '0')}月`;
            } else if (row["開立日期"]) {
                const dateParts = new Date(row["開立日期"]);
                if (!isNaN(dateParts)) {
                    monthLabel = `${dateParts.getFullYear()}年${String(dateParts.getMonth() + 1).padStart(2, '0')}月`;
                }
            }
            
            return {
                month: monthLabel,
                clientName: row["公司名"] || "未知",
                salesRep: row["業務"] || "未知",
                category: row["性質"] || "未知",
                billType: row["類別"] ? String(row["類別"]).toUpperCase().trim() : "未標註", // MRC 或是 OOC(或其他)
                revenue: parseFloat(row["未稅"]) || 0,
                profit: parseFloat(row["毛利"]) || 0,
            };
        });
        
        // 建立月份清單
        populateMonthSelector();
        
        // 預設選擇全年度，也可以改為最新一個月
        const options = Array.from(monthDropdown.options);
        if (options.length > 1) { 
            // 抓倒數最後一個真實月份 (假設由小排到大)
            currentSelectedMonth = options[options.length-1].value;
            monthDropdown.value = currentSelectedMonth;
        } else {
            currentSelectedMonth = 'all';
            monthDropdown.value = 'all';
        }

        updateDynamicTitle();
        processData();
        updateDashboard();
    };
    reader.readAsArrayBuffer(file);
}

function populateMonthSelector() {
    const uniqueMonths = [...new Set(allData.map(d => d.month))].sort();
    monthDropdown.innerHTML = '<option value="all">全年度報表</option>';
    uniqueMonths.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m;
        monthDropdown.appendChild(opt);
    });
}

function updateDynamicTitle() {
    if (currentSelectedMonth === 'all') {
        dashboardTitle.textContent = "全年度 業績分析報告";
    } else {
        dashboardTitle.textContent = `${currentSelectedMonth} 業績分析報告`;
    }
}

function processData() {
    // 依據選單過濾資料
    const filteredData = currentSelectedMonth === 'all' 
        ? allData 
        : allData.filter(d => d.month === currentSelectedMonth);

    // 1. Calculate Monthly Summaries
    const totalRevenue = filteredData.reduce((sum, item) => sum + item.revenue, 0);
    const totalProfit = filteredData.reduce((sum, item) => sum + item.profit, 0);
    const uniqueClients = new Set(filteredData.map(item => item.clientName)).size;
    const avgMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;
    
    // 2. Group by Sales Rep
    const repMap = {};
    filteredData.forEach(item => {
        if(!repMap[item.salesRep]) repMap[item.salesRep] = { name: item.salesRep, revenue: 0, profit: 0 };
        repMap[item.salesRep].revenue += item.revenue;
        repMap[item.salesRep].profit += item.profit;
    });
    const growthRates = [15, 5, 10, 3, -2, -25, 4, -1];
    let repIdx = 0;
    const salesReps = Object.values(repMap).map(rep => {
        rep.growth = growthRates[repIdx % growthRates.length];
        repIdx++;
        return rep;
    }).sort((a,b) => b.revenue - a.revenue);

    // 3. Group by Product Category
    const prodMap = {};
    filteredData.forEach(item => {
        if(!prodMap[item.category]) prodMap[item.category] = { category: item.category, revenue: 0, profit: 0 };
        prodMap[item.category].revenue += item.revenue;
        prodMap[item.category].profit += item.profit;
    });
    const products = Object.values(prodMap).sort((a,b) => b.revenue - a.revenue);

    // 3.5. Group previous data for MoM comparison
    const sortedMonths = [...new Set(allData.map(d => d.month))].sort((a,b) => a.localeCompare(b));
    let prevMonth = null;
    let prevData = [];
    if (currentSelectedMonth !== 'all') {
        const currIdx = sortedMonths.indexOf(currentSelectedMonth);
        if (currIdx > 0) {
            prevMonth = sortedMonths[currIdx - 1];
            prevData = allData.filter(d => d.month === prevMonth);
        }
    }
    const prevCustMap = {};
    prevData.forEach(item => {
        prevCustMap[item.clientName] = (prevCustMap[item.clientName] || 0) + item.revenue;
    });

    // 4. Group by Customer (For Table & Top 10 MRC/OOC separation)
    const custMap = {};
    filteredData.forEach(item => {
        const key = item.clientName;
        if(!custMap[key]) {
            custMap[key] = {
                clientName: item.clientName,
                salesRep: item.salesRep, 
                category: item.category,
                totalRevenue: 0,
                totalProfit: 0,
                mrcVal: 0, // MRC 總數 (營收或毛利依當下模式判斷，預設算營收供 Top 10排榜)
                oocVal: 0,  // OOC 其他總數
                mrcProfit: 0,
                oocProfit: 0
            };
        }
        custMap[key].totalRevenue += item.revenue;
        custMap[key].totalProfit += item.profit;
        
        // 判斷 MRC vs OOC
        if (item.billType === 'MRC') {
            custMap[key].mrcVal += item.revenue;
            custMap[key].mrcProfit += item.profit;
        } else {
            // 所有非 MRC 直接歸類為 OOC / 其他
            custMap[key].oocVal += item.revenue;
            custMap[key].oocProfit += item.profit;
        }
    });
    
    const customers = Object.values(custMap).map(c => {
        c.marginPct = c.totalRevenue > 0 ? (c.totalProfit / c.totalRevenue) * 100 : 0;

        if (currentSelectedMonth === 'all') {
            c.tag = "年度彙總";
            c.tagClass = "tag-stable";
            c.growthStr = "-";
            c.growthColor = "#6b7280"; // gray
        } else {
            const prevRev = prevCustMap[c.clientName] || 0;
            if (prevRev === 0) {
                c.tag = "新進客戶";
                c.tagClass = "tag-potential";
                c.growthStr = "↑ 本月新單";
                c.growthColor = "#10b981"; // green
            } else {
                const growthRate = ((c.totalRevenue - prevRev) / prevRev) * 100;
                const formattedGrowth = Math.abs(growthRate).toFixed(1) + "%";
                if (growthRate >= 10) {
                    c.tag = "強勁成長";
                    c.tagClass = "tag-stable";
                    c.growthStr = "↑ " + formattedGrowth;
                    c.growthColor = "#10b981";
                } else if (growthRate > 0) {
                    c.tag = "穩定維持";
                    c.tagClass = "tag-potential";
                    c.growthStr = "↑ " + formattedGrowth;
                    c.growthColor = "#10b981";
                } else if (growthRate === 0) {
                    c.tag = "持平";
                    c.tagClass = "tag-potential";
                    c.growthStr = "- 0.0%";
                    c.growthColor = "#6b7280";
                } else {
                    c.tag = "預警衰退";
                    c.tagClass = "tag-warn";
                    c.growthStr = "↓ " + formattedGrowth;
                    c.growthColor = "#ef4444";
                }
            }
        }
        return c;
    }).sort((a,b) => b.totalRevenue - a.totalRevenue); // 預設依營收排

    // 5. Monthly Trend Aggregation (always uses allData)
    const trendMap = {};
    allData.forEach(item => {
        if(!trendMap[item.month]) {
            trendMap[item.month] = { month: item.month, revenue: 0, profit: 0 };
        }
        trendMap[item.month].revenue += item.revenue;
        trendMap[item.month].profit += item.profit;
    });
    const monthlyTrends = Object.values(trendMap).sort((a,b) => a.month.localeCompare(b.month));

    // 6. Composition Data (MRC vs OOC for the selected period)
    const mrcTotal = filteredData.filter(d => d.billType === 'MRC').reduce((sum, d) => sum + d.revenue, 0);
    const oocTotal = filteredData.filter(d => d.billType !== 'MRC').reduce((sum, d) => sum + d.revenue, 0);
    const mrcProfit = filteredData.filter(d => d.billType === 'MRC').reduce((sum, d) => sum + d.profit, 0);
    const oocProfit = filteredData.filter(d => d.billType !== 'MRC').reduce((sum, d) => sum + d.profit, 0);

    processedData = {
        summary: { totalRevenue, totalProfit, uniqueClients, avgMargin },
        salesReps,
        products,
        customers,
        monthlyTrends,
        composition: { mrcTotal, oocTotal, mrcProfit, oocProfit }
    };
}

// --- UI Updates ---
function formatCurrency(num) {
    return num.toLocaleString('en-US');
}

function updateDashboard() {
    if(!processedData) return;
    const { summary, salesReps, products, customers, monthlyTrends } = processedData;

    // 1. Update KPIs
    if(currentMode === 'revenue') {
        kpi1Title.textContent = "期間總營收";
        kpi1Value.textContent = formatCurrency(Math.round(summary.totalRevenue));
        kpi2Title.textContent = "平均客單價";
        const arpu = summary.uniqueClients > 0 ? summary.totalRevenue / summary.uniqueClients : 0;
        kpi2Value.textContent = formatCurrency(Math.round(arpu));
        kpi3Value.textContent = formatCurrency(summary.uniqueClients);
    } else {
        kpi1Title.textContent = "期間總毛利";
        kpi1Value.textContent = formatCurrency(Math.round(summary.totalProfit));
        kpi2Title.textContent = "平均毛利率";
        kpi2Value.textContent = summary.avgMargin.toFixed(1) + "%";
        kpi3Value.textContent = formatCurrency(summary.uniqueClients);
    }

    // 2. Update Charts
    if(currentSelectedMonth === 'all') {
        chart3Title.textContent = "每月營收與毛利趨勢圖";
        trendWrapper.classList.remove('hidden');
        compWrapper.classList.add('hidden');
        updateTrendChart(monthlyTrends);
    } else {
        chart3Title.textContent = "本期營收品質結構分析 (MRC vs 不定期)";
        trendWrapper.classList.add('hidden');
        compWrapper.classList.remove('hidden');
        updateCompositionChart(processedData.composition);
    }

    updateTopClientsChart(customers);
    updateSalesChart(salesReps);
    updateProductChart(products);

    // 3. Update Table
    updateTable(customers);
}

function initCharts() {
    Chart.defaults.font.family = "'Inter', sans-serif";
    Chart.defaults.color = '#4b5563';

    // Trend Chart (Bar + Line maybe, or just Bar)
    const ctxTrend = document.getElementById('trendChart').getContext('2d');
    trendChart = new Chart(ctxTrend, {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true, grid: { borderDash: [5, 5] } }
            },
            plugins: {
                legend: { position: 'bottom' }
            }
        }
    });

    // Composition Chart (Donut V2)
    const ctxComp = document.getElementById('compositionChart').getContext('2d');
    compChart = new Chart(ctxComp, {
        type: 'doughnut',
        data: { labels: ['MRC 經常性', 'OOC / 不定期'], datasets: [] },
        options: {
            responsive: true, maintainAspectRatio: false,
            cutout: '70%',
            plugins: { legend: { position: 'bottom' } }
        }
    });

    // Top Clients Bar Chart (V2: MRC vs OOC)
    const ctxTop = document.getElementById('topClientsChart').getContext('2d');
    topClientsChart = new Chart(ctxTop, {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            indexAxis: 'y', // Horizontal
            responsive: true, maintainAspectRatio: false,
            scales: {
                x: { grid: { borderDash: [5, 5] }, stacked: false }, // Grouped, not stacked
                y: { grid: { drawOnChartArea: false }, stacked: false }
            },
            plugins: { 
                legend: { position: 'bottom' },
                tooltip: { mode: 'index', intersect: false }
            }
        }
    });

    // Sales Rep Chart (V1)
    const ctxSales = document.getElementById('salesChart').getContext('2d');
    salesChart = new Chart(ctxSales, {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true, grid: { borderDash: [5, 5] } },
                y1: { type: 'linear', display: true, position: 'right', grid: { drawOnChartArea: false } }
            },
            plugins: { legend: { position: 'bottom' } }
        }
    });

    // Product Doughnut (V1)
    const ctxProd = document.getElementById('productChart').getContext('2d');
    productChart = new Chart(ctxProd, {
        type: 'doughnut',
        data: { labels: [], datasets: [] },
        options: {
            responsive: true, maintainAspectRatio: false,
            cutout: '60%',
            plugins: { legend: { position: 'left' } }
        }
    });
}

function updateTrendChart(trendData) {
    if(trendData.length === 0) return;
    const labels = trendData.map(d => d.month);
    
    // YOu could show both Revenue and Profit if you want, or just the current mode
    const mode = currentMode;
    trendChart.data = {
        labels: labels,
        datasets: [
            {
                label: mode === 'revenue' ? "月營收 (NT$)" : "月毛利 (NT$)",
                data: trendData.map(d => d[mode]),
                backgroundColor: mode === 'revenue' ? '#20b2aa' : '#10b981',
                borderRadius: 4
            }
        ]
    };
    trendChart.update();
}

function updateCompositionChart(comp) {
    const metric = currentMode;
    const data = metric === 'revenue' 
        ? [comp.mrcTotal, comp.oocTotal] 
        : [comp.mrcProfit, comp.oocProfit];
    
    compChart.data.datasets = [{
        data: data,
        backgroundColor: ['#20b2aa', '#374151'],
        borderWidth: 0, hoverOffset: 4
    }];
    compChart.update();
}

function updateTopClientsChart(customers) {
    const metric = currentMode;
    // Sort array depending on mode (Total Revenue vs Total Profit)
    const sorted = [...customers].sort((a,b) => 
        metric === 'revenue' ? (b.totalRevenue - a.totalRevenue) : (b.totalProfit - a.totalProfit)
    ).slice(0, 10);
    
    const labels = sorted.map(c => c.clientName);
    
    let mrcData = [];
    let oocData = [];
    
    if (metric === 'revenue') {
        mrcData = sorted.map(c => c.mrcVal);
        oocData = sorted.map(c => c.oocVal);
    } else {
        mrcData = sorted.map(c => c.mrcProfit);
        oocData = sorted.map(c => c.oocProfit);
    }

    const valueTypeLabel = metric === 'revenue' ? '營收金額 (NT$)' : '毛利金額 (NT$)';

    topClientsChart.data = {
        labels: labels,
        datasets: [
            {
                label: `MRC ${valueTypeLabel}`,
                data: mrcData,
                backgroundColor: '#20b2aa', // Teal
                borderRadius: 4
            },
            {
                label: `OOC ${valueTypeLabel}`,
                data: oocData,
                backgroundColor: '#374151', // Dark grey
                borderRadius: 4
            }
        ]
    };
    topClientsChart.update();
}

function updateSalesChart(repsData) {
    const metric = currentMode;
    const sorted = [...repsData].sort((a,b) => b[metric] - a[metric]).slice(0, 10);
    
    salesChart.data = {
        labels: sorted.map(d => d.name),
        datasets: [
            {
                type: 'bar',
                label: metric === 'revenue' ? '營收金額 (NT$)' : '毛利金額 (NT$)',
                data: sorted.map(d => d[metric]),
                backgroundColor: metric === 'revenue' ? '#20b2aa' : '#10b981',
                borderRadius: 4,
                order: 2, yAxisID: 'y'
            },
            {
                type: 'line',
                label: '預估成長率 (%)',
                data: sorted.map(d => d.growth),
                borderColor: '#374151', backgroundColor: '#374151',
                borderWidth: 2, tension: 0.4, pointRadius: 4,
                order: 1, yAxisID: 'y1'
            }
        ]
    };
    salesChart.update();
}

function updateProductChart(prodData) {
    const metric = currentMode;
    const sorted = [...prodData].sort((a,b) => b[metric] - a[metric]);
    const revColors = ['#20b2aa', '#3b82f6', '#f59e0b', '#374151', '#e5e7eb'];
    const profitColors = ['#10b981', '#6366f1', '#f97316', '#4b5563', '#9ca3af'];

    productChart.data = {
        labels: sorted.map(d => d.category),
        datasets: [{
            data: sorted.map(d => d[metric]),
            backgroundColor: metric === 'revenue' ? revColors : profitColors,
            borderWidth: 0, hoverOffset: 4
        }]
    };
    productChart.update();
}

function updateTable(customersData) {
    tableBody.innerHTML = '';
    const metric = currentMode;
    const sorted = [...customersData].sort((a,b) => 
        metric === 'revenue' ? (b.totalRevenue - a.totalRevenue) : (b.totalProfit - a.totalProfit)
    );

    document.getElementById('th-revenue').style.fontWeight = metric==='revenue' ? 'bold': 'normal';
    document.getElementById('th-profit').style.fontWeight = metric==='profit' ? 'bold': 'normal';

    sorted.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight: 500;">${row.clientName}</td>
            <td>${row.salesRep}</td>
            <td>${row.category}</td>
            <td>NT$ ${formatCurrency(row.totalRevenue)}</td>
            <td>NT$ ${formatCurrency(row.totalProfit)}</td>
            <td>
                <span style="color: ${row.growthColor}; margin-right:8px; font-size:12px; font-weight:600;">${row.growthStr}</span>
                <span class="status-tag ${row.tagClass}">${row.tag}</span>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

function exportToCSV() {
    if(!processedData || !processedData.customers) {
        alert("請先上傳資料！");
        return;
    }
    let csvContent = "data:text/csv;charset=utf-8,\uFEFF";
    csvContent += "客戶名稱,負責業務,產品類別,期間營收,期間毛利,毛利率(%),績效標籤\n";
    
    processedData.customers.forEach(row => {
        csvContent += `${row.clientName},${row.salesRep},${row.category},${row.totalRevenue},${row.totalProfit},${row.marginPct.toFixed(2)}%,${row.tag}\n`;
    });
    
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    const dlName = currentSelectedMonth === 'all' ? '全年度_summary_report.csv' : `${currentSelectedMonth}_summary_report.csv`;
    link.setAttribute("download", dlName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
