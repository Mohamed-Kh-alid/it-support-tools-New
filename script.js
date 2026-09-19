const GOOGLE_SHEETS_DIRECT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQkUa8gJmOzW8vfd51PF0cXwKk9MjS9TEYHOZYN5TBT8I5xZBd4fxb9rjZj4s3mIVc5sYH0jO1437nx/pub?output=xlsx";

let cachedWorkData = [];
let cachedNightData = [];
let cachedVacationsData = [];
let activeTabName = 'workDistribution';
let currentModalBranches = [];

document.addEventListener('DOMContentLoaded', () => {
    loadOnlineDashboardData();
});

function formatExcelDate(serial) {
    if (!serial) return 'N/A';
    if (isNaN(serial)) return serial;
    const utcDays = Math.floor(serial - 25569);
    const date = new Date(utcDays * 86400 * 1000);
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

async function loadOnlineDashboardData() {
    try {
        const response = await fetch(GOOGLE_SHEETS_DIRECT_URL);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        if (workbook.Sheets['Night_Shift']) {
            cachedNightData = XLSX.utils.sheet_to_json(workbook.Sheets['Night_Shift'], { range: 3 });
            renderNightShiftTable(cachedNightData);
        }

        if (workbook.Sheets['Annual_Vacations']) {
            cachedVacationsData = XLSX.utils.sheet_to_json(workbook.Sheets['Annual_Vacations'], { range: 3 });
            renderVacationsTable(cachedVacationsData);
        }

        if (workbook.Sheets['Work_Distribution']) {
            cachedWorkData = XLSX.utils.sheet_to_json(workbook.Sheets['Work_Distribution'], { range: 3 });
            renderWorkDistributionTable(cachedWorkData);
        }
    } catch (error) {
        console.error("Error loading dashboard data:", error);
    }
}

function switchTab(tabId, element) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.menu-item').forEach(btn => btn.classList.remove('active'));

    document.getElementById(`tab-${tabId}`).classList.add('active');
    element.classList.add('active');
    activeTabName = tabId;

    // مسح خانة البحث عند الانتقال بين التبابيب لضمان تجربة مستخدم أفضل
    const searchInput = document.getElementById('globalSearch');
    if (searchInput) {
        searchInput.value = '';
        // إعادة عرض البيانات الأصلية للتبويب الجديد
        if (tabId === 'workDistribution') renderWorkDistributionTable(cachedWorkData);
        if (tabId === 'nightShift') renderNightShiftTable(cachedNightData);
        if (tabId === 'annualVacations') renderVacationsTable(cachedVacationsData);
    }

    const titles = {
        workDistribution: { title: "Work Distribution", subtitle: "Infrastructure Support & Group Branch Assignments" },
        nightShift: { title: "Night Shift Schedule", subtitle: "Weekly Rotation & Coverage Tracker" },
        annualVacations: { title: "Annual Vacations", subtitle: "Engineers Leave Schedule" }
    };

    document.getElementById('pageTitle').innerText = titles[tabId].title;
    document.getElementById('pageSubtitle').innerText = titles[tabId].subtitle;
}

function renderWorkDistributionTable(data) {
    const tbody = document.getElementById('workDistributionTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let totalBranchesSum = 0;

    data.forEach((row, index) => {
        const groupName = row['Group Name'] || row['Group_Name'] || `G${index + 1}`;
        const engineers = row['Engineers'] || '';
        const branches = row['Assigned Branches'] || row['Assigned_Branches'] || '';
        const totalBranches = row['Total Branches'] || row['Total_Branches'] || branches.split(',').length;

        totalBranchesSum += parseInt(totalBranches) || 0;

        const branchListArray = branches.split(',').map(b => b.trim()).filter(b => b.length > 0);
        let displayBranchesText = branchListArray.slice(0, 4).join(', ');
        if (branchListArray.length > 4) {
            displayBranchesText += `, ... and ${branchListArray.length - 4} more (Click to view all)`;
        }

        const tr = document.createElement('tr');
        tr.className = 'clickable-row';
        tr.onclick = () => openBranchModal(groupName, engineers, branchListArray);

        tr.innerHTML = `
            <td><span class="group-badge">${groupName}</span></td>
            <td><strong>${engineers}</strong></td>
            <td style="color: var(--text-secondary);">${displayBranchesText}</td>
            <td style="text-align: right; font-weight: 700; color: #FFFFFF;">${totalBranches}</td>
        `;
        tbody.appendChild(tr);
    });

    // تحديث العدّادات في أعلى الصفحة فقط لو دي البيانات الأصلية الكاملة
    if (data === cachedWorkData) {
        document.getElementById('metricBranchesCount').innerText = `${totalBranchesSum} Branches`;
        document.getElementById('metricGroupsCount').innerText = `${data.length} Groups (G1 to G${data.length})`;
    }
}

function renderNightShiftTable(data) {
    const tbody = document.getElementById('nightShiftTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (data.length > 0 && data === cachedNightData) {
        document.getElementById('metricNightShift').innerText = data[0]['Engineer Name'] || data[0]['Infrastructure Support'] || 'N/A';
    }

    data.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${row['Engineer Name'] || row['Infrastructure Support'] || ''}</strong></td>
            <td>${formatExcelDate(row['Start Date'])}</td>
            <td>${formatExcelDate(row['End Date'])}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderVacationsTable(data) {
    const tbody = document.getElementById('annualVacationsTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const validData = data.filter(row => {
        const name = row['Engineer Name'] || row['Infrastructure Support'];
        return name && name !== 'Available' && name !== 'NOT AVAILABLE';
    });

    validData.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${row['Engineer Name'] || row['Infrastructure Support'] || ''}</strong></td>
            <td>${formatExcelDate(row['Start Date'])}</td>
            <td>${formatExcelDate(row['End Date'])}</td>
            <td><span class="status-badge-leave">${row['Vacation Status'] || 'Scheduled'}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

// Modal Control Functions
function openBranchModal(groupName, engineers, branchesArray) {
    currentModalBranches = branchesArray;
    
    document.getElementById('modalGroupName').innerText = `Group: ${groupName}`;
    document.getElementById('modalEngineersName').innerText = `Engineers: ${engineers}`;
    document.getElementById('modalTotalCount').innerText = `Total: ${branchesArray.length} Branches`;
    document.getElementById('modalBranchSearch').value = '';

    renderModalBranchTags(branchesArray);
    document.getElementById('branchModal').classList.add('active');
}

function closeBranchModal(event) {
    document.getElementById('branchModal').classList.remove('active');
}

function renderModalBranchTags(branchesArray) {
    const container = document.getElementById('modalBranchesList');
    container.innerHTML = '';

    if (branchesArray.length === 0) {
        container.innerHTML = '<p style="color: var(--text-secondary);">No branches assigned.</p>';
        return;
    }

    branchesArray.forEach(branch => {
        const tag = document.createElement('div');
        tag.className = 'branch-tag-item';
        tag.innerText = branch;
        container.appendChild(tag);
    });
}

function filterModalBranches() {
    const query = document.getElementById('modalBranchSearch').value.toLowerCase();
    const filtered = currentModalBranches.filter(b => b.toLowerCase().includes(query));
    renderModalBranchTags(filtered);
}

// نظام بحث ذكي وشامل يعمل بكفاءة داخل التبويب النشط حالياً
function filterTableData() {
    const query = document.getElementById('globalSearch').value.toLowerCase();
    
    if (activeTabName === 'workDistribution') {
        const filtered = cachedWorkData.filter(row => {
            const text = `${row['Group Name'] || ''} ${row['Engineers'] || ''} ${row['Assigned Branches'] || ''}`.toLowerCase();
            return text.includes(query);
        });
        renderWorkDistributionTable(filtered);
    } 
    else if (activeTabName === 'nightShift') {
        const filtered = cachedNightData.filter(row => {
            const engName = row['Engineer Name'] || row['Infrastructure Support'] || '';
            const startDate = formatExcelDate(row['Start Date']);
            const endDate = formatExcelDate(row['End Date']);
            const text = `${engName} ${startDate} ${endDate}`.toLowerCase();
            return text.includes(query);
        });
        renderNightShiftTable(filtered);
    } 
    else if (activeTabName === 'annualVacations') {
        const filtered = cachedVacationsData.filter(row => {
            const engName = row['Engineer Name'] || row['Infrastructure Support'] || '';
            const startDate = formatExcelDate(row['Start Date']);
            const endDate = formatExcelDate(row['End Date']);
            const status = row['Vacation Status'] || '';
            const text = `${engName} ${startDate} ${endDate} ${status}`.toLowerCase();
            return text.includes(query);
        });
        renderVacationsTable(filtered);
    }
}
