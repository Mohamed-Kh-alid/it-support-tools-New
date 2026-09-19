/* ==========================================================
   Pharma Overseas - IT Service Desk Script Logic
   ========================================================== */

const GOOGLE_SHEETS_DIRECT_URL = "https://docs.google.com/spreadsheets/d/19eCMXNFvRVdgfWQi3j5hDOwyoAU-iExn_Fqv4D8VXYA/export?format=xlsx";

let cachedWorkData = [];
let cachedNightData = [];
let cachedVacationsData = [];
let usersDatabase = {};
let activeTabName = 'workDistribution';
let currentModalBranches = [];

let fileUsername = "pharma";
let filePassword = "Ph@rma1515";
let isVacationsUnlocked = false;

document.addEventListener('DOMContentLoaded', () => {
    loadOnlineDashboardData();
    initLiveClock();
    
    const passInput = document.getElementById('vacationPassword');
    if (passInput) {
        passInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') checkVacationsLogin();
        });
    }
});

// تحديث الساعة والتوقيت الحي تلقائياً
function initLiveClock() {
    function updateClock() {
        const now = new Date();
        let hours = now.getHours();
        const minutes = String(now.getMinutes()).padStart(2, '0');
        hours = String(hours).padStart(2, '0');
        
        const timeElement = document.getElementById('liveClockTime');
        const tzElement = document.getElementById('liveTimeZone');
        
        if (timeElement) timeElement.innerText = `${hours}:${minutes}`;
        
        if (tzElement) {
            const timeZoneOffset = -now.getTimezoneOffset() / 60;
            const sign = timeZoneOffset >= 0 ? '+' : '';
            tzElement.innerText = `GMT ${sign}${timeZoneOffset}`;
        }
    }
    updateClock();
    setInterval(updateClock, 10000);
}

function checkVacationsLogin() {
    const userInput = document.getElementById('vacationUsername').value.trim().toLowerCase();
    const passInput = document.getElementById('vacationPassword').value.trim();
    const errorMsg = document.getElementById('vacationErrorMsg');

    if (userInput === fileUsername.toLowerCase() && passInput === filePassword) {
        isVacationsUnlocked = true;
        document.getElementById('vacationsLoginBox').style.display = 'none';
        document.getElementById('vacationsTableContainer').style.display = 'block';
        renderVacationsTable(cachedVacationsData);
    } else {
        errorMsg.innerText = "Incorrect username or password!";
    }
}

function formatExcelDate(serial) {
    if (!serial) return 'N/A';
    if (isNaN(serial)) return serial;
    const utcDays = Math.floor(serial - 25569);
    const date = new Date(utcDays * 86400 * 1000);
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

function parseExcelDateObj(serial) {
    if (!serial) return null;
    if (typeof serial === 'string' && serial.includes('/')) {
        const parts = serial.split('/');
        if (parts.length === 3) {
            return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        }
    }
    if (!isNaN(serial)) {
        const utcDays = Math.floor(serial - 25569);
        const dateInfo = new Date(utcDays * 86400 * 1000);
        return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
    }
    const parsed = new Date(serial);
    return isNaN(parsed) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

async function loadOnlineDashboardData() {
    try {
        const response = await fetch(GOOGLE_SHEETS_DIRECT_URL);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        if (workbook.Sheets['User_PassWord']) {
            const rawCredsJson = XLSX.utils.sheet_to_json(workbook.Sheets['User_PassWord'], { header: 1 });
            if (rawCredsJson.length > 1) {
                if (rawCredsJson[1][0]) fileUsername = String(rawCredsJson[1][0]).trim();
                if (rawCredsJson[1][1]) filePassword = String(rawCredsJson[1][1]).trim();
            }
        }

        if (workbook.Sheets['Users_Data']) {
            const rawUsersRows = XLSX.utils.sheet_to_json(workbook.Sheets['Users_Data'], { header: 1 });
            for (let i = 1; i < rawUsersRows.length; i++) {
                const r = rawUsersRows[i];
                if (!r || !r[0]) continue;
                const name = String(r[0]).trim();
                if (name) {
                    usersDatabase[name.toLowerCase()] = {
                        id: r[1] ? String(r[1]).trim() : `PO-2026-${String(i).padStart(2, '0')}`,
                        email: r[2] ? String(r[2]).trim() : `${name.toLowerCase().replace(/[^a-z]/g, '.')}@pharmaoverseas.com`,
                        department: r[3] ? String(r[3]).trim() : 'Information Technology',
                        jobTitle: r[4] ? String(r[4]).trim() : 'IT Technical Support Specialist',
                        phone: r[5] ? String(r[5]).trim() : 'N/A'
                    };
                }
            }
        }

        if (workbook.Sheets['Night_Shift']) {
            cachedNightData = XLSX.utils.sheet_to_json(workbook.Sheets['Night_Shift'], { range: 3 });
            renderNightShiftTable(cachedNightData);
            updateCurrentNightShiftEngineer(cachedNightData);
        }

        if (workbook.Sheets['Annual_Vacations']) {
            cachedVacationsData = XLSX.utils.sheet_to_json(workbook.Sheets['Annual_Vacations'], { range: 3 });
        }

        if (workbook.Sheets['Work_Distribution']) {
            cachedWorkData = XLSX.utils.sheet_to_json(workbook.Sheets['Work_Distribution'], { range: 3 });
            renderWorkDistributionTable(cachedWorkData);
        }
    } catch (error) {
        console.error("Error loading dashboard data:", error);
    }
}

function updateCurrentNightShiftEngineer(data) {
    const metricNightShift = document.getElementById('metricNightShift');
    if (!metricNightShift) return;
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    let activeEngineer = 'No Active Shift';

    for (let row of data) {
        const engName = row['Engineer Name'] || row['Infrastructure Support'];
        const rawStart = row['Start Date'];
        const rawEnd = row['End Date'];
        if (!engName || !rawStart || !rawEnd) continue;

        const startDate = parseExcelDateObj(rawStart);
        const endDate = parseExcelDateObj(rawEnd);
        if (startDate && endDate) {
            const cleanStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()).getTime();
            const cleanEnd = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()).getTime();
            if (today >= cleanStart && today <= cleanEnd) {
                activeEngineer = engName;
                break;
            }
        }
    }
    metricNightShift.innerText = activeEngineer;
}

function toggleMobileMenu() {
    const sidebar = document.getElementById('appSidebar');
    sidebar.classList.toggle('show-sidebar');
}

function switchTab(tabId, element) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.menu-item').forEach(btn => btn.classList.remove('active'));
    document.getElementById(`tab-${tabId}`).classList.add('active');
    element.classList.add('active');
    activeTabName = tabId;

    // إغلاق القائمة الجانبية تلقائياً في الهواتف عند اختيار تبويب
    if (window.innerWidth <= 1024) {
        document.getElementById('appSidebar').classList.remove('show-sidebar');
    }

    const searchInput = document.getElementById('globalSearch');
    if (searchInput) {
        searchInput.value = '';
        if (tabId === 'workDistribution') renderWorkDistributionTable(cachedWorkData);
        if (tabId === 'nightShift') renderNightShiftTable(cachedNightData);
        if (tabId === 'annualVacations' && isVacationsUnlocked) renderVacationsTable(cachedVacationsData);
    }

    const titles = {
        workDistribution: { title: "Work Distribution", subtitle: "Infrastructure Support & Branch Assignments" },
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
        let displayBranchesText = branchListArray.slice(0, 3).join(', ');
        if (branchListArray.length > 3) displayBranchesText += `, ... (${branchListArray.length - 3} more)`;

        const tr = document.createElement('tr');
        tr.className = 'clickable-row';
        tr.onclick = () => openBranchModal(groupName, engineers, branchListArray);
        tr.innerHTML = `
            <td><strong>${engineers}</strong></td>
            <td style="color: var(--text-secondary);">${displayBranchesText}</td>
            <td style="text-align: right;"><span class="index-badge">${totalBranches} Branches</span></td>
        `;
        tbody.appendChild(tr);
    });

    if (data === cachedWorkData) {
        document.getElementById('metricBranchesCount').innerText = `${totalBranchesSum} Branches`;
        document.getElementById('metricGroupsCount').innerText = `${data.length} Teams`;
    }
}

function renderNightShiftTable(data) {
    const tbody = document.getElementById('nightShiftTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    document.getElementById('rotationCountBadge').innerText = `${data.length} Rotations`;

    data.forEach((row, index) => {
        const engName = row['Engineer Name'] || row['Infrastructure Support'] || '';
        const startDate = formatExcelDate(row['Start Date']);
        const endDate = formatExcelDate(row['End Date']);
        const rowNum = index + 1;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="index-badge">${rowNum}</span></td>
            <td><strong style="color: #FFFFFF;">${engName}</strong></td>
            <td>${startDate}</td>
            <td style="color: var(--success-text); font-weight: 600;">${endDate}</td>
            <td style="text-align: right;">
                <button class="action-btn" onclick="openEmployeeDrawer('${engName}', '${startDate}', '${endDate}', '${rowNum}')">
                    View ID &rarr;
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderVacationsTable(data) {
    const tbody = document.getElementById('annualVacationsTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    data.filter(row => {
        const name = row['Engineer Name'] || row['Infrastructure Support'];
        return name && name !== 'Available' && name !== 'NOT AVAILABLE';
    }).forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${row['Engineer Name'] || row['Infrastructure Support'] || ''}</strong></td>
            <td>${formatExcelDate(row['Start Date'])}</td>
            <td>${formatExcelDate(row['End Date'])}</td>
            <td><span style="background: var(--success-bg); color: var(--success-text); padding: 4px 10px; border-radius: 6px; font-size: 0.78rem;">${row['Vacation Status'] || 'Scheduled'}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function openEmployeeDrawer(engineerName, startDate, endDate, rowNum) {
    const cleanName = engineerName.trim();
    const lookupKey = cleanName.toLowerCase();
    
    const userData = usersDatabase[lookupKey] || {
        id: `PO-2026-${String(rowNum).padStart(2, '0')}`,
        email: `${cleanName.toLowerCase().replace(/[^a-z]/g, '.')}@pharmaoverseas.com`,
        phone: 'N/A',
        department: "Information Technology",
        jobTitle: "IT Technical Support Specialist"
    };

    document.getElementById('drawerEmpIdBadge').innerText = `#${userData.id}`;
    document.getElementById('drawerEmpName').innerText = cleanName;
    document.getElementById('drawerEmpTitle').innerText = userData.jobTitle;
    document.getElementById('drawerEmpEmail').innerText = userData.email;
    document.getElementById('drawerEmpPhone').innerText = userData.phone;
    document.getElementById('drawerEmpIdVal').innerText = userData.id;
    document.getElementById('drawerEmpDept').innerText = userData.department;
    document.getElementById('drawerStartDate').innerText = startDate;
    document.getElementById('drawerEndDate').innerText = endDate;

    document.getElementById('employeeDrawer').classList.add('active');
}

function closeEmployeeDrawer() {
    document.getElementById('employeeDrawer').classList.remove('active');
}

function openBranchModal(groupName, engineers, branchesArray) {
    currentModalBranches = branchesArray;
    document.getElementById('modalGroupName').innerText = groupName;
    document.getElementById('modalEngineersName').innerText = `Engineers: ${engineers}`;
    document.getElementById('modalTotalCount').innerText = `Total: ${branchesArray.length} Branches`;
    document.getElementById('modalBranchSearch').value = '';
    renderModalBranchTags(branchesArray);
    document.getElementById('branchModal').classList.add('active');
}

function closeBranchModal() {
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
        tag.style.cssText = "background: rgba(255,255,255,0.04); border: 1px solid var(--border-color); color: var(--text-primary); padding: 6px 12px; border-radius: 6px; font-size: 0.82rem;";
        tag.innerText = branch;
        container.appendChild(tag);
    });
}

function filterModalBranches() {
    const query = document.getElementById('modalBranchSearch').value.toLowerCase();
    const filtered = currentModalBranches.filter(b => b.toLowerCase().includes(query));
    renderModalBranchTags(filtered);
}

function filterTableData() {
    const query = document.getElementById('globalSearch').value.toLowerCase();
    if (activeTabName === 'nightShift') {
        const filtered = cachedNightData.filter(row => {
            const name = row['Engineer Name'] || row['Infrastructure Support'] || '';
            return name.toLowerCase().includes(query);
        });
        renderNightShiftTable(filtered);
    } else if (activeTabName === 'workDistribution') {
        const filtered = cachedWorkData.filter(row => {
            const engineers = row['Engineers'] || '';
            const groupName = row['Group Name'] || '';
            return engineers.toLowerCase().includes(query) || groupName.toLowerCase().includes(query);
        });
        renderWorkDistributionTable(filtered);
    }
}
