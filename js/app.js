let allCases = [];
let userCases = [];
let currentCase = null;

async function initApp() {
    try {
        await initializeMsal();
        if (!isUserLoggedIn()) {
            showLoginScreen();
        }
    } catch (error) {
        console.error("Error inicializando app:", error);
        showError("Error al inicializar la aplicación");
    }
}

function showLoginScreen() {
    document.getElementById('loginScreen').style.display = 'flex';
    document.getElementById('appScreen').style.display = 'none';
    document.getElementById('loginBtn').addEventListener('click', signIn);
}

async function showApp() {
    try {
        showLoading("Cargando datos...");
        document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('appScreen').style.display = 'block';
        const userInfo = await getUserInfo();
        displayUserInfo(userInfo);
        await loadCases();
        setupEventListeners();
        hideLoading();
    } catch (error) {
        console.error("Error mostrando app:", error);
        showError("Error cargando la aplicación: " + error.message);
        hideLoading();
    }
}

function displayUserInfo(userInfo) {
    document.getElementById('userName').textContent = userInfo.displayName;
    document.getElementById('userEmail').textContent = userInfo.mail || userInfo.userPrincipalName;
}

async function loadCases() {
    try {
        showLoading("Cargando casos...");
        allCases = await readAllRows();
        const userName = getCurrentUserName();
        userCases = filterUserCases(allCases, userName);
        updateStatistics();
        displayCases(userCases);
        hideLoading();
    } catch (error) {
        console.error("Error cargando casos:", error);
        showError("Error al cargar los casos: " + error.message);
        hideLoading();
    }
}

function updateStatistics() {
    const totalCases = userCases.length;
    const pending = userCases.filter(c => !c[excelConfig.columns.fechaCierre]).length;
    const closed = userCases.filter(c => c[excelConfig.columns.fechaCierre]).length;
    const urgent = userCases.filter(c => {
        const days = calculateDaysRemaining(c[excelConfig.columns.fechaVencimiento]);
        return days !== null && days < 3;
    }).length;
    
    document.getElementById('totalCases').textContent = totalCases;
    document.getElementById('pendingCases').textContent = pending;
    document.getElementById('closedCases').textContent = closed;
    document.getElementById('urgentCases').textContent = urgent;
}

function displayCases(cases) {
    const tbody = document.getElementById('casesTableBody');
    tbody.innerHTML = '';
    
    if (cases.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center">No tienes casos asignados</td></tr>';
        return;
    }
    
    cases.forEach((caso, index) => {
        const row = createCaseRow(caso, index);
        tbody.appendChild(row);
    });
}

function createCaseRow(caso, index) {
    const tr = document.createElement('tr');
    const daysRemaining = calculateDaysRemaining(caso[excelConfig.columns.fechaVencimiento]);
    let priorityClass = 'priority-normal';
    let priorityIcon = '🟢';
    
    if (daysRemaining !== null) {
        if (daysRemaining < 3) {
            priorityClass = 'priority-high';
            priorityIcon = '🔴';
        } else if (daysRemaining < 7) {
            priorityClass = 'priority-medium';
            priorityIcon = '🟡';
        }
    }
    
    const isClosed = caso[excelConfig.columns.fechaCierre];
    
    tr.className = priorityClass;
    tr.innerHTML = `
        <td>${priorityIcon}</td>
        <td>${caso[excelConfig.columns.idCaso] || '-'}</td>
        <td>${caso[excelConfig.columns.paciente] || '-'}</td>
        <td>${caso[excelConfig.columns.servicio] || '-'}</td>
        <td>${caso[excelConfig.columns.fechaSolicitud] || '-'}</td>
        <td>${caso[excelConfig.columns.estado] || 'PENDIENTE'}</td>
        <td>${daysRemaining !== null ? daysRemaining + ' días' : '-'}</td>
        <td>
            ${isClosed ? 
                '<span class="badge bg-secondary">Cerrado</span>' : 
                `<button class="btn btn-sm btn-primary" onclick="openCaseDetail(${index})">Gestionar</button>`
            }
        </td>
    `;
    
    return tr;
}

function openCaseDetail(caseIndex) {
    currentCase = userCases[caseIndex];
    displayCaseDetail(currentCase, caseIndex);
    const modal = new bootstrap.Modal(document.getElementById('caseDetailModal'));
    modal.show();
}

function displayCaseDetail(caso, caseIndex) {
    document.getElementById('detailIdCaso').textContent = caso[excelConfig.columns.idCaso] || '-';
    document.getElementById('detailPaciente').textContent = caso[excelConfig.columns.paciente] || '-';
    document.getElementById('detailServicio').textContent = caso[excelConfig.columns.servicio] || '-';
    document.getElementById('detailFechaSolicitud').textContent = caso[excelConfig.columns.fechaSolicitud] || '-';
    document.getElementById('detailFechaVencimiento').textContent = caso[excelConfig.columns.fechaVencimiento] || '-';
    document.getElementById('detailTelefono').textContent = caso[excelConfig.columns.telefono] || '-';
    document.getElementById('detailDireccion').textContent = caso[excelConfig.columns.direccion] || '-';
    document.getElementById('detailObservacionesRISS').textContent = caso[excelConfig.columns.observacionesRISS] || '-';
    
    populateStatusDropdown();
    document.getElementById('estadoSelect').value = caso[excelConfig.columns.estado] || 'PENDIENTE';
    document.getElementById('observacionesTextarea').value = caso[excelConfig.columns.observaciones] || '';
    
    displayCallHistory(caso);
    setupCaseDetailButtons(caseIndex);
    
    if (shouldAutoClose(caso)) {
        showAutoCloseAlert();
    }
}

function populateStatusDropdown() {
    const select = document.getElementById('estadoSelect');
    select.innerHTML = '';
    excelConfig.estadosPermitidos.forEach(estado => {
        const option = document.createElement('option');
        option.value = estado;
        option.textContent = estado;
        select.appendChild(option);
    });
}

function displayCallHistory(caso) {
    const container = document.getElementById('callHistory');
    container.innerHTML = '';
    
    for (let i = 1; i <= 3; i++) {
        const callData = caso[excelConfig.columns[`llamado${i}`]];
        const callDiv = document.createElement('div');
        callDiv.className = 'call-record mb-2';
        
        if (callData) {
            callDiv.innerHTML = `<strong>Llamado ${i}:</strong><br>${callData}`;
        } else {
            callDiv.innerHTML = `<strong>Llamado ${i}:</strong> <em>No registrado</em>`;
        }
        container.appendChild(callDiv);
    }
}

function setupCaseDetailButtons(caseIndex) {
    document.getElementById('saveChangesBtn').onclick = () => saveCaseChanges(caseIndex);
    document.getElementById('registerCall1Btn').onclick = () => openCallModal(caseIndex, 1);
    document.getElementById('registerCall2Btn').onclick = () => openCallModal(caseIndex, 2);
    document.getElementById('registerCall3Btn').onclick = () => openCallModal(caseIndex, 3);
    document.getElementById('closeCaseBtn').onclick = () => confirmCloseCase(caseIndex);
}

async function saveCaseChanges(caseIndex) {
    try {
        showLoading("Guardando cambios...");
        const estado = document.getElementById('estadoSelect').value;
        const observaciones = document.getElementById('observacionesTextarea').value;
        const rowIndex = caseIndex + 1;
        
        await updateStatus(rowIndex, estado);
        await updateObservations(rowIndex, observaciones);
        
        userCases[caseIndex][excelConfig.columns.estado] = estado;
        userCases[caseIndex][excelConfig.columns.observaciones] = observaciones;
        
        showSuccess("Cambios guardados exitosamente");
        await loadCases();
        hideLoading();
    } catch (error) {
        console.error("Error guardando cambios:", error);
        showError("Error al guardar cambios: " + error.message);
        hideLoading();
    }
}

function openCallModal(caseIndex, callNumber) {
    const caso = userCases[caseIndex];
    
    if (callNumber === 2 && !caso[excelConfig.columns.llamado1]) {
        showError("Debes registrar el Llamado 1 primero");
        return;
    }
    
    if (callNumber === 3 && !caso[excelConfig.columns.llamado2]) {
        showError("Debes registrar el Llamado 2 primero");
        return;
    }
    
    document.getElementById('callModalLabel').textContent = `Registrar Llamado ${callNumber}`;
    document.getElementById('currentDateTime').textContent = formatDateTime(new Date());
    
    const resultSelect = document.getElementById('callResultSelect');
    resultSelect.innerHTML = '';
    excelConfig.resultadosLlamado.forEach(resultado => {
        const option = document.createElement('option');
        option.value = resultado;
        option.textContent = resultado;
        resultSelect.appendChild(option);
    });
    
    document.getElementById('callObservationTextarea').value = '';
    document.getElementById('saveCallBtn').onclick = () => saveCall(caseIndex, callNumber);
    
    const modal = new bootstrap.Modal(document.getElementById('registerCallModal'));
    modal.show();
}

async function saveCall(caseIndex, callNumber) {
    try {
        const result = document.getElementById('callResultSelect').value;
        const observation = document.getElementById('callObservationTextarea').value;
        
        if (!observation.trim()) {
            showError("Debes agregar una observación");
            return;
        }
        
        showLoading("Registrando llamado...");
        const rowIndex = caseIndex + 1;
        await registerCall(rowIndex, callNumber, result, observation);
        
        const timestamp = formatDateTime(new Date());
        const callText = `${timestamp} - ${result} - ${observation}`;
        userCases[caseIndex][excelConfig.columns[`llamado${callNumber}`]] = callText;
        
        showSuccess("Llamado registrado exitosamente");
        bootstrap.Modal.getInstance(document.getElementById('registerCallModal')).hide();
        displayCaseDetail(userCases[caseIndex], caseIndex);
        await loadCases();
        hideLoading();
    } catch (error) {
        console.error("Error guardando llamado:", error);
        showError("Error al registrar llamado: " + error.message);
        hideLoading();
    }
}

function confirmCloseCase(caseIndex) {
    const caso = userCases[caseIndex];
    const hasCall = caso[excelConfig.columns.llamado1];
    const hasObservation = caso[excelConfig.columns.observaciones];
    
    if (!hasCall && !hasObservation) {
        showError("No puedes cerrar un caso sin al menos 1 intento de contacto o una observación");
        return;
    }
    
    if (confirm("¿Estás seguro que deseas cerrar este caso?")) {
        closeCaseNow(caseIndex);
    }
}

async function closeCaseNow(caseIndex) {
    try {
        showLoading("Cerrando caso...");
        let estado = document.getElementById('estadoSelect').value;
        
        if (shouldAutoClose(userCases[caseIndex])) {
            estado = "CERRADO POR INTENTOS";
        }
        
        const rowIndex = caseIndex + 1;
        await closeCase(rowIndex, estado);
        
        showSuccess("Caso cerrado exitosamente");
        bootstrap.Modal.getInstance(document.getElementById('caseDetailModal')).hide();
        await loadCases();
        hideLoading();
    } catch (error) {
        console.error("Error cerrando caso:", error);
        showError("Error al cerrar caso: " + error.message);
        hideLoading();
    }
}

function showAutoCloseAlert() {
    const alertDiv = document.createElement('div');
    alertDiv.className = 'alert alert-warning mt-3';
    alertDiv.innerHTML = '⚠️ Este caso tiene 3 intentos fallidos. Se recomienda cerrarlo.';
    document.getElementById('caseDetailContent').prepend(alertDiv);
}

function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', handleSearch);
    document.getElementById('filterAll').addEventListener('click', () => filterCases('all'));
    document.getElementById('filterUrgent').addEventListener('click', () => filterCases('urgent'));
    document.getElementById('filterPending').addEventListener('click', () => filterCases('pending'));
    document.getElementById('refreshBtn').addEventListener('click', loadCases);
    document.getElementById('logoutBtn').addEventListener('click', signOut);
}

function handleSearch(e) {
    const searchTerm = e.target.value.toLowerCase();
    if (!searchTerm) {
        displayCases(userCases);
        return;
    }
    const filtered = userCases.filter(caso => {
        return Object.values(caso).some(value => 
            value && value.toString().toLowerCase().includes(searchTerm)
        );
    });
    displayCases(filtered);
}

function filterCases(filter) {
    let filtered = userCases;
    switch (filter) {
        case 'urgent':
            filtered = userCases.filter(c => {
                const days = calculateDaysRemaining(c[excelConfig.columns.fechaVencimiento]);
                return days !== null && days < 3;
            });
            break;
        case 'pending':
            filtered = userCases.filter(c => !c[excelConfig.columns.fechaCierre]);
            break;
        case 'all':
        default:
            filtered = userCases;
    }
    displayCases(filtered);
}

function showLoading(message = "Cargando...") {
    document.getElementById('loadingMessage').textContent = message;
    document.getElementById('loadingOverlay').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loadingOverlay').style.display = 'none';
}

function showError(message) {
    alert("❌ " + message);
}

function showSuccess(message) {
    alert("✅ " + message);
}

document.addEventListener('DOMContentLoaded', initApp);
