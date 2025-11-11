let workbookSession = null;

async function getSiteInfo() {
    try {
        const token = await getToken();
        const url = `${graphConfig.graphSitesEndpoint}/${excelConfig.siteUrl}:${excelConfig.sitePath}`;
        
        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) {
            throw new Error(`Error obteniendo sitio: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error("Error obteniendo sitio:", error);
        throw error;
    }
}

async function createWorkbookSession(itemId) {
    try {
        const token = await getToken();
        const url = `https://graph.microsoft.com/v1.0/sites/${excelConfig.siteUrl}/drive/items/${itemId}/workbook/createSession`;
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                persistChanges: true
            })
        });
        
        if (!response.ok) {
            throw new Error(`Error creando sesión: ${response.status}`);
        }
        
        const session = await response.json();
        workbookSession = session.id;
        return session.id;
    } catch (error) {
        console.error("Error creando sesión de Excel:", error);
        throw error;
    }
}

async function readAllRows() {
    try {
        const token = await getToken();
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${excelConfig.fileId}/workbook/worksheets/${excelConfig.sheetName}/usedRange`;
        
        const headers = {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
        
        if (workbookSession) {
            headers['workbook-session-id'] = workbookSession;
        }
        
        const response = await fetch(url, { headers });
        
        if (!response.ok) {
            throw new Error(`Error leyendo Excel: ${response.status} - ${await response.text()}`);
        }
        
        const data = await response.json();
        return data.values;
    } catch (error) {
        console.error("Error leyendo filas:", error);
        throw error;
    }
}

function filterUserCases(allRows, userName) {
    const headerRow = allRows[0];
    const dataRows = allRows.slice(1);
    
    const userCases = dataRows.filter(row => {
        const assignedTo = row[excelConfig.columns.asignadoA];
        return assignedTo && assignedTo.toString().toLowerCase().includes(userName.toLowerCase());
    });
    
    return userCases;
}

async function updateCell(rowIndex, columnIndex, value) {
    try {
        const token = await getToken();
        const cellAddress = columnToLetter(columnIndex) + (rowIndex + 1);
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${excelConfig.fileId}/workbook/worksheets/${excelConfig.sheetName}/range(address='${cellAddress}')`;
        
        const headers = {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
        
        if (workbookSession) {
            headers['workbook-session-id'] = workbookSession;
        }
        
        const response = await fetch(url, {
            method: 'PATCH',
            headers: headers,
            body: JSON.stringify({
                values: [[value]]
            })
        });
        
        if (!response.ok) {
            throw new Error(`Error actualizando celda: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error("Error actualizando celda:", error);
        throw error;
    }
}

async function updateRow(rowIndex, updates) {
    try {
        const promises = Object.entries(updates).map(([colIndex, value]) => {
            return updateCell(rowIndex, parseInt(colIndex), value);
        });
        
        await Promise.all(promises);
        return true;
    } catch (error) {
        console.error("Error actualizando fila:", error);
        throw error;
    }
}

async function registerCall(rowIndex, callNumber, result, observation) {
    const now = new Date();
    const timestamp = formatDateTime(now);
    const callText = `${timestamp} - ${result} - ${observation}`;
    
    let columnIndex;
    switch (callNumber) {
        case 1:
            columnIndex = excelConfig.columns.llamado1;
            break;
        case 2:
            columnIndex = excelConfig.columns.llamado2;
            break;
        case 3:
            columnIndex = excelConfig.columns.llamado3;
            break;
        default:
            throw new Error("Número de llamado inválido");
    }
    
    return await updateCell(rowIndex, columnIndex, callText);
}

async function closeCase(rowIndex, estado) {
    const now = new Date();
    const userName = getCurrentUserName();
    
    const updates = {
        [excelConfig.columns.resolutor]: userName,
        [excelConfig.columns.fechaCierre]: formatDate(now),
        [excelConfig.columns.estado]: estado
    };
    
    return await updateRow(rowIndex, updates);
}

async function updateObservations(rowIndex, observations) {
    return await updateCell(rowIndex, excelConfig.columns.observaciones, observations);
}

async function updateStatus(rowIndex, status) {
    return await updateCell(rowIndex, excelConfig.columns.estado, status);
}

function shouldAutoClose(caso) {
    const call1 = caso[excelConfig.columns.llamado1];
    const call2 = caso[excelConfig.columns.llamado2];
    const call3 = caso[excelConfig.columns.llamado3];
    
    if (!call1 || !call2 || !call3) {
        return false;
    }
    
    const hasSuccessfulContact = [call1, call2, call3].some(call => 
        call.includes("Contactado exitosamente")
    );
    
    return !hasSuccessfulContact;
}

function columnToLetter(column) {
    let temp, letter = '';
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function formatDateTime(date) {
    const dateStr = formatDate(date);
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${dateStr} ${hours}:${minutes}`;
}

function calculateDaysRemaining(fechaVencimiento) {
    if (!fechaVencimiento) return null;
    
    const vencimiento = new Date(fechaVencimiento);
    const hoy = new Date();
    const diffTime = vencimiento - hoy;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays;
}
