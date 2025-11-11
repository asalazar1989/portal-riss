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

async function updateRow(rowIndex, up
