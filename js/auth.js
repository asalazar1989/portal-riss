let msalInstance;
let currentUser = null;

function initializeMsal() {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    return msalInstance.handleRedirectPromise()
        .then(handleResponse)
        .catch(err => {
            console.error("Error en redirect:", err);
        });
}

function handleResponse(response) {
    if (response !== null && response.account) {
        currentUser = response.account;
        showApp();
        return;
    }
    
    const currentAccounts = msalInstance.getAllAccounts();
    if (currentAccounts.length > 0) {
        currentUser = currentAccounts[0];
        showApp();
    }
}

function signIn() {
    msalInstance.loginRedirect(loginRequest);
}

function signOut() {
    msalInstance.logoutRedirect({
        account: currentUser,
        postLogoutRedirectUri: msalConfig.auth.redirectUri
    });
}

async function getToken() {
    const account = currentUser || msalInstance.getAllAccounts()[0];
    
    if (!account) {
        throw new Error("No hay usuario autenticado");
    }

    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes: loginRequest.scopes,
            account: account
        });
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            await msalInstance.acquireTokenRedirect({
                scopes: loginRequest.scopes,
                account: account
            });
        }
        throw error;
    }
}

async function getUserInfo() {
    try {
        const token = await getToken();
        const response = await fetch(graphConfig.graphMeEndpoint, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) {
            throw new Error(`Error obteniendo usuario: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error("Error obteniendo info de usuario:", error);
        throw error;
    }
}

function isUserLoggedIn() {
    return currentUser !== null;
}

function getCurrentUserName() {
    return currentUser ? currentUser.name : null;
}
