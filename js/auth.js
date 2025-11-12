let msalInstance;
let currentUser = null;

function initializeMsal() {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    return msalInstance.handleRedirectPromise()
        .then(handleResponse)
        .catch(err => {
            console.error("Error en redirect:", err);
            throw err;
        });
}

function handleResponse(response) {
    if (response !== null) {
        currentUser = response.account;
        showApp();
    } else {
        const currentAccounts = msalInstance.getAllAccounts();
        if (currentAccounts.length === 0) {
            return;
        } else if (currentAccounts.length === 1) {
            currentUser = currentAccounts[0];
            showApp();
        }
    }
}

async function signIn() {
    try {
        await msalInstance.loginRedirect(loginRequest);
    } catch (error) {
        console.error("Error en login:", error);
        showError("Error al iniciar sesión: " + error.message);
    }
}

function signOut() {
    const logoutRequest = {
        account: currentUser,
        postLogoutRedirectUri: msalConfig.auth.redirectUri
    };
    msalInstance.logoutRedirect(logoutRequest);
}

async function getToken() {
    const account = currentUser || msalInstance.getAllAccounts()[0];
    
    if (!account) {
        throw new Error("No hay usuario autenticado");
    }

    const silentRequest = {
        scopes: loginRequest.scopes,
        account: account
    };

    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error) {
        console.warn("Token silencioso falló, intentando redirect:", error);
        if (error instanceof msal.InteractionRequiredAuthError) {
            await msalInstance.acquireTokenRedirect(loginRequest);
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
