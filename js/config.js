const msalConfig = {
    auth: {
        clientId: "f5c0aa2-c940-42c1-9a18-c0c3b96a58d5",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://asalazar1989.github.io/portal-riss/"
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
    }
};

const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite", "Files.ReadWrite.All"]
};

const excelConfig = {
    siteUrl: "rangelrehabilitacion.sharepoint.com",
    sitePath: "/sites/CentrodeServiciosRangelRHB",
    fileName: "Productividad RISS_Version12022025.xlsx",
    fileId: "BBC2BD08-61EE-45EF-BC83-034F5B1C6157",
    sheetName: "Casos RISS",
    
    columns: {
        idCaso: 0,
        paciente: 1,
        servicio: 2,
        fechaSolicitud: 3,
        fechaVencimiento: 4,
        telefono: 5,
        direccion: 6,
        observacionesRISS: 7,
        asignadoA: 8,
        fechaAsignacion: 9,
        resolutor: 10,
        fechaCierre: 11,
        observaciones: 12,
        estado: 13,
        llamado1: 14,
        llamado2: 15,
        llamado3: 16
    },
    
    estadosPermitidos: [
        "PENDIENTE",
        "EN GESTIÓN",
        "CONTACTADO",
        "NO CONTACTADO",
        "PROGRAMADO",
        "CERRADO",
        "CERRADO POR INTENTOS"
    ],
    
    resultadosLlamado: [
        "Contactado exitosamente",
        "No contesta",
        "Número erróneo",
        "Buzón de voz",
        "Solicita llamar más tarde",
        "Otro"
    ]
};

const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphSitesEndpoint: "https://graph.microsoft.com/v1.0/sites",
    graphFilesEndpoint: "https://graph.microsoft.com/v1.0/me/drive/items"
};
