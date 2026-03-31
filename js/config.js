const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzWTmJ1ZF3FG3INrVSWPjzXzSbyOV-8BvFnKT5c6Lxqwqy5DyctwoCPIUw4tmBKx_Af/exec";

const APP_CONFIG = Object.freeze({
  webAppUrl: WEB_APP_URL,
  defaultProtection: 1.3,
  initialRows: 6,
  costuraCsvHeaders: [
    "CODIGO",
    "BLOQUE",
    "OPERACIONES",
    "TIEMPOS ESTIMADO",
    "TIPO MAQ",
    "% PROTECCION",
    "TIPO PTA",
    "TIEMPOS MAQ C/PROTECC.",
    "TIEMPOS MANUAL C/PROTECC.",
    "TIEMPOS COTIZACION",
  ],
  corteCsvHeaders: [
    "OPERACIONES",
    "TIEMPOS ESTIMADO CORTE",
    "TIEMPOS ESTIMADO HABILITADO",
    "% PROTECCION",
    "AREA",
    "TIEMPOS CORTE C/PROTECC.",
    "TIEMPOS HAB C/PROTECC.",
    "TIEMPOS COTIZACION"
  ],
  defaultCorteOperations: [
    { operaciones: "TENDER PAÑOS", area: "CORT" },
    { operaciones: "PRETENDIDO", area: "CORT" },
    { operaciones: "TENDIDO TELA", area: "CORT" },
    { operaciones: "CORTE AUTOMATICO", area: "CORT" },
    { operaciones: "NUMERADO", area: "CORT" },
    { operaciones: "Complemento Cuello", area: "HAB" },
    { operaciones: "Complemento Tapeta", area: "HAB" },
    { operaciones: "DESHILADO+ATRAQUE+ ENTALLE CLLO RECTILINEO", area: "HAB" },
    { operaciones: "FUSIONADO + ENTALLE PECHERA X2", area: "HAB" },
    { operaciones: "TRANSFER", area: "HAB" },
    { operaciones: "OTROS", area: "HAB" }
  ],
  defaultAcabadoOperations: [
    "INSPECCION",
    "PLANCHA",
    "HANTAG",
    "MEDIDAS criticas",
    "DOBLADO+ EMBOLSADO",
    "EMBALAJE",
    "PREPARAR HANG TAG",
    "PASAR DETECTOR",
    "COLOCAR PIN DE SEGURIDAD",
    "ESTAMPADO",
    "LAVADO"
  ],
  acabadoCsvHeaders: [
    "OPERACIONES",
    "TIEMPOS ESTIMADO",
    "% PROTECCION",
    "TIEMPOS COTIZACION"
  ],
});
