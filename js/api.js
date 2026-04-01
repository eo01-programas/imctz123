(function () {
  function isConfigured() {
    return Boolean(APP_CONFIG.webAppUrl) && !APP_CONFIG.webAppUrl.includes("PEGA_AQUI");
  }

  function buildUrl(params) {
    const url = new URL(APP_CONFIG.webAppUrl);
    Object.entries(params).forEach(([key, value]) => {
      url.searchParams.set(key, value);
    });
    return url.toString();
  }

  function jsonpRequest(params) {
    return new Promise((resolve, reject) => {
      const callbackName = `costuraCallback_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
      const script = document.createElement("script");
      const timeoutId = window.setTimeout(() => {
        cleanup();
        reject(new Error("La API demoró demasiado en responder."));
      }, 15000);

      function cleanup() {
        window.clearTimeout(timeoutId);
        if (script.parentNode) {
          script.parentNode.removeChild(script);
        }
        delete window[callbackName];
      }

      window[callbackName] = (payload) => {
        cleanup();
        resolve(payload);
      };

      script.onerror = () => {
        cleanup();
        reject(new Error("No se pudo conectar con Google Apps Script."));
      };

      script.src = buildUrl({
        ...params,
        callback: callbackName,
        _: Date.now().toString(),
      });

      document.body.appendChild(script);
    });
  }

  async function fetchCatalog() {
    return jsonpRequest({ action: "catalog" });
  }

  async function fetchBasedatosSheet() {
    return jsonpRequest({ action: "basedatosSheet" });
  }

  async function searchByProto(proto) {
    return jsonpRequest({ action: "searchProto", proto });
  }

  async function saveCotizacion(payload) {
    const body = new URLSearchParams({
      payload: JSON.stringify(payload),
    });

    try {
      const response = await fetch(APP_CONFIG.webAppUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        body,
      });

      const text = await response.text();

      if (!text) {
        return {
          success: response.ok,
          message: response.ok ? "Registro enviado." : "No se recibió respuesta de la API.",
        };
      }

      return JSON.parse(text);
    } catch (error) {
      await fetch(APP_CONFIG.webAppUrl, {
        method: "POST",
        mode: "no-cors",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        body,
      });

      return {
        success: true,
        message: "Registro enviado. La respuesta no pudo leerse por restricciones del navegador.",
      };
    }
  }

  async function saveBasedatosRow(payload) {
    const body = new URLSearchParams({
      payload: JSON.stringify({
        action: "SAVE_BASEDATOS_ROW",
        ...payload,
      }),
    });

    try {
      const response = await fetch(APP_CONFIG.webAppUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        body,
      });

      const text = await response.text();

      if (!text) {
        return {
          success: response.ok,
          message: response.ok ? "Fila enviada." : "No se recibió respuesta de la API.",
        };
      }

      return JSON.parse(text);
    } catch (error) {
      await fetch(APP_CONFIG.webAppUrl, {
        method: "POST",
        mode: "no-cors",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        body,
      });

      return {
        success: true,
        message: "Fila enviada. La respuesta no pudo leerse por restricciones del navegador.",
      };
    }
  }

  window.AppsScriptAPI = {
    isConfigured,
    fetchCatalog,
    fetchBasedatosSheet,
    searchByProto,
    saveCotizacion,
    saveBasedatosRow,
  };
})();
