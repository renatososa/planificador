let token = null;
let isLoggedOut = localStorage.getItem("isLoggedOut") === "true";
let spreadsheetId = localStorage.getItem("sheetId");
let spreadsheetName = localStorage.getItem("sheetName") || "";
let tokenClient;
let registros = [];
let modoEdicion = false;
let registroEnEdicion = null;
const TEMPLATE_SHEET_URL = "";
const SHEET_HEADERS = ["Fecha", "Hora Inicio", "Hora Fin", "Tipo", "Realizado", "Detalle"];
const HORA_INICIO = 7;
const HORA_FIN = 19;
const MINUTOS_INICIO = HORA_INICIO * 60;
const MINUTOS_FIN_INICIO = HORA_FIN * 60;
const MINUTOS_FIN = (HORA_FIN * 60) + 30;
let campoActivo = null;
let ajusteCampoViewportTimer = null;
let calendarioInstancia = null;

// ðŸ•’ LÃ³gica de Tiempo
function minutosAHora(minutos) {
  const horas = Math.floor(minutos / 60);
  const mins = minutos % 60;
  return `${String(horas).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function horaAMinutos(hora) {
  if (!hora) {
    return null;
  }

  const [horas, minutos] = hora.split(":").map(Number);
  return (horas * 60) + minutos;
}

function sumarTreintaMinutosHora(hora) {
  return minutosAHora(Math.min((horaAMinutos(hora) ?? MINUTOS_INICIO) + 30, MINUTOS_FIN));
}

function poblarSelectHoras(select, minimo, maximo) {
  select.innerHTML = "";

  for (let minutos = minimo; minutos <= maximo; minutos += 30) {
    const hora = minutosAHora(minutos);
    const option = document.createElement("option");
    option.value = hora;
    option.text = hora;
    select.appendChild(option);
  }
}

function generarHoras() {
  poblarSelectHoras(document.getElementById("horaInicio"), MINUTOS_INICIO, MINUTOS_FIN_INICIO);
  poblarSelectHoras(document.getElementById("horaFin"), MINUTOS_INICIO + 30, MINUTOS_FIN);
}

function normalizarRealizado(valor) {
  return String(valor || "").trim().toLowerCase() === "si" ? "Si" : "No";
}

function estaRealizado(valor) {
  return normalizarRealizado(valor) === "Si";
}

function normalizarRegistro(fila) {
  if (!fila || !fila.length) {
    return ["", "", "", "", "No", ""];
  }

  if (fila.length >= 6) {
    return [fila[0], fila[1], fila[2], fila[3], normalizarRealizado(fila[4]), fila[5] || ""];
  }

  if (fila.length >= 5) {
    return [fila[0], fila[1], fila[2], fila[3], "No", fila[4] || ""];
  }

  const [fecha, horaInicio, tipo = "", detalle = ""] = fila;
  return [fecha, horaInicio, sumarTreintaMinutosHora(horaInicio), tipo, "No", detalle];
}

function actualizarHoraFinPredeterminada() {
  const horaInicio = document.getElementById("horaInicio").value;
  document.getElementById("horaFin").value = sumarTreintaMinutosHora(horaInicio);
}

function ajustarAlturaDetalle() {
  const campoDetalle = document.getElementById("detalle");
  if (!campoDetalle) {
    return;
  }

  campoDetalle.style.height = "auto";
  campoDetalle.style.height = `${Math.max(campoDetalle.scrollHeight, 56)}px`;
}

function fechaISOAFormatoHoja(fechaISO) {
  if (!fechaISO) {
    return "";
  }

  const [year, month, day] = fechaISO.split("-");
  return `${day}/${month}/${year}`;
}

function formatearFechaVisible(fechaISO) {
  if (!fechaISO) {
    return "";
  }

  const [year, month, day] = fechaISO.split("-").map(Number);
  const fecha = new Date(year, month - 1, day);
  return fecha.toLocaleDateString("es-UY", {
    day: "numeric",
    month: "short",
    year: "numeric"
  });
}

function actualizarDisplayFecha(idCampo, textoVacio = "Seleccionar fecha") {
  const campo = document.getElementById(idCampo);
  const display = document.querySelector(`[data-date-display-for="${idCampo}"]`);
  if (!campo || !display) {
    return;
  }

  if (!campo.value) {
    display.textContent = textoVacio;
    display.classList.add("is-placeholder");
    return;
  }

  display.textContent = formatearFechaVisible(campo.value);
  display.classList.remove("is-placeholder");
}

function abrirSelectorFecha(idCampo) {
  const campo = document.getElementById(idCampo);
  if (!campo || campo.disabled) {
    return;
  }

  campo.focus({ preventScroll: true });

  if (typeof campo.showPicker === "function") {
    campo.showPicker();
    return;
  }

  campo.click();
}

function setFechaHoraActual() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  document.getElementById("fecha").value = `${year}-${month}-${day}`;
  actualizarDisplayFecha("fecha");

  let minutosInicio = (now.getHours() * 60) + (now.getMinutes() < 30 ? 0 : 30);
  minutosInicio = Math.min(Math.max(minutosInicio, MINUTOS_INICIO), MINUTOS_FIN_INICIO);

  document.getElementById("horaInicio").value = minutosAHora(minutosInicio);
  document.getElementById("horaFin").value = minutosAHora(Math.min(minutosInicio + 30, MINUTOS_FIN));
}

function extraerSpreadsheetId(valor) {
  if (!valor) {
    return "";
  }

  const valorLimpio = valor.trim();
  const coincidencia = valorLimpio.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (coincidencia) {
    return coincidencia[1];
  }

  return /^[a-zA-Z0-9-_]{20,}$/.test(valorLimpio) ? valorLimpio : "";
}

function limpiarDatosLocalesHoja() {
  spreadsheetId = null;
  spreadsheetName = "";
  localStorage.removeItem("sheetId");
  localStorage.removeItem("sheetName");
}

function guardarDatosLocalesHoja(id, nombre = "") {
  spreadsheetId = id;
  spreadsheetName = nombre;
  localStorage.setItem("sheetId", id);

  if (nombre) {
    localStorage.setItem("sheetName", nombre);
  } else {
    localStorage.removeItem("sheetName");
  }
}

function limpiarVistasRegistros() {
  registros = [];

  const lista = document.getElementById("lista");
  const resultados = document.getElementById("searchResults");

  if (lista) {
    lista.innerHTML = "";
  }

  if (resultados) {
    resultados.innerHTML = "";
  }

  if (calendarioInstancia) {
    actualizarEventosCalendario(calendarioInstancia, calendarioInstancia.view.type);
  }
}

function mostrarEstadoHoja(mensaje = "", tipo = "") {
  const estado = document.getElementById("sheetLinkStatus");
  if (!estado) {
    return;
  }

  estado.className = "sheet-link-status";

  if (!mensaje) {
    estado.hidden = true;
    estado.textContent = "";
    return;
  }

  if (tipo) {
    estado.classList.add(`is-${tipo}`);
  }

  estado.hidden = false;
  estado.textContent = mensaje;
}

function actualizarPanelHoja() {
  const seccion = document.getElementById("sheetLinkSection");
  const formulario = document.getElementById("sheetLinkForm");
  const panelVinculado = document.getElementById("sheetLinkedPanel");
  const botonVincular = document.getElementById("btnVincularHoja");
  const botonDesvincular = document.getElementById("btnDesvincularHoja");
  const info = document.getElementById("sheetLinkedInfo");

  if (!seccion || !formulario || !panelVinculado || !botonVincular || !botonDesvincular || !info) {
    return;
  }

  if (!token) {
    seccion.style.display = "none";
    return;
  }

  seccion.style.display = "block";

  formulario.style.display = spreadsheetId ? "none" : "flex";
  panelVinculado.style.display = spreadsheetId ? "flex" : "none";
  info.textContent = spreadsheetId ? `Hoja vinculada: ${spreadsheetName || spreadsheetId}` : "";
}

// ðŸ” AutenticaciÃ³n
function initClient() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: "487461126477-5m2p50n3se85n1btppmh5h0vk95nfvnd.apps.googleusercontent.com",
    scope: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets",
    callback: async (response) => {
      token = response.access_token;

      isLoggedOut = false;
      localStorage.setItem("isLoggedOut", "false");
      document.getElementById("loginBtn").style.display = "none";
      document.getElementById("logoutBtn").style.display = "block";
      actualizarPanelHoja();

      if (spreadsheetId) {
        await restaurarHojaVinculada();
      } else {
        limpiarVistasRegistros();
        setCamposHabilitados(false);
        mostrarEstadoHoja("");
      }
    },
    error_callback: () => {
      document.getElementById("loginBtn").style.display = "block";
    }
  });
}

function loginManual() { tokenClient.requestAccessToken({ prompt: "consent" }); }
function autoLogin() { tokenClient.requestAccessToken({ prompt: "" }); }

function logout() {
  token = null;
  localStorage.setItem("isLoggedOut", "true");
  limpiarVistasRegistros();
  document.getElementById("loginBtn").style.display = "block";
  document.getElementById("logoutBtn").style.display = "none";
  setCamposHabilitados(false);
  actualizarPanelHoja();
  mostrarEstadoHoja("");
}

// ðŸ“„ Google Sheets API
async function obtenerMetadatosHoja(id) {
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=spreadsheetId,properties.title`, {
    headers: { Authorization: "Bearer " + token },
  });

  if (!res.ok) {
    throw new Error("No se pudo acceder a la hoja seleccionada.");
  }

  return res.json();
}

async function asegurarCabecerasHoja(id) {
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values/A1:F1`, {
    headers: { Authorization: "Bearer " + token },
  });

  if (!res.ok) {
    throw new Error("No se pudo validar la hoja seleccionada.");
  }

  const data = await res.json();
  const cabecerasActuales = data.values?.[0] || [];
  const yaCoinciden = SHEET_HEADERS.every((valor, indice) => cabecerasActuales[indice] === valor);

  if (yaCoinciden) {
    return;
  }

  const actualizarRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values/A1:F1?valueInputOption=RAW`, {
    method: "PUT",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ values: [SHEET_HEADERS] }),
  });

  if (!actualizarRes.ok) {
    throw new Error("No se pudieron preparar las cabeceras de la hoja.");
  }
}

async function activarHojaVinculada(id, nombreSugerido = "") {
  await asegurarCabecerasHoja(id);
  const metadata = nombreSugerido ? { properties: { title: nombreSugerido } } : await obtenerMetadatosHoja(id);
  guardarDatosLocalesHoja(id, metadata.properties?.title || nombreSugerido || "");
  actualizarPanelHoja();
  setCamposHabilitados(true);
  await cargar();

  if (document.getElementById("calendarView").style.display !== "none") {
    programarRenderCalendario();
  }
}

async function restaurarHojaVinculada() {
  try {
    await activarHojaVinculada(spreadsheetId, spreadsheetName);
  } catch (error) {
    limpiarDatosLocalesHoja();
    limpiarVistasRegistros();
    setCamposHabilitados(false);
    actualizarPanelHoja();
    mostrarEstadoHoja("No se pudo acceder a la hoja vinculada. Pega de nuevo la URL de tu copia.", "error");
  }
}

async function vincularHojaDesdeInput() {
  if (!token) {
    mostrarEstadoHoja("Primero inicia sesi\u00f3n con Google.", "error");
    return;
  }

  const input = document.getElementById("sheetUrl");
  const id = extraerSpreadsheetId(input.value);

  if (!id) {
    mostrarEstadoHoja("Pega una URL v\u00e1lida de Google Sheets o un ID de hoja v\u00e1lido.", "error");
    return;
  }

  try {
    await activarHojaVinculada(id);
    input.value = "";
    mostrarEstadoHoja("Hoja vinculada correctamente.", "success");
  } catch (error) {
    mostrarEstadoHoja(error.message || "No se pudo vincular la hoja indicada.", "error");
  }
}

function desvincularHoja() {
  limpiarDatosLocalesHoja();
  limpiarVistasRegistros();
  setCamposHabilitados(false);
  mostrarHome();
  actualizarPanelHoja();
  mostrarEstadoHoja("La hoja fue desvinculada de este dispositivo.", "success");
}

async function guardar() {
  if (!token || !spreadsheetId) {
    mostrarEstadoHoja("Primero vincula una hoja de Google Sheets.", "error");
    return;
  }

  const fechaInput = document.getElementById("fecha").value;
  const horaInicio = document.getElementById("horaInicio").value;
  const horaFin = document.getElementById("horaFin").value;
  const tipo = document.getElementById("tipo").value;
  const realizado = document.getElementById("realizado").checked ? "Si" : "No";
  const detalle = document.getElementById("detalle").value;

  if (!fechaInput || !horaInicio || !horaFin || !detalle) {
    alert("Por favor, completa todos los campos.");
    return;
  }

  if ((horaAMinutos(horaFin) ?? 0) <= (horaAMinutos(horaInicio) ?? 0)) {
    alert("La hora de fin debe ser posterior a la hora de inicio.");
    return;
  }

  const fechaFormateada = fechaISOAFormatoHoja(fechaInput);
  if (modoEdicion && registroEnEdicion) {
    // Modo ediciÃ³n: actualizar el registro existente
    let rowIndex = -1;
    for (let i = 0; i < registros.length; i++) {
      if (
        registros[i][0] === registroEnEdicion[0] &&
        registros[i][1] === registroEnEdicion[1] &&
        registros[i][2] === registroEnEdicion[2] &&
        registros[i][3] === registroEnEdicion[3] &&
        registros[i][4] === registroEnEdicion[4] &&
        registros[i][5] === registroEnEdicion[5]
      ) {
        rowIndex = i + 2; // +2 porque fila 1 es encabezado
        break;
      }
    }

    if (rowIndex === -1) {
      alert("No se encontrÃ³ el registro para actualizar.");
      return;
    }

    const range = `A${rowIndex}:F${rowIndex}`;
    const values = [[fechaFormateada, horaInicio, horaFin, tipo, realizado, detalle]];

    await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=RAW`, {
      method: "PUT",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ values }),
    });

    modoEdicion = false;
    registroEnEdicion = null;
    document.getElementById("btnGuardar").textContent = "Guardar Registro";
    document.getElementById("detalle").value = "";
    ajustarAlturaDetalle();
  } else {
    // Modo crear: verificar duplicados y crear nuevo registro
    const registroDuplicado = registros.some(fila => fila[0] === fechaFormateada && fila[1] === horaInicio);
    if (registroDuplicado) {
      alert("Ya existe un registro con la misma fecha y hora de inicio. Por favor, selecciona una hora diferente.");
      return;
    }

    const values = [[fechaFormateada, horaInicio, horaFin, tipo, realizado, detalle]];

    await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A2:F:append?valueInputOption=RAW`, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ values }),
    });

    document.getElementById("detalle").value = "";
    ajustarAlturaDetalle();
  }

  await cargar();

  if (document.getElementById('calendarView').style.display !== 'none') {
    programarRenderCalendario();
  }
}

function crearTarjetaRegistro(fila) {
  const [fecha, horaInicio, horaFin, tipoRegistro, realizadoRegistro, detalleRegistro] = fila;
  const div = document.createElement("div");
  div.className = `card card-tipo-${normalizarTipoClase(tipoRegistro)}`;
  if (estaRealizado(realizadoRegistro)) {
    div.classList.add("card-realizado");
  }

  const content = document.createElement("div");
  content.className = "card-content";
  const heading = document.createElement("b");
  heading.textContent = `${fecha} - ${horaInicio} a ${horaFin}`;
  const tipo = document.createElement("small");
  tipo.textContent = estaRealizado(realizadoRegistro) ? `${tipoRegistro} · Realizado` : (tipoRegistro || "");
  const detalle = document.createElement("p");
  detalle.textContent = detalleRegistro || "";

  content.appendChild(heading);
  content.appendChild(tipo);
  content.appendChild(detalle);

  const actions = document.createElement("div");
  actions.className = "card-actions";

  const editBtn = document.createElement("button");
  editBtn.className = "btn-edit";
  editBtn.type = "button";
  editBtn.setAttribute("aria-label", "Editar registro");
  editBtn.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8" d="M5.75 5.25h5.25M5.75 5.25a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-5.25"/><path fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8" d="M13.9 5.3l4.8 4.8M10.2 17.1l1.3-3.9 7.3-7.3a1.35 1.35 0 0 1 1.9 0l1.15 1.15a1.35 1.35 0 0 1 0 1.9l-7.3 7.3-4.3.85z"/></svg>';
  editBtn.onclick = () => editarRegistro(fila);

  const deleteBtn = document.createElement("button");
  deleteBtn.className = "btn-delete";
  deleteBtn.type = "button";
  deleteBtn.textContent = "\u2715";
  deleteBtn.onclick = () => borrarRegistro(fila);

  actions.appendChild(editBtn);
  actions.appendChild(deleteBtn);

  div.appendChild(content);
  div.appendChild(actions);

  return div;
}

async function cargar() {
  if (!token || !spreadsheetId) {
    limpiarVistasRegistros();
    return;
  }

  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A2:F100`, {
    headers: { Authorization: "Bearer " + token },
  });
  const data = await res.json();
  const lista = document.getElementById("lista");
  lista.innerHTML = "";
  registros = (data.values || []).map(normalizarRegistro);
  if (!data.values) return;

  const fechaSeleccionada = document.getElementById("fecha").value;
  const fechaFormateada = fechaISOAFormatoHoja(fechaSeleccionada);

  const registrosFiltrados = registros
    .filter(fila => fila[0] === fechaFormateada)
    .reverse();

  if (!registrosFiltrados.length) {
    lista.innerHTML = '<p style="color:#555; font-size:.95rem;">No hay registros para la fecha seleccionada.</p>';
    return;
  }

  registrosFiltrados.forEach((fila) => {
    lista.appendChild(crearTarjetaRegistro(fila));
  });
}

function setCamposHabilitados(habilitado) {
  document.querySelectorAll("[data-requires-sheet=\"true\"]").forEach(el => el.disabled = !habilitado);
}

function esCampoEditable(elemento) {
  return elemento instanceof HTMLElement && elemento.matches('input, select, textarea');
}

function desplazarCampoVisible(elemento, opciones = {}) {
  if (!esCampoEditable(elemento)) {
    return;
  }

  const { centrar = false, suave = false } = opciones;
  const rect = elemento.getBoundingClientRect();
  const viewportHeight = window.visualViewport?.height ?? window.innerHeight;
  const margenSuperior = 24;
  const margenInferior = 24;
  const fueraPorArriba = rect.top < margenSuperior;
  const fueraPorAbajo = rect.bottom > (viewportHeight - margenInferior);

  if (!centrar && !fueraPorArriba && !fueraPorAbajo) {
    return;
  }

  elemento.scrollIntoView({
    behavior: suave ? 'smooth' : 'auto',
    block: centrar ? 'center' : 'nearest',
    inline: 'nearest'
  });
}

function asegurarCampoVisible(elemento) {
  if (!esCampoEditable(elemento)) {
    return;
  }

  campoActivo = elemento;

  if (ajusteCampoViewportTimer) {
    window.clearTimeout(ajusteCampoViewportTimer);
  }

  ajusteCampoViewportTimer = window.setTimeout(() => {
    if (campoActivo !== elemento) {
      return;
    }

    desplazarCampoVisible(elemento, { centrar: true, suave: true });
  }, 220);
}

function inicializarAjusteCamposMovil() {
  document.addEventListener('focusin', (event) => {
    if (esCampoEditable(event.target)) {
      asegurarCampoVisible(event.target);
    }
  });

  document.addEventListener('focusout', (event) => {
    if (campoActivo === event.target) {
      campoActivo = null;
    }

    if (ajusteCampoViewportTimer) {
      window.clearTimeout(ajusteCampoViewportTimer);
      ajusteCampoViewportTimer = null;
    }
  });

  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', () => {
      if (ajusteCampoViewportTimer) {
        window.clearTimeout(ajusteCampoViewportTimer);
      }

      ajusteCampoViewportTimer = window.setTimeout(() => {
        if (!campoActivo) {
          return;
        }

        desplazarCampoVisible(campoActivo, { centrar: false, suave: false });
      }, 120);
    });

    window.visualViewport.addEventListener('scroll', () => {
      if (campoActivo) {
        desplazarCampoVisible(campoActivo, { centrar: false, suave: false });
      }
    });
  }
}

function mostrarHome() {
  document.getElementById('mainView').style.display = 'block';
  document.getElementById('calendarView').style.display = 'none';
  document.getElementById('searchView').style.display = 'none';
  document.getElementById('navHome').classList.add('active');
  document.getElementById('navCalendar').classList.remove('active');
  document.getElementById('navSearch').classList.remove('active');
}

function programarRenderCalendario() {
  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => {
      inicializarCalendario();
    });
  });
}

function mostrarCalendario() {
  document.getElementById('mainView').style.display = 'none';
  document.getElementById('calendarView').style.display = 'block';
  document.getElementById('searchView').style.display = 'none';
  document.getElementById('navHome').classList.remove('active');
  document.getElementById('navCalendar').classList.add('active');
  document.getElementById('navSearch').classList.remove('active');

  if (registros.length === 0 && token && spreadsheetId) {
    cargar().then(programarRenderCalendario);
  } else {
    programarRenderCalendario();
  }
}

function mostrarBuscar() {
  document.getElementById('mainView').style.display = 'none';
  document.getElementById('calendarView').style.display = 'none';
  document.getElementById('searchView').style.display = 'block';
  document.getElementById('navHome').classList.remove('active');
  document.getElementById('navCalendar').classList.remove('active');
  document.getElementById('navSearch').classList.add('active');
  
  // Inicializar horas en el selector de bÃºsqueda
  inicializarHorasBusqueda();
  
  document.getElementById('searchFecha').value = '';
  actualizarDisplayFecha('searchFecha', 'Todas las fechas');
  document.getElementById('searchHora').value = '';
  document.getElementById('searchKeyword').value = '';
  document.getElementById('searchResults').innerHTML = '';
}

function inicializarHorasBusqueda() {
  const select = document.getElementById("searchHora");
  // Limpiar opciones existentes excepto la primera
  while (select.children.length > 1) {
    select.removeChild(select.lastChild);
  }
  
  // Agregar todas las horas disponibles
  for (let minutos = MINUTOS_INICIO; minutos <= MINUTOS_FIN; minutos += 30) {
    const hora = minutosAHora(minutos);
    const option = document.createElement("option");
    option.value = hora;
    option.text = hora;
    select.appendChild(option);
  }
}

function limpiarBusqueda() {
  document.getElementById('searchFecha').value = '';
  actualizarDisplayFecha('searchFecha', 'Todas las fechas');
  document.getElementById('searchHora').value = '';
  document.getElementById('searchKeyword').value = '';
  document.getElementById('searchResults').innerHTML = '';
}

function normalizarTipoClase(tipo) {
  return tipo
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '-');
}

function construirFechaHoraISO(fecha, hora) {
  const [dia, mes, anio] = fecha.split('/');
  return `${anio}-${mes.padStart(2, '0')}-${dia.padStart(2, '0')}T${hora}:00`;
}

function construirEventosDetallados() {
  return registros
    .filter(registro => registro[0] && registro[1] && registro[2] && registro[3])
    .map(registro => {
      const [fecha, horaInicio, horaFin, tipo, realizado, detalle = ''] = registro;

      return {
        title: detalle ? `${tipo}: ${detalle}` : tipo,
        start: construirFechaHoraISO(fecha, horaInicio),
        end: construirFechaHoraISO(fecha, horaFin),
        allDay: false,
        classNames: [
          `event-${normalizarTipoClase(tipo)}`,
          ...(estaRealizado(realizado) ? ['event-realizado'] : [])
        ],
        extendedProps: {
          agrupado: false,
          tipo,
          realizado,
          detalle,
          fecha,
          horaInicio,
          horaFin
        }
      };
    });
}

function construirEventosAgrupadosMes() {
  const acumulado = {};

  registros.forEach(registro => {
    const [fecha, , , tipo, realizado] = registro;
    if (!fecha || !tipo) {
      return;
    }

    const clave = `${fecha}|${tipo}|${normalizarRealizado(realizado)}`;
    if (!acumulado[clave]) {
      acumulado[clave] = { fecha, tipo, realizado: normalizarRealizado(realizado), cantidad: 0 };
    }

    acumulado[clave].cantidad += 1;
  });

  return Object.values(acumulado).map(({ fecha, tipo, realizado, cantidad }) => ({
    title: String(cantidad),
    start: construirFechaHoraISO(fecha, '00:00').slice(0, 10),
    allDay: true,
    classNames: [
      `event-${normalizarTipoClase(tipo)}`,
      ...(estaRealizado(realizado) ? ['event-realizado'] : [])
    ],
    extendedProps: {
      agrupado: true,
      tipo,
      realizado,
      cantidad,
      fecha
    }
  }));
}

function actualizarEventosCalendario(calendar, tipoVista) {
  const eventos = tipoVista === 'dayGridMonth'
    ? construirEventosAgrupadosMes()
    : construirEventosDetallados();

  calendar.getEventSources().forEach(source => source.remove());
  calendar.addEventSource(eventos);
}

function renderizarLeyendaCalendario(calendarEl) {
  const toolbar = calendarEl.querySelector('.fc-header-toolbar');
  const leftChunk = toolbar?.querySelector('.fc-toolbar-chunk');

  if (!leftChunk) {
    return;
  }

  const prevButton = leftChunk.querySelector('.fc-prev-button');
  const nextButton = leftChunk.querySelector('.fc-next-button');
  const todayButton = leftChunk.querySelector('.fc-today-button');

  if (prevButton && nextButton && todayButton && !leftChunk.querySelector('.calendar-nav-group')) {
    const navGroup = document.createElement('div');
    navGroup.className = 'calendar-nav-group';
    const buttonGroup = prevButton.parentElement;

    if (buttonGroup?.classList.contains('fc-button-group')) {
      navGroup.appendChild(buttonGroup);
    } else {
      navGroup.appendChild(prevButton);
      navGroup.appendChild(nextButton);
    }

    navGroup.appendChild(todayButton);
    leftChunk.insertBefore(navGroup, leftChunk.firstChild);
  }

  if (leftChunk.querySelector('.calendar-legend')) {
    return;
  }

  const items = [
    { label: 'Informe', className: 'event-informe' },
    { label: 'Visita', className: 'event-visita' },
    { label: 'Entrevista', className: 'event-entrevista' },
    { label: 'Reunión', className: 'event-reunion' },
    { label: 'Contactar', className: 'event-contactar' },
    { label: 'Entrega', className: 'event-entrega' },
    { label: 'Otro', className: 'event-otro' }
  ];

  const legend = document.createElement('div');
  legend.className = 'calendar-legend';

  items.forEach(item => {
    const legendItem = document.createElement('div');
    legendItem.className = 'calendar-legend-item';

    const swatch = document.createElement('span');
    swatch.className = `calendar-legend-swatch ${item.className}`;

    const text = document.createElement('span');
    text.className = 'calendar-legend-text';
    text.textContent = item.label;

    legendItem.appendChild(swatch);
    legendItem.appendChild(text);
    legend.appendChild(legendItem);
  });

  leftChunk.appendChild(legend);
}

function inicializarCalendario() {
  const calendarEl = document.getElementById('calendar');

  if (calendarioInstancia) {
    calendarioInstancia.destroy();
    calendarioInstancia = null;
  }

  calendarEl.innerHTML = '';

  calendarioInstancia = new FullCalendar.Calendar(calendarEl, {
    locale: 'es',
    firstDay: 1,
    initialView: 'dayGridMonth',
    height: 'auto',
    contentHeight: 'auto',
    expandRows: true,
    stickyHeaderDates: false,
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,timeGridWeek,timeGridDay'
    },
    buttonText: {
      today: 'HOY',
      month: 'MES',
      week: 'SEMANA',
      day: 'DÃA'
    },
    monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
    monthNamesShort: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
    dayNames: ['Domingo', 'Lunes', 'Martes', 'MiÃ©rcoles', 'Jueves', 'Viernes', 'SÃ¡bado'],
    dayNamesShort: ['Dom', 'Lun', 'Mar', 'MiÃ©', 'Jue', 'Vie', 'SÃ¡b'],
    buttonText: {
      today: 'Hoy',
      month: 'Mes',
      week: 'Semana',
      day: 'D\u00eda'
    },
    allDayText: 'Hora',
    slotMinTime: '07:00:00',
    slotMaxTime: '19:30:00',
    slotDuration: '00:30:00',
    slotLabelInterval: '00:30:00',
    slotLabelFormat: {
      hour: 'numeric',
      minute: '2-digit',
      hour12: false
    },
    eventTimeFormat: {
      hour: 'numeric',
      minute: '2-digit',
      hour12: false
    },
    scrollTime: '07:00:00',
    events: construirEventosAgrupadosMes(),
    datesSet: function(info) {
      actualizarEventosCalendario(info.view.calendar, info.view.type);
      renderizarLeyendaCalendario(calendarEl);
    },
    eventClick: function(info) {
      if (info.event.extendedProps.agrupado) {
        const tipo = info.event.extendedProps.tipo;
        const realizado = info.event.extendedProps.realizado;
        const cantidad = info.event.extendedProps.cantidad;
        const fecha = info.event.extendedProps.fecha;

        alert(`Fecha: ${fecha}\nTipo: ${tipo}\nRealizado: ${estaRealizado(realizado) ? 'Si' : 'No'}\nCantidad: ${cantidad} registro(s)`);
        return;
      }

      const tipo = info.event.extendedProps.tipo;
      const realizado = info.event.extendedProps.realizado;
      const detalle = info.event.extendedProps.detalle;
      const fecha = info.event.extendedProps.fecha;
      const horaInicio = info.event.extendedProps.horaInicio;
      const horaFin = info.event.extendedProps.horaFin;

      alert(`Fecha: ${fecha}\nHora: ${horaInicio} a ${horaFin}\nTipo: ${tipo}\nRealizado: ${estaRealizado(realizado) ? 'Si' : 'No'}\nDetalle: ${detalle || 'Sin detalle'}`);
    },
    dayMaxEvents: 3,
    moreLinkClick: 'popover',
    showNonCurrentDates: false,
    fixedWeekCount: false
  });

  calendarioInstancia.render();
  calendarioInstancia.updateSize();
  renderizarLeyendaCalendario(calendarEl);
}

function buscar() {
  const fechaBusqueda = document.getElementById('searchFecha').value;
  const horaBusqueda = document.getElementById('searchHora').value;
  const keywordBusqueda = document.getElementById('searchKeyword').value.trim().toLowerCase();
  
  const dest = document.getElementById('searchResults');
  dest.innerHTML = '';
  
  let matches = registros;
  
  if (fechaBusqueda) {
    const fechaFormateada = fechaISOAFormatoHoja(fechaBusqueda);
    matches = matches.filter(fila => fila[0] === fechaFormateada);
  }
  
  if (horaBusqueda) {
    const horaBusquedaMinutos = horaAMinutos(horaBusqueda) ?? 0;
    matches = matches.filter(fila => {
      const inicio = horaAMinutos(fila[1]) ?? 0;
      const fin = horaAMinutos(fila[2]) ?? inicio;
      return horaBusquedaMinutos >= inicio && horaBusquedaMinutos <= fin;
    });
  }
  
  if (keywordBusqueda) {
    matches = matches.filter(fila => 
      fila[3].toLowerCase().includes(keywordBusqueda) || 
      fila[4].toLowerCase().includes(keywordBusqueda) ||
      fila[5].toLowerCase().includes(keywordBusqueda)
    );
  }
  
  if (!fechaBusqueda && !horaBusqueda && !keywordBusqueda) {
    dest.innerHTML = '<p style="color:#555; font-size:.95rem;">Especifica al menos un criterio de b\u00fasqueda.</p>';
    return;
  }
  
  if (!matches.length) {
    dest.innerHTML = '<p style="color:#555; font-size:.95rem;">No se encontraron registros que coincidan con los criterios.</p>';
    return;
  }
  
  matches.sort((a, b) => {
    const fechaA = a[0].split('/').reverse().join('');
    const fechaB = b[0].split('/').reverse().join('');
    if (fechaA !== fechaB) return fechaB.localeCompare(fechaA);
    return b[1].localeCompare(a[1]);
  });

  matches.forEach(fila => {
    dest.appendChild(crearTarjetaRegistro(fila));
  });
}

async function borrarRegistro(fila) {
  if (!confirm(`\u00bfEst\u00e1s seguro de que deseas borrar el registro de ${fila[0]} de ${fila[1]} a ${fila[2]}?`)) {
    return;
  }

  // Encontrar el Ã­ndice del registro en Google Sheets (offset de 2 porque la primera fila es encabezado)
  const indices = [];
  for (let i = 0; i < registros.length; i++) {
    if (
      registros[i][0] === fila[0] &&
      registros[i][1] === fila[1] &&
      registros[i][2] === fila[2] &&
      registros[i][3] === fila[3] &&
      registros[i][4] === fila[4] &&
      registros[i][5] === fila[5]
    ) {
      indices.push(i + 2); // +2 porque la fila 1 es encabezado y array es 0-indexed
      break;
    }
  }

  if (indices.length === 0) {
    alert("No se encontrÃ³ el registro para borrar.");
    return;
  }

  // Usar batchUpdate para eliminar filas de Google Sheets
  const deleteRequest = {
    requests: [{
      deleteDimension: {
        range: {
          sheetId: 0,
          dimension: "ROWS",
          startIndex: indices[0] - 1,
          endIndex: indices[0]
        }
      }
    }]
  };

  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify(deleteRequest),
  });

  await cargar();
  if (document.getElementById('calendarView').style.display !== 'none') {
    programarRenderCalendario();
  }
}

function editarRegistro(fila) {
  // Cargar los datos del registro en los campos del formulario
  modoEdicion = true;
  registroEnEdicion = fila;

  // Convertir fecha de dd/mm/yyyy a yyyy-mm-dd
  const [day, month, year] = fila[0].split("/");
  const fechaInput = `${year}-${month}-${day}`;

  document.getElementById("fecha").value = fechaInput;
  actualizarDisplayFecha("fecha");
  document.getElementById("horaInicio").value = fila[1];
  document.getElementById("horaFin").value = fila[2];
  document.getElementById("tipo").value = fila[3];
  document.getElementById("realizado").checked = estaRealizado(fila[4]);
  document.getElementById("detalle").value = fila[5];
  ajustarAlturaDetalle();

  // Cambiar el texto del botÃ³n y mostrar cancelar
  document.getElementById("btnGuardar").textContent = "Actualizar Registro";
  document.getElementById("btnCancelar").style.display = "block";

  // Desplazarse hacia el formulario
  document.getElementById("mainView").scrollIntoView({ behavior: "smooth", block: "start" });
}

function cancelarEdicion() {
  // Limpiar los campos y volver al modo guardar
  modoEdicion = false;
  registroEnEdicion = null;

  document.getElementById("fecha").value = "";
  actualizarDisplayFecha("fecha");
  document.getElementById("horaInicio").value = "";
  document.getElementById("horaFin").value = "";
  document.getElementById("tipo").value = "Informe";
  document.getElementById("realizado").checked = false;
  document.getElementById("detalle").value = "";
  ajustarAlturaDetalle();

  // Restablecer el botÃ³n a su estado original
  document.getElementById("btnGuardar").textContent = "Guardar Registro";
  document.getElementById("btnCancelar").style.display = "none";

  // Establecer la fecha y hora actual
  setFechaHoraActual();
}

function registrarEventosUI() {
  document.getElementById("loginBtn").addEventListener("click", loginManual);
  document.getElementById("logoutBtn").addEventListener("click", logout);
  document.getElementById("btnVincularHoja").addEventListener("click", vincularHojaDesdeInput);
  document.getElementById("btnDesvincularHoja").addEventListener("click", desvincularHoja);
  document.getElementById("btnGuardar").addEventListener("click", guardar);
  document.getElementById("btnCancelar").addEventListener("click", cancelarEdicion);
  document.getElementById("horaInicio").addEventListener("change", actualizarHoraFinPredeterminada);
  document.getElementById("detalle").addEventListener("input", ajustarAlturaDetalle);
  document.getElementById("btnBuscar").addEventListener("click", buscar);
  document.getElementById("btnLimpiarBusqueda").addEventListener("click", limpiarBusqueda);
  document.getElementById("btnVolverBusqueda").addEventListener("click", mostrarHome);
  document.getElementById("navHome").addEventListener("click", mostrarHome);
  document.getElementById("navCalendar").addEventListener("click", mostrarCalendario);
  document.getElementById("navSearch").addEventListener("click", mostrarBuscar);
  document.getElementById("fecha").addEventListener("change", () => actualizarDisplayFecha("fecha"));
  document.getElementById("searchFecha").addEventListener("change", () => actualizarDisplayFecha("searchFecha", "Todas las fechas"));
  document.querySelectorAll("[data-date-shell-for]").forEach((shell) => {
    shell.addEventListener("click", (event) => {
      const idCampo = shell.getAttribute("data-date-shell-for");
      if (!idCampo) {
        return;
      }

      event.preventDefault();
      abrirSelectorFecha(idCampo);
    });
  });
}

function estaEnModoStandalone() {
  return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
}

function esDispositivoIOS() {
  return /iPhone|iPad|iPod/i.test(window.navigator.userAgent);
}

function actualizarEstadoPWA() {
  const estado = document.getElementById('appStatus');
  if (!estado) {
    return;
  }

  const enModoStandalone = estaEnModoStandalone();
  document.body.classList.toggle('app-standalone', enModoStandalone);

  estado.className = 'app-status';
  estado.hidden = false;

  if (!navigator.onLine) {
    estado.classList.add('is-offline');
    estado.textContent = 'Sin conexión. Esta app requiere internet para funcionar.';
    return;
  }

  if (enModoStandalone) {
    estado.classList.add('is-installed');
    estado.textContent = 'Estás usando la app instalada.';
    return;
  }

  if (esDispositivoIOS()) {
    estado.textContent = 'En iPhone puedes instalarla desde Compartir > Añadir a pantalla de inicio.';
    return;
  }

  estado.hidden = true;
  estado.textContent = '';
}

function registrarPWA() {
  actualizarEstadoPWA();
  window.addEventListener('online', actualizarEstadoPWA);
  window.addEventListener('offline', actualizarEstadoPWA);
  window.addEventListener('appinstalled', actualizarEstadoPWA);

  const displayMode = window.matchMedia('(display-mode: standalone)');
  if (typeof displayMode.addEventListener === 'function') {
    displayMode.addEventListener('change', actualizarEstadoPWA);
  } else if (typeof displayMode.addListener === 'function') {
    displayMode.addListener(actualizarEstadoPWA);
  }

  if (!('serviceWorker' in navigator)) {
    return;
  }

  navigator.serviceWorker.register('./service-worker.js').catch((error) => {
    console.error('No se pudo registrar el service worker.', error);
  });
}

window.onload = () => {
  generarHoras();
  setFechaHoraActual();
  actualizarDisplayFecha("searchFecha", "Todas las fechas");
  ajustarAlturaDetalle();
  initClient();
  registrarEventosUI();
  registrarPWA();
  inicializarAjusteCamposMovil();
  setCamposHabilitados(false);
  document.getElementById("fecha").addEventListener("change", () => {
    if (token && spreadsheetId) {
      cargar();
    }
  });
  if (!isLoggedOut) autoLogin();
  else document.getElementById("loginBtn").style.display = "block";
  mostrarHome();
};

