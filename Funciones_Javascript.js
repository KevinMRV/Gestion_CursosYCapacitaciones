/* =========================================================
   Gestión de Capacitaciones - Funciones_Javascript.js
   ========================================================= */
"use strict";

/* ---------- Constantes y datos base ---------- */

const CLAVE_STORAGE = "Funcionarios_Data";
const MENSAJE_SIN_DESCRIPCION = "Sin descripción disponible para esta capacitación.";

const Descripciones_Cursos = {
    "Declaración Renta": "Capacitación sobre las normativas, leyes vigentes y procedimiento paso a paso para la declaración anual de impuesto a la renta.",
    "Fiscalización Base": "Taller práctico sobre procesos de auditoría, revisión documental y fiscalización tributaria en terreno.",
    "Atención al Contribuyente": "Módulo de formación orientado a la atención al cliente, resolución de conflictos y trámites presenciales o digitales."
};

// Las fechas se guardan en formato ISO (aaaa-mm-dd) y se muestran como dd/mm/aaaa.
const DatosIniciales = [
    {
        ID: "1",
        Codigo: "ACT-2026-01",
        Curso: "Declaración Renta",
        Fecha: "2026-10-21",
        Examen: "2026-10-28",
        Modalidad: "Presencial",
        Lugar: "Punta Arenas",
        TipoInscripcion: "Directa",
        Origen: "Nacional",
        DireccionRegional: "XII DIRECCIÓN REGIONAL PUNTA ARENAS",
        Departamento: "Fiscalización",
        Funcionario: "Oscar Ortega",
        Observaciones: "Sin observaciones"
    }
];

// Una sola definición de columnas sirve para la tabla, el orden y la exportación.
const COLUMNAS = [
    { campo: "Codigo", titulo: "Código Versión" },
    { campo: "Curso", titulo: "Actividad" },
    { campo: "Fecha", titulo: "Inicio", tipo: "fecha" },
    { campo: "Examen", titulo: "Término", tipo: "fecha" },
    { campo: "Estado", titulo: "Estado" },
    { campo: "Modalidad", titulo: "Modalidad" },
    { campo: "Lugar", titulo: "Lugar" },
    { campo: "TipoInscripcion", titulo: "Tipo de Inscripción" },
    { campo: "Origen", titulo: "Origen" },
    { campo: "DireccionRegional", titulo: "Dirección Regional" },
    { campo: "Departamento", titulo: "Departamento" },
    { campo: "Funcionario", titulo: "Funcionario Seleccionado" },
    { campo: "Observaciones", titulo: "Observaciones" }
];

const ETIQUETAS_ESTADO = {
    "proxima": "Próxima",
    "en-curso": "En curso",
    "finalizada": "Finalizada",
    "sin-fecha": "Sin fecha"
};

const CamposFormulario = {
    Codigo: "Nuevo_Codigo",
    Curso: "Nuevo_Curso",
    Fecha: "Nueva_Fecha",
    Examen: "Nuevo_Examen",
    Modalidad: "Nueva_Modalidad",
    Lugar: "Nuevo_Lugar",
    TipoInscripcion: "Nuevo_TipoInscripcion",
    Origen: "Nuevo_Origen",
    DireccionRegional: "Nueva_DireccionRegional",
    Departamento: "Nuevo_Departamento",
    Funcionario: "Nuevo_Funcionario",
    Observaciones: "Nueva_Observacion"
};

const MapeoColumnasExcel = {
    "Código Versión": "Codigo",
    "Actividad": "Curso",
    "Inicio": "Fecha",
    "Término": "Examen",
    "Modalidad": "Modalidad",
    "Lugar": "Lugar",
    "Tipo de Inscripción": "TipoInscripcion",
    "Origen": "Origen",
    "Dirección Regional": "DireccionRegional",
    "Departamento": "Departamento",
    "Funcionario Seleccionado": "Funcionario",
    "Observaciones": "Observaciones"
};

/* ---------- Estado de la interfaz ---------- */

let idSeleccionado = null;
let idParaEliminar = null;
let registroEnEdicion = null;
let cacheDatos = null;

const Estado_UI = {
    orden: { campo: "Fecha", asc: false },
    pagina: 1,
    porPagina: 10
};

/* ---------- Utilidades ---------- */

const $ = (id) => document.getElementById(id);

function Escapar(texto) {
    const mapa = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
    return String(texto ?? "").replace(/[&<>"']/g, (c) => mapa[c]);
}

function NormalizarTexto(texto) {
    return String(texto ?? "")
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, ""); // quita tildes
}

function GenerarID() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function FechaLocalISO(fecha = new Date()) {
    const p = (n) => String(n).padStart(2, "0");
    return `${fecha.getFullYear()}-${p(fecha.getMonth() + 1)}-${p(fecha.getDate())}`;
}

// Acepta número de serie de Excel, dd/mm/aaaa, dd-mm-aaaa o aaaa-mm-dd. Devuelve aaaa-mm-dd o "".
function NormalizarFechaISO(valor) {
    if (valor === undefined || valor === null || valor === "") return "";

    if (typeof valor === "number") {
        const fecha = new Date(Date.UTC(1899, 11, 30) + Math.floor(valor) * 86400000);
        return fecha.toISOString().slice(0, 10);
    }

    const texto = String(valor).trim();

    let m = texto.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;

    m = texto.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;

    return "";
}

function FormatearFechaVisible(iso) {
    return iso ? iso.split("-").reverse().join("/") : "-";
}

function ObtenerEstado(registro) {
    if (!registro.Fecha) return "sin-fecha";
    const hoy = FechaLocalISO();
    if (hoy < registro.Fecha) return "proxima";
    if (registro.Examen && hoy > registro.Examen) return "finalizada";
    return "en-curso";
}

function ClaveDuplicado(r) {
    return [r.Codigo, r.Curso, r.Funcionario].map(NormalizarTexto).join("|") + "|" + (r.Fecha || "");
}

function BuscarDescripcion(curso) {
    const clave = Object.keys(Descripciones_Cursos)
        .find((k) => NormalizarTexto(k) === NormalizarTexto(curso));
    return clave ? Descripciones_Cursos[clave] : MENSAJE_SIN_DESCRIPCION;
}

function MostrarMensaje(elemento, texto, tipo) {
    if (!elemento) return;
    elemento.textContent = texto;
    elemento.className = `Mensaje_Importacion ${tipo}`;
}

function MostrarToast(mensaje, tipo = "exito") {
    let contenedor = $("ToastContenedor");
    if (!contenedor) {
        contenedor = document.createElement("div");
        contenedor.id = "ToastContenedor";
        contenedor.setAttribute("aria-live", "polite");
        document.body.appendChild(contenedor);
    }
    const toast = document.createElement("div");
    toast.className = `toast ${tipo}`;
    toast.textContent = mensaje;
    contenedor.appendChild(toast);
    setTimeout(() => toast.remove(), 3500);
}

/* ---------- Almacenamiento ---------- */

// Convierte registros antiguos (fechas dd/mm/aaaa, IDs numéricos, typo en "Declación").
function MigrarRegistro(r) {
    return {
        ...r,
        ID: String(r.ID ?? GenerarID()),
        Curso: String(r.Curso ?? "").replace("Declación", "Declaración"),
        Fecha: NormalizarFechaISO(r.Fecha),
        Examen: NormalizarFechaISO(r.Examen)
    };
}

function GuardarDatos(datos) {
    cacheDatos = datos;
    try {
        localStorage.setItem(CLAVE_STORAGE, JSON.stringify(datos));
        return true;
    } catch (error) {
        console.error(error);
        MostrarToast("No se pudo guardar la información en el navegador.", "error");
        return false;
    }
}

function ObtenerDatos() {
    if (cacheDatos) return cacheDatos;

    let datos = null;
    try {
        datos = JSON.parse(localStorage.getItem(CLAVE_STORAGE));
    } catch (error) {
        console.error("Datos almacenados ilegibles:", error);
    }
    if (!Array.isArray(datos)) datos = JSON.parse(JSON.stringify(DatosIniciales));

    GuardarDatos(datos.map(MigrarRegistro));
    return cacheDatos;
}

/* ---------- Tabla: filtros, orden, paginación ---------- */

function ObtenerFiltrados() {
    const busqueda = NormalizarTexto($("IngresarBusqueda").value);
    const depto = $("FiltroDepartamento").value;
    const modalidad = $("FiltroModalidad").value;
    const estado = $("FiltroEstado").value;

    return ObtenerDatos().filter((r) => {
        const texto = NormalizarTexto(`${r.Funcionario} ${r.Curso} ${r.Codigo}`);
        return texto.includes(busqueda)
            && (depto === "" || r.Departamento === depto)
            && (modalidad === "" || r.Modalidad === modalidad)
            && (estado === "" || ObtenerEstado(r) === estado);
    });
}

function ValorOrden(registro, campo) {
    return campo === "Estado" ? ETIQUETAS_ESTADO[ObtenerEstado(registro)] : (registro[campo] ?? "");
}

function OrdenarRegistros(lista) {
    const { campo, asc } = Estado_UI.orden;
    const factor = asc ? 1 : -1;
    return [...lista].sort((a, b) =>
        factor * String(ValorOrden(a, campo)).localeCompare(String(ValorOrden(b, campo)), "es", { numeric: true, sensitivity: "base" })
    );
}

function GenerarEncabezado() {
    $("EncabezadoTabla").innerHTML = COLUMNAS.map((c) =>
        `<th scope="col" data-campo="${c.campo}" tabindex="0" aria-sort="none">${c.titulo}<span class="icono_orden" aria-hidden="true"></span></th>`
    ).join("");
}

function ActualizarIndicadoresOrden() {
    document.querySelectorAll("#EncabezadoTabla th").forEach((th) => {
        const activo = th.dataset.campo === Estado_UI.orden.campo;
        th.setAttribute("aria-sort", activo ? (Estado_UI.orden.asc ? "ascending" : "descending") : "none");
    });
}

function CambiarOrden(campo) {
    if (!campo) return;
    if (Estado_UI.orden.campo === campo) {
        Estado_UI.orden.asc = !Estado_UI.orden.asc;
    } else {
        Estado_UI.orden = { campo, asc: true };
    }
    Estado_UI.pagina = 1;
    Generar_Tabla();
}

function CeldaHTML(registro, columna) {
    if (columna.campo === "Estado") {
        const estado = ObtenerEstado(registro);
        return `<span class="badge badge_${estado}">${ETIQUETAS_ESTADO[estado]}</span>`;
    }
    if (columna.tipo === "fecha") return FormatearFechaVisible(registro[columna.campo]);
    return Escapar(registro[columna.campo]) || "-";
}

function FilaHTML(registro) {
    const celdas = COLUMNAS.map((c) => `<td>${CeldaHTML(registro, c)}</td>`).join("");
    return `<tr class="fila_clickable" data-id="${Escapar(registro.ID)}" tabindex="0" title="Haz clic para ver los detalles">${celdas}</tr>`;
}

function ActualizarResumen(lista) {
    const conteo = { "proxima": 0, "en-curso": 0, "finalizada": 0 };
    lista.forEach((r) => {
        const estado = ObtenerEstado(r);
        if (estado in conteo) conteo[estado]++;
    });
    const item = (clase, numero, texto) =>
        `<div class="resumen_item resumen_${clase}"><span class="resumen_numero">${numero}</span><span class="resumen_texto">${texto}</span></div>`;

    $("Resumen").innerHTML =
        item("total", lista.length, "Registros") +
        item("proxima", conteo["proxima"], "Próximas") +
        item("en-curso", conteo["en-curso"], "En curso") +
        item("finalizada", conteo["finalizada"], "Finalizadas");
}

function ActualizarPaginacion(total, inicio, cantidadVisible, totalPaginas) {
    $("InfoPaginacion").textContent = total === 0
        ? "Sin resultados"
        : `Mostrando ${inicio + 1}-${inicio + cantidadVisible} de ${total}`;
    $("NumeroPagina").textContent = `Página ${Estado_UI.pagina} de ${totalPaginas}`;
    $("BtnPaginaAnterior").disabled = Estado_UI.pagina <= 1;
    $("BtnPaginaSiguiente").disabled = Estado_UI.pagina >= totalPaginas;
}

function Generar_Tabla() {
    const tbody = $("DatosTabla");
    if (!tbody) return;

    const filtrados = OrdenarRegistros(ObtenerFiltrados());
    ActualizarResumen(filtrados);

    const totalPaginas = Math.max(1, Math.ceil(filtrados.length / Estado_UI.porPagina));
    Estado_UI.pagina = Math.min(Estado_UI.pagina, totalPaginas);
    const inicio = (Estado_UI.pagina - 1) * Estado_UI.porPagina;
    const visibles = filtrados.slice(inicio, inicio + Estado_UI.porPagina);

    tbody.innerHTML = filtrados.length === 0
        ? `<tr><td colspan="${COLUMNAS.length}" class="fila_vacia">No se encontraron registros con los filtros actuales.</td></tr>`
        : visibles.map(FilaHTML).join("");

    ActualizarPaginacion(filtrados.length, inicio, visibles.length, totalPaginas);
    ActualizarIndicadoresOrden();
}

function AplicarFiltros() {
    Estado_UI.pagina = 1;
    Generar_Tabla();
}

function LimpiarFiltros() {
    $("IngresarBusqueda").value = "";
    $("FiltroDepartamento").value = "";
    $("FiltroModalidad").value = "";
    $("FiltroEstado").value = "";
    AplicarFiltros();
}

/* ---------- Exportar ---------- */

function ValorExportar(registro, columna) {
    if (columna.campo === "Estado") return ETIQUETAS_ESTADO[ObtenerEstado(registro)];
    if (columna.tipo === "fecha") return registro[columna.campo] ? FormatearFechaVisible(registro[columna.campo]) : "";
    return registro[columna.campo];
}

// CSV con ";" y BOM para que Excel en español lo abra con tildes y columnas correctas.
function ExportarCSV() {
    const filas = OrdenarRegistros(ObtenerFiltrados());
    if (filas.length === 0) {
        MostrarToast("No hay registros para exportar.", "error");
        return;
    }

    const celda = (valor) => {
        let texto = String(valor ?? "");
        if (/^[=+\-@]/.test(texto)) texto = "'" + texto; // evita fórmulas al abrir en Excel
        return `"${texto.replace(/"/g, '""')}"`;
    };

    const lineas = [COLUMNAS.map((c) => celda(c.titulo)).join(";")];
    filas.forEach((r) => lineas.push(COLUMNAS.map((c) => celda(ValorExportar(r, c))).join(";")));

    const blob = new Blob(["\uFEFF" + lineas.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download = `capacitaciones_${FechaLocalISO()}.csv`;
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();
    URL.revokeObjectURL(url);
}

/* ---------- Modales ---------- */

function verDetalles(id) {
    const item = ObtenerDatos().find((r) => r.ID === String(id));
    if (!item) return;

    idSeleccionado = item.ID;
    const estado = ObtenerEstado(item);

    $("Lista_Detalles").innerHTML = `
        <li><strong>Código:</strong> ${Escapar(item.Codigo) || "-"}</li>
        <li><strong>Actividad:</strong> ${Escapar(item.Curso)}</li>
        <li><strong>Funcionario:</strong> ${Escapar(item.Funcionario)}</li>
        <li><strong>Departamento:</strong> ${Escapar(item.Departamento) || "-"}</li>
        <li><strong>Inicio / Término:</strong> ${FormatearFechaVisible(item.Fecha)} a ${FormatearFechaVisible(item.Examen)}</li>
        <li><strong>Estado:</strong> <span class="badge badge_${estado}">${ETIQUETAS_ESTADO[estado]}</span></li>
        <li><strong>Lugar / Modalidad:</strong> ${Escapar(item.Lugar) || "-"} (${Escapar(item.Modalidad) || "-"})</li>
        <li><strong>Observaciones:</strong> ${Escapar(item.Observaciones) || "Sin observaciones"}</li>
    `;
    $("Descripcion_Curso").value = BuscarDescripcion(item.Curso);

    $("Modal_Detalles").style.display = "flex";
    $("BtnAceptarModal").focus();
}

function CerrarModal() {
    $("Modal_Detalles").style.display = "none";
    idSeleccionado = null;
}

function AbrirModalEliminar(id) {
    idParaEliminar = String(id);
    $("InputConfirmarEliminar").value = "";
    $("MensajeErrorConfirmacion").textContent = "";
    $("MensajeErrorConfirmacion").className = "Mensaje_Importacion";
    $("Modal_ConfirmarEliminar").style.display = "flex";
    $("InputConfirmarEliminar").focus();
}

function CerrarModalEliminar() {
    $("Modal_ConfirmarEliminar").style.display = "none";
    idParaEliminar = null;
}

function ConfirmarEliminacionDefinitiva() {
    if (idParaEliminar === null) return;

    if ($("InputConfirmarEliminar").value.trim() !== "Confirmar") {
        MostrarMensaje($("MensajeErrorConfirmacion"), 'Debes escribir exactamente "Confirmar" para continuar.', "error");
        return;
    }

    GuardarDatos(ObtenerDatos().filter((r) => r.ID !== idParaEliminar));
    CerrarModalEliminar();
    Generar_Tabla();
    MostrarToast("Registro eliminado.");
}

/* ---------- Formulario de registro / edición ---------- */

function LlenarListaCursos() {
    const lista = $("ListaCursos");
    if (!lista) return;
    const cursos = new Set([...Object.keys(Descripciones_Cursos), ...ObtenerDatos().map((r) => r.Curso)]);
    lista.innerHTML = [...cursos].filter(Boolean).map((c) => `<option value="${Escapar(c)}"></option>`).join("");
}

function LeerFormulario() {
    const datos = {};
    Object.entries(CamposFormulario).forEach(([campo, idInput]) => {
        datos[campo] = $(idInput).value.trim();
    });
    return datos;
}

function GuardarRegistro(event) {
    event.preventDefault();
    const mensaje = $("MensajeFormulario");
    const datos = LeerFormulario();

    if (datos.Fecha && datos.Examen && datos.Examen < datos.Fecha) {
        MostrarMensaje(mensaje, "La fecha de término no puede ser anterior a la fecha de inicio.", "error");
        $("Nuevo_Examen").focus();
        return;
    }

    const registros = ObtenerDatos();
    const clave = ClaveDuplicado(datos);
    if (registros.some((r) => r.ID !== registroEnEdicion && ClaveDuplicado(r) === clave)) {
        MostrarMensaje(mensaje, "Este funcionario ya está registrado en esa actividad con la misma fecha de inicio.", "error");
        return;
    }

    const resultado = registroEnEdicion
        ? registros.map((r) => (r.ID === registroEnEdicion ? { ...r, ...datos } : r))
        : [...registros, { ID: GenerarID(), ...datos }];

    if (!GuardarDatos(resultado)) return;

    sessionStorage.setItem("ToastPendiente", registroEnEdicion ? "Registro actualizado correctamente." : "Registro guardado correctamente.");
    window.location.href = "index.html";
}

function InicializarFormulario() {
    const form = $("FormCapacitacion");
    if (!form) return;

    LlenarListaCursos();

    const idEditar = new URLSearchParams(window.location.search).get("editar");
    if (idEditar) {
        const registro = ObtenerDatos().find((r) => r.ID === idEditar);
        if (registro) {
            registroEnEdicion = registro.ID;
            Object.entries(CamposFormulario).forEach(([campo, idInput]) => {
                $(idInput).value = registro[campo] ?? "";
            });
            $("TituloFormulario").textContent = "Editar Registro";
            $("BtnGuardarRegistro").textContent = "Guardar cambios";
        } else {
            MostrarMensaje($("MensajeFormulario"), "No se encontró el registro que intentas editar.", "error");
        }
    }

    form.addEventListener("submit", GuardarRegistro);
}

/* ---------- Importación desde Excel ---------- */

function ProcesarArchivoExcel(file) {
    const mensaje = $("MensajeImportacion");
    const reader = new FileReader();

    reader.onerror = () => MostrarMensaje(mensaje, "No se pudo leer el archivo seleccionado.", "error");

    reader.onload = (e) => {
        try {
            const libro = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
            const hoja = libro.Sheets[libro.SheetNames[0]];
            const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" });

            if (!filas || filas.length < 2) {
                MostrarMensaje(mensaje, "El archivo no contiene datos para importar.", "error");
                return;
            }

            const encabezados = filas[0].map(NormalizarTexto);
            const indice = {};
            Object.entries(MapeoColumnasExcel).forEach(([nombreEsperado, campo]) => {
                const posicion = encabezados.indexOf(NormalizarTexto(nombreEsperado));
                if (posicion !== -1) indice[campo] = posicion;
            });

            if (indice.Funcionario === undefined || indice.Curso === undefined) {
                MostrarMensaje(mensaje, "Faltan las columnas mínimas: Actividad y Funcionario Seleccionado. Revisa los encabezados o descarga la plantilla.", "error");
                return;
            }

            const actuales = ObtenerDatos();
            const clavesExistentes = new Set(actuales.map(ClaveDuplicado));
            const nuevos = [];
            let duplicados = 0, incompletas = 0, fechasNoReconocidas = 0;

            for (let i = 1; i < filas.length; i++) {
                const fila = filas[i];
                if (fila.every((celda) => String(celda).trim() === "")) continue;

                const texto = (campo) => String(indice[campo] !== undefined ? fila[indice[campo]] : "").trim();
                const fecha = (campo) => {
                    const bruto = indice[campo] !== undefined ? fila[indice[campo]] : "";
                    const iso = NormalizarFechaISO(bruto);
                    if (bruto !== "" && !iso) fechasNoReconocidas++;
                    return iso;
                };

                const registro = {
                    ID: GenerarID(),
                    Codigo: texto("Codigo"),
                    Curso: texto("Curso"),
                    Fecha: fecha("Fecha"),
                    Examen: fecha("Examen"),
                    Modalidad: texto("Modalidad"),
                    Lugar: texto("Lugar"),
                    TipoInscripcion: texto("TipoInscripcion"),
                    Origen: texto("Origen"),
                    DireccionRegional: texto("DireccionRegional"),
                    Departamento: texto("Departamento"),
                    Funcionario: texto("Funcionario"),
                    Observaciones: texto("Observaciones")
                };

                if (!registro.Funcionario || !registro.Curso) { incompletas++; continue; }

                const clave = ClaveDuplicado(registro);
                if (clavesExistentes.has(clave)) { duplicados++; continue; }

                clavesExistentes.add(clave);
                nuevos.push(registro);
            }

            if (nuevos.length === 0) {
                MostrarMensaje(mensaje, "No se importó ningún registro: las filas estaban incompletas o ya existían.", "error");
                return;
            }

            if (!GuardarDatos(actuales.concat(nuevos))) return;

            let resumen = `Se importaron ${nuevos.length} registro(s) correctamente.`;
            if (duplicados) resumen += ` ${duplicados} omitido(s) por estar duplicados.`;
            if (incompletas) resumen += ` ${incompletas} fila(s) sin actividad o funcionario.`;
            if (fechasNoReconocidas) resumen += ` ${fechasNoReconocidas} fecha(s) no reconocida(s) quedaron vacías.`;
            MostrarMensaje(mensaje, resumen, "exito");

            $("ArchivoExcel").value = "";
            LlenarListaCursos();
            Generar_Tabla();
        } catch (error) {
            console.error(error);
            MostrarMensaje(mensaje, "Ocurrió un error al procesar el archivo. Verifica el formato.", "error");
        }
    };

    reader.readAsArrayBuffer(file);
}

function DescargarPlantilla() {
    if (typeof XLSX === "undefined") {
        MostrarMensaje($("MensajeImportacion"), "No se pudo cargar la librería de Excel. Revisa tu conexión.", "error");
        return;
    }
    const encabezados = Object.keys(MapeoColumnasExcel);
    const ejemplo = [
        "ACT-2026-02", "Atención al Contribuyente", "10/11/2026", "12/11/2026", "Presencial",
        "Punta Arenas", "Directa", "Nacional", "XII DIRECCIÓN REGIONAL PUNTA ARENAS",
        "Administración", "Nombre Apellido", "Sin observaciones"
    ];
    const hoja = XLSX.utils.aoa_to_sheet([encabezados, ejemplo]);
    hoja["!cols"] = encabezados.map(() => ({ wch: 24 }));
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Funcionarios");
    XLSX.writeFile(libro, "Plantilla_Importacion.xlsx");
}

function InicializarImportacion() {
    const btnImportar = $("BtnImportarExcel");
    if (!btnImportar) return;

    btnImportar.addEventListener("click", () => {
        const mensaje = $("MensajeImportacion");
        const archivos = $("ArchivoExcel").files;

        if (!archivos || archivos.length === 0) {
            MostrarMensaje(mensaje, "Debes seleccionar un archivo primero.", "error");
            return;
        }
        if (typeof XLSX === "undefined") {
            MostrarMensaje(mensaje, "No se pudo cargar la librería de Excel. Revisa tu conexión.", "error");
            return;
        }
        ProcesarArchivoExcel(archivos[0]);
    });

    $("BtnDescargarPlantilla").addEventListener("click", DescargarPlantilla);
}

/* ---------- Inicialización ---------- */

function InicializarTabla() {
    if (!$("DatosTabla")) return;

    GenerarEncabezado();

    const toastPendiente = sessionStorage.getItem("ToastPendiente");
    if (toastPendiente) {
        sessionStorage.removeItem("ToastPendiente");
        MostrarToast(toastPendiente);
    }

    // Filtros
    $("IngresarBusqueda").addEventListener("input", AplicarFiltros);
    ["FiltroDepartamento", "FiltroModalidad", "FiltroEstado"].forEach((id) => $(id).addEventListener("change", AplicarFiltros));
    $("BtnLimpiarFiltros").addEventListener("click", LimpiarFiltros);
    $("BtnExportarCSV").addEventListener("click", ExportarCSV);

    // Orden por columna (clic o teclado)
    const encabezado = $("EncabezadoTabla");
    encabezado.addEventListener("click", (e) => {
        const th = e.target.closest("th");
        if (th) CambiarOrden(th.dataset.campo);
    });
    encabezado.addEventListener("keydown", (e) => {
        const th = e.target.closest("th");
        if (th && (e.key === "Enter" || e.key === " ")) {
            e.preventDefault();
            CambiarOrden(th.dataset.campo);
        }
    });

    // Filas: una sola escucha para toda la tabla
    const tbody = $("DatosTabla");
    tbody.addEventListener("click", (e) => {
        const fila = e.target.closest("tr[data-id]");
        if (fila) verDetalles(fila.dataset.id);
    });
    tbody.addEventListener("keydown", (e) => {
        const fila = e.target.closest("tr[data-id]");
        if (fila && e.key === "Enter") verDetalles(fila.dataset.id);
    });

    // Paginación
    $("BtnPaginaAnterior").addEventListener("click", () => { Estado_UI.pagina--; Generar_Tabla(); });
    $("BtnPaginaSiguiente").addEventListener("click", () => { Estado_UI.pagina++; Generar_Tabla(); });

    // Modales
    $("BtnAceptarModal").addEventListener("click", CerrarModal);
    $("BtnEditarModal").addEventListener("click", () => {
        if (idSeleccionado !== null) {
            window.location.href = `Registro_Funcionarios.html?editar=${encodeURIComponent(idSeleccionado)}`;
        }
    });
    $("BtnEliminarModal").addEventListener("click", () => {
        const idTemp = idSeleccionado;
        CerrarModal();
        if (idTemp !== null) AbrirModalEliminar(idTemp);
    });
    $("BtnCancelarEliminar").addEventListener("click", CerrarModalEliminar);
    $("BtnConfirmarEliminarDefinitivo").addEventListener("click", ConfirmarEliminacionDefinitiva);
    $("InputConfirmarEliminar").addEventListener("keydown", (e) => {
        if (e.key === "Enter") ConfirmarEliminacionDefinitiva();
    });

    // Cerrar con clic en el fondo o con Escape
    const cierres = { Modal_Detalles: CerrarModal, Modal_ConfirmarEliminar: CerrarModalEliminar };
    Object.entries(cierres).forEach(([id, cerrar]) => {
        $(id).addEventListener("click", (e) => { if (e.target === e.currentTarget) cerrar(); });
    });
    document.addEventListener("keydown", (e) => {
        if (e.key !== "Escape") return;
        if ($("Modal_ConfirmarEliminar").style.display === "flex") CerrarModalEliminar();
        else if ($("Modal_Detalles").style.display === "flex") CerrarModal();
    });

    Generar_Tabla();
}

document.addEventListener("DOMContentLoaded", () => {
    InicializarTabla();
    InicializarFormulario();
    InicializarImportacion();
});

// Si los datos cambian en otra pestaña, se refresca la tabla.
window.addEventListener("storage", (e) => {
    if (e.key === CLAVE_STORAGE) {
        cacheDatos = null;
        Generar_Tabla();
    }
});
