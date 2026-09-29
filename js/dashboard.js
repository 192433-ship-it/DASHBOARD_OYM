// ==========================================================
// DASHBOARD OYM - DASHBOARD.JS
// ==========================================================

// ----------------------------------------------------------
// CONFIGURACIÓN
// ----------------------------------------------------------

const URL_DATOS =
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQEH6rFknTB0x7RFIHLKlujRxRCCBLsJE7a6wZ06zzZaRXOY2sLnp4dmv2kkNABIEt7HzbGfhIyPHf1/pub?output=csv";


// ----------------------------------------------------------
// VARIABLES
// ----------------------------------------------------------

let datosOriginales = [];
let datosVisibles = [];

let graficoDia = null;
let graficoEstado = null;
let graficoResponsable = null;
let graficoPrioridad = null;

let ordenColumna = "";
let ordenAscendente = true;


// ----------------------------------------------------------
// INICIO
// ----------------------------------------------------------

document.addEventListener("DOMContentLoaded", function () {

    console.log("Dashboard OYM iniciado");

    configurarEventos();

    cargarDatos();

});


// ----------------------------------------------------------
// CONFIGURAR EVENTOS
// ----------------------------------------------------------

function configurarEventos() {

    const buscador = document.getElementById("buscador");
    const filtroRegion = document.getElementById("filtroRegion");
    const filtroEstado = document.getElementById("filtroEstado");
    const filtroResponsable = document.getElementById("filtroResponsable");
    const filtroPrioridad = document.getElementById("filtroPrioridad");
    const filtroTipo = document.getElementById("filtroTipo");
    const filtroArea = document.getElementById("filtroArea");

    const btnLimpiar = document.getElementById("btnLimpiarFiltros");
    const btnActualizar = document.getElementById("btnActualizar");


    if (buscador) {
        buscador.addEventListener("input", aplicarFiltros);
    }

    if (filtroRegion) {
        filtroRegion.addEventListener("change", aplicarFiltros);
    }

    if (filtroEstado) {
        filtroEstado.addEventListener("change", aplicarFiltros);
    }

    if (filtroResponsable) {
        filtroResponsable.addEventListener("change", aplicarFiltros);
    }

    if (filtroPrioridad) {
        filtroPrioridad.addEventListener("change", aplicarFiltros);
    }

    if (filtroTipo) {
        filtroTipo.addEventListener("change", aplicarFiltros);
    }

    if (filtroArea) {
        filtroArea.addEventListener("change", aplicarFiltros);
    }


    if (btnLimpiar) {
        btnLimpiar.addEventListener("click", limpiarFiltros);
    }

    if (btnActualizar) {
        btnActualizar.addEventListener("click", cargarDatos);
    }


    // Ordenamiento de columnas

    const encabezados =
        document.querySelectorAll("th[data-columna]");


    encabezados.forEach(function (th) {

        th.addEventListener("click", function () {

            const columna =
                th.getAttribute("data-columna");

            ordenarDatos(columna);

        });

    });

}


// ----------------------------------------------------------
// CARGAR DATOS
// ----------------------------------------------------------

async function cargarDatos() {

    const boton =
        document.getElementById("btnActualizar");


    try {

        if (boton) {

            boton.disabled = true;

            boton.textContent =
                "⟳ Actualizando...";

        }


        console.log("Cargando datos...");


        const respuesta =
            await fetch(
                URL_DATOS + "&t=" + Date.now()
            );


        if (!respuesta.ok) {

            throw new Error(
                "No se pudo obtener el archivo de Google Sheets"
            );

        }


        const textoCSV =
            await respuesta.text();


        console.log("CSV recibido");


        if (!textoCSV.trim()) {

            throw new Error(
                "Google Sheets no devolvió datos"
            );

        }


        if (typeof XLSX === "undefined") {

            throw new Error(
                "La librería XLSX no está cargada"
            );

        }


        const workbook =
            XLSX.read(textoCSV, {
                type: "string"
            });


        const nombreHoja =
            workbook.SheetNames[0];


        const hoja =
            workbook.Sheets[nombreHoja];


        const datos =
            XLSX.utils.sheet_to_json(
                hoja,
                {
                    defval: ""
                }
            );


        console.log(
            "Filas recibidas:",
            datos.length
        );


        datosOriginales =
            normalizarDatos(datos);


        datosVisibles =
            [...datosOriginales];


        cargarFiltros();


        actualizarTodo();


        if (boton) {

            boton.textContent =
                "✓ Datos actualizados";


            setTimeout(function () {

                boton.textContent =
                    "↻ Actualizar datos";

                boton.disabled = false;

            }, 1200);

        }


    } catch (error) {

        console.error(
            "ERROR AL CARGAR DATOS:",
            error
        );


        if (boton) {

            boton.textContent =
                "⚠ Error al actualizar";

            boton.disabled = false;

        }


        mostrarError(
            "No se pudieron cargar los datos"
        );

    }

}


// ----------------------------------------------------------
// NORMALIZAR DATOS
// ----------------------------------------------------------

function normalizarDatos(datos) {

    return datos.map(function (fila) {

        return {

            Codigo:
                obtenerValor(
                    fila,
                    ["Codigo", "Código"]
                ),

            Site:
                obtenerValor(
                    fila,
                    ["Site", "SITE"]
                ),

            Region:
                obtenerValor(
                    fila,
                    [
                        "Región",
                        "RegiÃ³n",
                        "Region"
                    ]
                ),

            Actividad:
                obtenerValor(
                    fila,
                    ["Actividad"]
                ),

            Responsable:
                obtenerValor(
                    fila,
                    ["Responsable"]
                ),

            Estado:
                obtenerValor(
                    fila,
                    ["Estado"]
                ),

            Prioridad:
                obtenerValor(
                    fila,
                    ["Prioridad"]
                ),

            Tipo:
                obtenerValor(
                    fila,
                    ["Tipo"]
                ),

            Area:
                obtenerValor(
                    fila,
                    [
                        "Area",
                        "Área"
                    ]
                ),

            FechaInicio:
                obtenerValor(
                    fila,
                    ["Fecha Inicio"]
                ),

            FechaCierre:
                obtenerValor(
                    fila,
                    ["Fecha Cierre"]
                )

        };

    });

}


// ----------------------------------------------------------
// OBTENER VALOR
// ----------------------------------------------------------

function obtenerValor(objeto, nombres) {

    for (
        let i = 0;
        i < nombres.length;
        i++
    ) {

        const nombre =
            nombres[i];


        if (
            Object.prototype.hasOwnProperty.call(
                objeto,
                nombre
            ) &&
            objeto[nombre] !== undefined &&
            objeto[nombre] !== null
        ) {

            return String(
                objeto[nombre]
            ).trim();

        }

    }


    return "";

}


// ----------------------------------------------------------
// ACTUALIZAR TODO
// ----------------------------------------------------------

function actualizarTodo() {

    actualizarKPIs();

    actualizarTabla();

    actualizarGraficoDia();

    actualizarGraficoEstado();

    actualizarGraficoResponsable();

    actualizarGraficoPrioridad();

}


// ----------------------------------------------------------
// KPI
// ----------------------------------------------------------

function actualizarKPIs() {

    const total =
        datosVisibles.length;


    let completadas = 0;

    let pendientes = 0;

    let enProceso = 0;


    datosVisibles.forEach(
        function (fila) {

            const estado =
                normalizarTexto(
                    fila.Estado
                );


            if (
                estado === "completado" ||
                estado === "completada"
            ) {

                completadas++;

            }

            else if (
                estado === "pendiente"
            ) {

                pendientes++;

            }

            else if (
                estado === "en proceso" ||
                estado === "enproceso"
            ) {

                enProceso++;

            }

        }
    );


    const porcentaje =
        total > 0
            ? Math.round(
                (completadas / total) * 100
            )
            : 0;


    ponerTexto(
        "totalActividades",
        total
    );


    ponerTexto(
        "completadas",
        completadas
    );


    ponerTexto(
        "pendientes",
        pendientes
    );


    ponerTexto(
        "enProceso",
        enProceso
    );


    ponerTexto(
        "cumplimiento",
        porcentaje + "%"
    );

}


// ----------------------------------------------------------
// PONER TEXTO
// ----------------------------------------------------------

function ponerTexto(id, texto) {

    const elemento =
        document.getElementById(id);


    if (elemento) {

        elemento.textContent =
            texto;

    }

}


// ----------------------------------------------------------
// GRÁFICO POR DÍA
// ----------------------------------------------------------

// ----------------------------------------------------------
// GRÁFICO ACTIVIDADES EN EJECUCIÓN POR MES
// ----------------------------------------------------------

function actualizarGraficoDia() {

    const canvas =
        document.getElementById(
            "miGrafico"
        );


    if (!canvas) {

        console.warn(
            "No existe el canvas miGrafico"
        );

        return;

    }


    const actividadesPorMes = {};


    datosVisibles.forEach(
        function (fila) {

            const fechaInicio =
                convertirFecha(
                    fila.FechaInicio
                );


            if (!fechaInicio) {

                return;

            }


            const inicio =
                crearFechaLocal(
                    fechaInicio
                );


            if (!inicio) {

                return;

            }


            let fin;


            if (fila.FechaCierre) {

                const fechaCierre =
                    convertirFecha(
                        fila.FechaCierre
                    );


                fin =
                    crearFechaLocal(
                        fechaCierre
                    );

            }


            // Si no tiene fecha de cierre,
            // se considera activa hasta el mes actual

            if (!fin) {

                const ahora =
                    new Date();

                fin =
                    new Date(
                        ahora.getFullYear(),
                        ahora.getMonth(),
                        1
                    );

            }


            // Primer día del mes de inicio

            let mesActual =
                new Date(
                    inicio.getFullYear(),
                    inicio.getMonth(),
                    1
                );


            // Primer día del mes de cierre

            const mesFin =
                new Date(
                    fin.getFullYear(),
                    fin.getMonth(),
                    1
                );


            // Recorrer todos los meses
            // en los que la actividad estuvo activa

            while (
                mesActual <= mesFin
            ) {

                const clave =
                    mesActual.getFullYear() +
                    "-" +
                    String(
                        mesActual.getMonth() + 1
                    ).padStart(
                        2,
                        "0"
                    );


                if (
                    !actividadesPorMes[clave]
                ) {

                    actividadesPorMes[clave] =
                        0;

                }


                actividadesPorMes[clave]++;


                mesActual.setMonth(
                    mesActual.getMonth() + 1
                );

            }

        }
    );


    const meses =
        Object.keys(
            actividadesPorMes
        ).sort();


    const etiquetas =
        meses.map(
            function (mes) {

                const partes =
                    mes.split("-");


                const fecha =
                    new Date(
                        Number(partes[0]),
                        Number(partes[1]) - 1,
                        1
                    );


                return fecha.toLocaleDateString(
                    "es-PE",
                    {
                        month: "short",
                        year: "numeric"
                    }
                );

            }
        );


    const valores =
        meses.map(
            function (mes) {

                return actividadesPorMes[mes];

            }
        );


    if (graficoDia) {

        graficoDia.destroy();

    }


    graficoDia =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        etiquetas,

                    datasets: [

                        {

                            label:
                                "Actividades en ejecución",

                            data:
                                valores,

                            borderWidth:
                                3,

                            tension:
                                0.3,

                            fill:
                                false,

                            pointRadius:
                                5

                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    plugins: {

                        legend: {

                            display:
                                true

                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    function (
                                        contexto
                                    ) {

                                        return (
                                            "Actividades activas: " +
                                            contexto.raw
                                        );

                                    }

                            }

                        }

                    },

                    scales: {

                        y: {

                            beginAtZero:
                                true,

                            ticks: {

                                precision:
                                    0

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Actividades"

                            }

                        },

                        x: {

                            title: {

                                display:
                                    true,

                                text:
                                    "Mes"

                            }

                        }

                    }

                }

            }
        );

}


// ----------------------------------------------------------
// CREAR FECHA LOCAL
// ----------------------------------------------------------

function crearFechaLocal(fecha) {

    if (!fecha) {

        return null;

    }


    const partes =
        fecha.split("-");


    if (
        partes.length !== 3
    ) {

        return null;

    }


    const anio =
        Number(partes[0]);

    const mes =
        Number(partes[1]) - 1;

    const dia =
        Number(partes[2]);


    const resultado =
        new Date(
            anio,
            mes,
            dia
        );


    if (
        isNaN(
            resultado.getTime()
        )
    ) {

        return null;

    }


    return resultado;

}

// ----------------------------------------------------------
// GRÁFICO ESTADO
// ----------------------------------------------------------

function actualizarGraficoEstado() {

    const canvas =
        document.getElementById(
            "graficoEstado"
        );


    if (!canvas) {

        console.warn(
            "No existe el canvas graficoEstado"
        );

        return;

    }


    let completadas = 0;

    let pendientes = 0;

    let enProceso = 0;


    datosVisibles.forEach(
        function (fila) {

            const estado =
                normalizarTexto(
                    fila.Estado
                );


            if (
                estado === "completado" ||
                estado === "completada"
            ) {

                completadas++;

            }

            else if (
                estado === "pendiente"
            ) {

                pendientes++;

            }

            else if (
                estado === "en proceso" ||
                estado === "enproceso"
            ) {

                enProceso++;

            }

        }
    );


    if (graficoEstado) {

        graficoEstado.destroy();

    }


    graficoEstado =
        new Chart(
            canvas,
            {

                type: "doughnut",

                data: {

                    labels: [

                        "Completadas",

                        "Pendientes",

                        "En proceso"

                    ],

                    datasets: [

                        {

                            data: [

                                completadas,

                                pendientes,

                                enProceso

                            ],

                            borderWidth:
                                2

                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    plugins: {

                        legend: {

                            position:
                                "bottom"

                        }

                    }

                }

            }
        );

}


// ----------------------------------------------------------
// GRÁFICO RESPONSABLE
// ----------------------------------------------------------

function actualizarGraficoResponsable() {

    const canvas =
        document.getElementById(
            "graficoResponsable"
        );


    if (!canvas) {

        console.warn(
            "No existe el canvas graficoResponsable"
        );

        return;

    }


    const cantidades = {};


    datosVisibles.forEach(
        function (fila) {

            const responsable =
                fila.Responsable &&
                fila.Responsable.trim()
                    ? fila.Responsable.trim()
                    : "Sin asignar";


            if (!cantidades[responsable]) {

                cantidades[responsable] =
                    0;

            }


            cantidades[responsable]++;

        }
    );


    const responsables =
        Object.keys(
            cantidades
        );


    const valores =
        responsables.map(
            function (responsable) {

                return cantidades[
                    responsable
                ];

            }
        );


    if (graficoResponsable) {

        graficoResponsable.destroy();

    }


    graficoResponsable =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels:
                        responsables,

                    datasets: [

                        {

                            label:
                                "Actividades",

                            data:
                                valores,

                            borderWidth:
                                1

                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },

                    scales: {

                        y: {

                            beginAtZero:
                                true,

                            ticks: {

                                precision:
                                    0

                            }

                        }

                    }

                }

            }
        );

}


// ----------------------------------------------------------
// GRÁFICO PRIORIDAD - DOUGHNUT
// ----------------------------------------------------------

function actualizarGraficoPrioridad() {

    const canvas =
        document.getElementById(
            "graficoPrioridad"
        );


    if (!canvas) {

        console.warn(
            "No existe el canvas graficoPrioridad"
        );

        return;

    }


    let alta = 0;

    let media = 0;

    let baja = 0;

    let sinPrioridad = 0;


    datosVisibles.forEach(
        function (fila) {

            const prioridad =
                normalizarTexto(
                    fila.Prioridad
                );


            if (prioridad === "alta") {

                alta++;

            }

            else if (prioridad === "media") {

                media++;

            }

            else if (prioridad === "baja") {

                baja++;

            }

            else {

                sinPrioridad++;

            }

        }
    );


    if (graficoPrioridad) {

        graficoPrioridad.destroy();

    }


    graficoPrioridad =
        new Chart(
            canvas,
            {

                type: "doughnut",

                data: {

                    labels: [

                        "Alta",

                        "Media",

                        "Baja",

                        "Sin prioridad"

                    ],

                    datasets: [

                        {

                            data: [

                                alta,

                                media,

                                baja,

                                sinPrioridad

                            ],

                            borderWidth: 2

                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    cutout: "60%",

                    plugins: {

                        legend: {

                            position:
                                "bottom",

                            labels: {

                                padding: 15

                            }

                        }

                    }

                }

            }
        );

}


// ----------------------------------------------------------
// TABLA
// ----------------------------------------------------------

function actualizarTabla() {

    const tbody =
        document.getElementById(
            "tablaDatos"
        );


    if (!tbody) {

        console.warn(
            "No existe tablaDatos"
        );

        return;

    }


    tbody.innerHTML = "";


    datosVisibles.forEach(
        function (fila, indice) {

            const tr =
                document.createElement(
                    "tr"
                );


            tr.innerHTML = `

                <td>${indice + 1}</td>

                <td>
                    ${escaparHTML(
                        fila.Codigo
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        fila.Site
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        fila.Region
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        fila.Actividad
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        fila.Responsable
                    )}
                </td>

                <td>
                    ${crearEstadoHTML(
                        fila.Estado
                    )}
                </td>

                <td>
                    ${crearPrioridadHTML(
                        fila.Prioridad
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        fila.Tipo
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        fila.Area
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        formatearFecha(
                            fila.FechaInicio
                        )
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        formatearFecha(
                            fila.FechaCierre
                        )
                    )}
                </td>

            `;


            tbody.appendChild(
                tr
            );

        }
    );


    const contador =
        document.getElementById(
            "contadorResultados"
        );


    if (contador) {

        contador.textContent =
            "Mostrando " +
            datosVisibles.length +
            " actividades";

    }

}


// ----------------------------------------------------------
// ESTADO HTML
// ----------------------------------------------------------

function crearEstadoHTML(estado) {

    const texto =
        estado && estado.trim()
            ? estado.trim()
            : "Sin estado";


    const normalizado =
        normalizarTexto(
            texto
        );


    let clase =
        "estado";


    if (
        normalizado === "completado" ||
        normalizado === "completada"
    ) {

        clase +=
            " estado-completado";

    }

    else if (
        normalizado === "pendiente"
    ) {

        clase +=
            " estado-pendiente";

    }

    else if (
        normalizado === "en proceso" ||
        normalizado === "enproceso"
    ) {

        clase +=
            " estado-en-proceso";

    }


    return `

        <span class="${clase}">

            ${escaparHTML(texto)}

        </span>

    `;

}


// ----------------------------------------------------------
// PRIORIDAD HTML
// ----------------------------------------------------------

function crearPrioridadHTML(prioridad) {

    const texto =
        prioridad && prioridad.trim()
            ? prioridad.trim()
            : "Sin prioridad";


    const normalizado =
        normalizarTexto(
            texto
        );


    let clase =
        "prioridad";


    if (
        normalizado === "alta"
    ) {

        clase +=
            " prioridad-alta";

    }

    else if (
        normalizado === "media"
    ) {

        clase +=
            " prioridad-media";

    }

    else if (
        normalizado === "baja"
    ) {

        clase +=
            " prioridad-baja";

    }


    return `

        <span class="${clase}">

            ${escaparHTML(texto)}

        </span>

    `;

}


// ----------------------------------------------------------
// FILTROS
// ----------------------------------------------------------

function aplicarFiltros() {

    const buscador =
        obtenerValorElemento(
            "buscador"
        );


    const region =
        obtenerValorElemento(
            "filtroRegion"
        );


    const estado =
        obtenerValorElemento(
            "filtroEstado"
        );


    const responsable =
        obtenerValorElemento(
            "filtroResponsable"
        );


    const prioridad =
        obtenerValorElemento(
            "filtroPrioridad"
        );


    const tipo =
        obtenerValorElemento(
            "filtroTipo"
        );


    const area =
        obtenerValorElemento(
            "filtroArea"
        );


    const textoBusqueda =
        normalizarTexto(
            buscador
        );


    datosVisibles =
        datosOriginales.filter(
            function (fila) {


                const textoCompleto =
                    normalizarTexto(

                        [

                            fila.Codigo,

                            fila.Site,

                            fila.Region,

                            fila.Actividad,

                            fila.Responsable,

                            fila.Estado,

                            fila.Prioridad,

                            fila.Tipo,

                            fila.Area

                        ].join(" ")

                    );


                const cumpleBusqueda =
                    !textoBusqueda ||
                    textoCompleto.includes(
                        textoBusqueda
                    );


                const cumpleRegion =
                    !region ||
                    fila.Region === region;


                const cumpleEstado =
                    !estado ||
                    fila.Estado === estado;


                const cumpleResponsable =
                    !responsable ||
                    fila.Responsable === responsable;


                const cumplePrioridad =
                    !prioridad ||
                    fila.Prioridad === prioridad;


                const cumpleTipo =
                    !tipo ||
                    fila.Tipo === tipo;


                const cumpleArea =
                    !area ||
                    fila.Area === area;


                return (

                    cumpleBusqueda &&

                    cumpleRegion &&

                    cumpleEstado &&

                    cumpleResponsable &&

                    cumplePrioridad &&

                    cumpleTipo &&

                    cumpleArea

                );

            }
        );


    actualizarTodo();

}


// ----------------------------------------------------------
// OBTENER VALOR ELEMENTO
// ----------------------------------------------------------

function obtenerValorElemento(id) {

    const elemento =
        document.getElementById(
            id
        );


    if (!elemento) {

        return "";

    }


    return elemento.value.trim();

}


// ----------------------------------------------------------
// CARGAR FILTROS
// ----------------------------------------------------------

function cargarFiltros() {

    llenarSelect(

        "filtroRegion",

        datosOriginales.map(
            function (fila) {

                return fila.Region;

            }
        ),

        "Todas las regiones"

    );


    llenarSelect(

        "filtroEstado",

        datosOriginales.map(
            function (fila) {

                return fila.Estado;

            }
        ),

        "Todos los estados"

    );


    llenarSelect(

        "filtroResponsable",

        datosOriginales.map(
            function (fila) {

                return fila.Responsable;

            }
        ),

        "Todos los responsables"

    );


    llenarSelect(

        "filtroPrioridad",

        datosOriginales.map(
            function (fila) {

                return fila.Prioridad;

            }
        ),

        "Todas las prioridades"

    );


    llenarSelect(

        "filtroTipo",

        datosOriginales.map(
            function (fila) {

                return fila.Tipo;

            }
        ),

        "Todos los tipos"

    );


    llenarSelect(

        "filtroArea",

        datosOriginales.map(
            function (fila) {

                return fila.Area;

            }
        ),

        "Todas las áreas"

    );

}


// ----------------------------------------------------------
// LLENAR SELECT
// ----------------------------------------------------------

function llenarSelect(
    id,
    valores,
    textoInicial
) {

    const select =
        document.getElementById(
            id
        );


    if (!select) {

        return;

    }


    const valorActual =
        select.value;


    const valoresUnicos = [

        ...new Set(

            valores

                .filter(
                    function (valor) {

                        return (
                            valor &&
                            valor.trim()
                        );

                    }
                )

                .map(
                    function (valor) {

                        return valor.trim();

                    }
                )

        )

    ].sort(
        function (a, b) {

            return a.localeCompare(
                b,
                "es",
                {
                    sensitivity:
                        "base"
                }
            );

        }
    );


    select.innerHTML =
        "";


    const opcionInicial =
        document.createElement(
            "option"
        );


    opcionInicial.value =
        "";


    opcionInicial.textContent =
        textoInicial;


    select.appendChild(
        opcionInicial
    );


    valoresUnicos.forEach(
        function (valor) {

            const opcion =
                document.createElement(
                    "option"
                );


            opcion.value =
                valor;


            opcion.textContent =
                valor;


            select.appendChild(
                opcion
            );

        }
    );


    if (
        valoresUnicos.includes(
            valorActual
        )
    ) {

        select.value =
            valorActual;

    }

}


// ----------------------------------------------------------
// LIMPIAR FILTROS
// ----------------------------------------------------------

function limpiarFiltros() {

    const ids = [

        "buscador",

        "filtroRegion",

        "filtroEstado",

        "filtroResponsable",

        "filtroPrioridad",

        "filtroTipo",

        "filtroArea"

    ];


    ids.forEach(
        function (id) {

            const elemento =
                document.getElementById(
                    id
                );


            if (!elemento) {

                return;

            }


            elemento.value =
                "";

        }
    );


    datosVisibles =
        [...datosOriginales];


    actualizarTodo();

}


// ----------------------------------------------------------
// ORDENAR TABLA
// ----------------------------------------------------------

function ordenarDatos(columna) {

    if (
        ordenColumna === columna
    ) {

        ordenAscendente =
            !ordenAscendente;

    }

    else {

        ordenColumna =
            columna;

        ordenAscendente =
            true;

    }


    const mapaColumnas = {

        "Codigo":
            "Codigo",

        "Site":
            "Site",

        "RegiÃ³n":
            "Region",

        "Región":
            "Region",

        "Actividad":
            "Actividad",

        "Responsable":
            "Responsable",

        "Estado":
            "Estado",

        "Prioridad":
            "Prioridad",

        "Tipo":
            "Tipo",

        "Area":
            "Area",

        "Área":
            "Area",

        "Fecha Inicio":
            "FechaInicio",

        "Fecha Cierre":
            "FechaCierre"

    };


    const propiedad =
        mapaColumnas[
            columna
        ];


    if (!propiedad) {

        return;

    }


    datosVisibles.sort(
        function (a, b) {

            let valorA =
                a[propiedad] || "";


            let valorB =
                b[propiedad] || "";


            valorA =
                String(
                    valorA
                ).toLowerCase();


            valorB =
                String(
                    valorB
                ).toLowerCase();


            if (
                valorA < valorB
            ) {

                return ordenAscendente
                    ? -1
                    : 1;

            }


            if (
                valorA > valorB
            ) {

                return ordenAscendente
                    ? 1
                    : -1;

            }


            return 0;

        }
    );


    actualizarTabla();

}


// ----------------------------------------------------------
// FECHAS
// ----------------------------------------------------------

function convertirFecha(fecha) {

    if (!fecha) {

        return "";

    }


    const texto =
        String(fecha).trim();


    const partes =
        texto.split("/");


    if (
        partes.length === 3
    ) {

        const dia =
            partes[0].padStart(
                2,
                "0"
            );


        const mes =
            partes[1].padStart(
                2,
                "0"
            );


        const anio =
            partes[2];


        return (

            anio +
            "-" +
            mes +
            "-" +
            dia

        );

    }


    if (
        /^\d{4}-\d{2}-\d{2}$/.test(
            texto
        )
    ) {

        return texto;

    }


    return texto;

}


// ----------------------------------------------------------
// FORMATEAR FECHA
// ----------------------------------------------------------

function formatearFecha(fecha) {

    if (!fecha) {

        return "";

    }


    const texto =
        String(fecha).trim();


    if (
        /^\d{4}-\d{2}-\d{2}$/.test(
            texto
        )
    ) {

        const partes =
            texto.split("-");


        return (

            partes[2] +
            "/" +
            partes[1] +
            "/" +
            partes[0]

        );

    }


    if (
        /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(
            texto
        )
    ) {

        return texto;

    }


    return texto;

}


// ----------------------------------------------------------
// NORMALIZAR TEXTO
// ----------------------------------------------------------

function normalizarTexto(texto) {

    return String(
        texto || ""
    )
        .toLowerCase()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .trim();

}


// ----------------------------------------------------------
// SEGURIDAD HTML
// ----------------------------------------------------------

function escaparHTML(texto) {

    return String(
        texto || ""
    )
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );

}


// ----------------------------------------------------------
// MOSTRAR ERROR
// ----------------------------------------------------------

function mostrarError(mensaje) {

    const tbody =
        document.getElementById(
            "tablaDatos"
        );


    if (tbody) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="12"
                    style="
                        text-align:center;
                        padding:30px;
                    "
                >

                    ⚠️
                    ${escaparHTML(
                        mensaje
                    )}

                </td>

            </tr>

        `;

    }


    ponerTexto(
        "totalActividades",
        "0"
    );


    ponerTexto(
        "completadas",
        "0"
    );


    ponerTexto(
        "pendientes",
        "0"
    );


    ponerTexto(
        "enProceso",
        "0"
    );


    ponerTexto(
        "cumplimiento",
        "0%"
    );

}