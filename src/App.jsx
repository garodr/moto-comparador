import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

// IMPORTACIÓN ÚNICA: Logo unificado de la app
import logoApp from "./assets/logoapp.png";

// ==========================================
// FUNCIONES AUXILIARES (Heurísticas y Normalización)
// ==========================================

const normalizarTexto = (texto) => {
  if (!texto) return "";
  return String(texto)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remueve acentos
    .trim();
};

const detectarNombreProveedor = (nombreArchivo) => {
  const textoLimpio = nombreArchivo.toUpperCase();
  const proveedores = [
    "ARAX",
    "HADA",
    "REPCOR",
    "ROHAN",
    "CATALANO",
    "WSTANDARD",
    "STANDARD",
    "GIROLDI",
  ];

  const encontrado = proveedores.find((p) => textoLimpio.includes(p));
  if (encontrado) return encontrado;

  return nombreArchivo
    .replace(/\.[^/.]+$/, "")
    .replace(/LISTA DE PRECIOS/i, "")
    .replace(/MAYO|JUNIO|JULIO|AGOSTO|DE/i, "")
    .trim()
    .toUpperCase();
};

const obtenerValorFlexible = (item, textoBuscar) => {
  if (!item) return "No disponible";
  const terminoLimpio = normalizarTexto(textoBuscar);

  const claveReales = Object.keys(item).find((key) =>
    normalizarTexto(key).includes(terminoLimpio),
  );

  return claveReales && item[claveReales] !== undefined
    ? item[claveReales]
    : "No disponible";
};

const cumpleBusquedaSegura = (item, terminosBusqueda) => {
  if (terminosBusqueda.length === 0) return true;
  const stringItem = item._textoBusqueda || "";
  return terminosBusqueda.every((termino) => stringItem.includes(termino));
};

const formatearMonedaArgentina = (numero) => {
  if (numero === null || isNaN(numero)) return "S/D";
  const numeroRedondeado = Math.round(numero);
  return (
    "$" +
    numeroRedondeado.toLocaleString("es-AR", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    })
  );
};

const parsearPrecio = (precioRaw) => {
  if (!precioRaw || precioRaw === "No disponible") return null;

  let limpio = String(precioRaw).trim();
  if (limpio.includes(",") && limpio.includes(".")) {
    limpio = limpio.replace(/\./g, "").replace(/,/g, ".");
  } else if (limpio.includes(",")) {
    limpio = limpio.replace(/,/g, ".");
  }

  limpio = limpio.replace(/[^0-9.-]+/g, "");
  const numero = parseFloat(limpio);
  return isNaN(numero) ? null : numero;
};

const obtenerOpcionesValidas = (opciones) => {
  return opciones.filter((o) => o.precioNum !== null && !isNaN(o.precioNum));
};

// HEURÍSTICA DE AGRUPACIÓN POR DETALLE
const agruparPorProducto = (itemsFiltrados) => {
  const grupos = {};

  itemsFiltrados.forEach((item) => {
    const detalle = obtenerValorFlexible(item, "DETALLE");
    const codigo = obtenerValorFlexible(item, "CODIGO");

    const detalleLimpioBase =
      detalle !== "No disponible"
        ? detalle.replace(/\s*\([^)]*\)\s*$/, "").trim()
        : "Repuesto sin descripción";

    const claveGrupo = `DET-${normalizarTexto(detalleLimpioBase)}`;

    const precioRaw = obtenerValorFlexible(item, "PRECIO FINAL");
    const precioFinalNum = parsearPrecio(precioRaw);

    const ofertaProveedor = {
      proveedor: item.proveedorOrigen || "Desconocido",
      codigo: codigo !== "No disponible" ? codigo : null,
      precioNum: precioFinalNum,
      precioLista: obtenerValorFlexible(item, "PRECIO LISTA"),
      contado: obtenerValorFlexible(item, "CONTADO"),
    };

    if (!grupos[claveGrupo]) {
      grupos[claveGrupo] = {
        idUnico: claveGrupo,
        detalle: detalleLimpioBase,
        opciones: [],
      };
    }
    grupos[claveGrupo].opciones.push(ofertaProveedor);
  });

  return Object.values(grupos).map((grupo) => {
    const opcionesConPrecio = obtenerOpcionesValidas(grupo.opciones);

    let precioMinimo = Infinity;
    let mejorProveedor = null;
    let mejorCodigo = null;

    if (opcionesConPrecio.length > 0) {
      opcionesConPrecio.forEach((opc) => {
        if (opc.precioNum !== null && opc.precioNum < precioMinimo) {
          precioMinimo = opc.precioNum;
          mejorProveedor = opc.proveedor;
          mejorCodigo = opc.codigo;
        }
      });
    }

    grupo.opciones = grupo.opciones.map((opc) => {
      const esMejor =
        opc.precioNum === precioMinimo &&
        opc.proveedor === mejorProveedor &&
        opc.codigo === mejorCodigo;
      let diferenciaTexto = "";

      if (opc.precioNum !== null && precioMinimo !== Infinity && !esMejor) {
        const difPorcentaje = Math.round(
          ((opc.precioNum - precioMinimo) / precioMinimo) * 100,
        );
        diferenciaTexto = `(+${difPorcentaje}%)`;
      }

      return {
        ...opc,
        esElMasBarato: esMejor,
        diferencia: diferenciaTexto,
      };
    });

    grupo.opciones.sort(
      (a, b) => (a.precioNum || Infinity) - (b.precioNum || Infinity),
    );

    return grupo;
  });
};

// ==========================================
// COMPONENTE PRINCIPAL
// ==========================================
export default function App() {
  const [productos, setProductos] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [errores, setErrores] = useState([]);
  const [archivosCargados, setArchivosCargados] = useState([]);

  // Estados de Configuración y Menú
  const [menuAbierto, setMenuAbierto] = useState(false);
  const [darkMode, setDarkMode] = useState(false);

  // Doble Margen de Ganancia
  const [gananciaA, setGananciaA] = useState(0);
  const [gananciaB, setGananciaB] = useState(0);

  // Estados de Modales / Alertas
  const [modalGananciaAbierto, setModalGananciaAbierto] = useState(false);
  const [modalAcercaDe, setModalAcercaDe] = useState(false);

  // Inputs temporales para el Modal
  const [inputGananciaA, setInputGananciaA] = useState("0");
  const [inputGananciaB, setInputGananciaB] = useState("0");

  const leerCarpetaDeExcel = async (e) => {
    const archivosLista = Array.from(e.target.files);
    if (archivosLista.length === 0) return;

    const archivosExcel = archivosLista.filter(
      (archivo) =>
        (archivo.name.toLowerCase().endsWith(".xlsx") ||
          archivo.name.toLowerCase().endsWith(".xls")) &&
        !archivo.name.startsWith("~$"),
    );

    if (archivosExcel.length === 0) {
      setErrores([
        "No se encontraron archivos de Excel (.xlsx o .xls) válidos en la carpeta elegida.",
      ]);
      e.target.value = "";
      return;
    }

    setErrores([]);
    let nuevosProductos = [];
    let nuevosArchivosCargados = [...archivosCargados];
    let listaErrores = [];

    for (const archivo of archivosExcel) {
      if (nuevosArchivosCargados.length >= 6) {
        listaErrores.push(
          "Se alcanzó el límite máximo de 6 listas. Algunos archivos se omitieron.",
        );
        break;
      }

      if (nuevosArchivosCargados.some((a) => a.nombre === archivo.name)) {
        continue;
      }

      const promesaLectura = new Promise((resolve) => {
        const reader = new FileReader();

        reader.onload = (evt) => {
          try {
            const data = evt.target.result;
            let workbook;

            // INTENTO DE LECTURA ROBUSTA HÍBRIDA (Soporta .xls binarios viejos y .xlsx nuevos)
            try {
              workbook = XLSX.read(data, { type: "binary" });
            } catch {
              const bytes = new Uint8Array(data);
              workbook = XLSX.read(bytes, { type: "array" });
            }

            const nombreHoja = workbook.SheetNames[0];
            const hoja = workbook.Sheets[nombreHoja];

            const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });
            if (filas.length === 0) {
              resolve({ error: `"${archivo.name}" está vacío.` });
              return;
            }

            const indexEncabezados = filas.findIndex((fila) =>
              fila.some((celda) => {
                const textoCelda = normalizarTexto(celda);
                return (
                  textoCelda.includes("codigo de articulo") ||
                  textoCelda.includes("codigo") ||
                  textoCelda.includes("detalle") ||
                  textoCelda.includes("descripcion")
                );
              }),
            );

            const filaInicioDatos =
              indexEncabezados !== -1 ? indexEncabezados : 0;
            const datos = XLSX.utils.sheet_to_json(hoja, {
              range: filaInicioDatos,
            });

            if (datos.length === 0) {
              resolve({
                error: `"${archivo.name}" no tiene filas estructuradas.`,
              });
              return;
            }

            const nombreProveedorLimpio = detectarNombreProveedor(archivo.name);
            const idUnico = `${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;

            const datosConProveedor = datos.map((item) => {
              const codigo = obtenerValorFlexible(item, "CODIGO");
              const detalle = obtenerValorFlexible(item, "DETALLE");
              const textoIndexado = normalizarTexto(
                `${codigo} ${detalle} ${nombreProveedorLimpio}`,
              );

              return {
                ...item,
                archivoId: idUnico,
                proveedorOrigen: nombreProveedorLimpio,
                _textoBusqueda: textoIndexado,
              };
            });

            resolve({
              productos: datosConProveedor,
              archivoInfo: {
                id: idUnico,
                nombre: archivo.name,
                proveedor: nombreProveedorLimpio,
              },
            });
          } catch {
            resolve({ error: `Error de formato en "${archivo.name}".` });
          }
        };

        // Si es un .xls viejo, se lee mejor con readAsArrayBuffer para no romper codificaciones
        if (archivo.name.toLowerCase().endsWith(".xls")) {
          reader.readAsArrayBuffer(archivo);
        } else {
          reader.readAsBinaryString(archivo);
        }
      });

      const resultado = await promesaLectura;
      if (resultado.error) {
        listaErrores.push(resultado.error);
      } else if (resultado.productos) {
        nuevosProductos = [...nuevosProductos, ...resultado.productos];
        nuevosArchivosCargados.push(resultado.archivoInfo);
      }
    }

    if (nuevosProductos.length > 0) {
      setProductos((prev) => [...prev, ...nuevosProductos]);
      setArchivosCargados(nuevosArchivosCargados);
    }

    if (listaErrores.length > 0) {
      setErrores(listaErrores);
    }

    e.target.value = "";
  };

  const eliminarListaIndividual = (idParaEliminar) => {
    setProductos((prevProductos) =>
      prevProductos.filter((p) => p.archivoId !== idParaEliminar),
    );
    setArchivosCargados((prevArchivos) =>
      prevArchivos.filter((a) => a.id !== idParaEliminar),
    );
    setErrores([]);
  };

  const limpiarTodo = () => {
    setProductos([]);
    setArchivosCargados([]);
    setErrores([]);
    setBusqueda("");
  };

  const guardarGanancias = () => {
    const valA = parseFloat(inputGananciaA);
    const valB = parseFloat(inputGananciaB);
    setGananciaA(isNaN(valA) || valA < 0 ? 0 : valA);
    setGananciaB(isNaN(valB) || valB < 0 ? 0 : valB);
    setModalGananciaAbierto(false);
  };

  const abrirConsultasWhatsapp = () => {
    window.open(
      "https://wa.me/5493512345678?text=Hola!%20Tengo%20una%20consulta%20sobre%20Motolist",
      "_blank",
    );
    setMenuAbierto(false);
  };

  const terminosBusqueda = useMemo(() => {
    return normalizarTexto(busqueda)
      .split(" ")
      .filter((t) => t.length > 0);
  }, [busqueda]);

  const rawFiltrados = useMemo(() => {
    return productos.filter((item) =>
      cumpleBusquedaSegura(item, terminosBusqueda),
    );
  }, [productos, terminosBusqueda]);

  const productosAgrupados = useMemo(() => {
    return agruparPorProducto(rawFiltrados);
  }, [rawFiltrados]);

  return (
    <div
      className={`min-h-screen p-3 md:p-6 font-sans antialiased flex flex-col justify-between transition-colors duration-300 ${
        darkMode ? "bg-gray-900 text-gray-100" : "bg-gray-50 text-gray-800"
      }`}
    >
      {/* MENÚ HAMBURGUESA */}
      <div className="fixed top-3 left-3 z-50">
        <button
          onClick={() => setMenuAbierto(!menuAbierto)}
          className={`p-2.5 rounded-xl shadow-md border focus:outline-none transition-all ${
            darkMode
              ? "bg-gray-800 border-gray-700 text-white"
              : "bg-white border-gray-200 text-gray-800"
          }`}
        >
          {menuAbierto ? "✕" : "☰"}
        </button>

        {menuAbierto && (
          <div
            className={`absolute top-14 left-0 w-52 rounded-2xl shadow-xl border p-2 animate-fadeIn ${
              darkMode
                ? "bg-gray-800 border-gray-700 text-white"
                : "bg-white border-gray-100 text-gray-800"
            }`}
          >
            <button
              onClick={() => {
                setModalGananciaAbierto(true);
                setMenuAbierto(false);
              }}
              className={`w-full text-left p-2.5 text-xs font-bold rounded-xl flex items-center gap-2 ${darkMode ? "hover:bg-gray-700" : "hover:bg-gray-50"}`}
            >
              📈 <span>Editar ganancia</span>
              {(gananciaA > 0 || gananciaB > 0) && (
                <span className="ml-auto bg-blue-600 text-white px-1 py-0.5 text-[8px] rounded font-black">
                  SET
                </span>
              )}
            </button>

            <button
              onClick={() => setDarkMode(!darkMode)}
              className={`w-full text-left p-2.5 text-xs font-bold rounded-xl flex items-center gap-2 ${darkMode ? "hover:bg-gray-700" : "hover:bg-gray-50"}`}
            >
              🌙 <span>Modo oscuro: {darkMode ? "ON" : "OFF"}</span>
            </button>

            <button
              onClick={abrirConsultasWhatsapp}
              className={`w-full text-left p-2.5 text-xs font-bold rounded-xl flex items-center gap-2 ${darkMode ? "hover:bg-gray-700" : "hover:bg-gray-50"}`}
            >
              💬 <span>Consultas</span>
            </button>

            <button
              onClick={() => {
                setModalAcercaDe(true);
                setMenuAbierto(false);
              }}
              className={`w-full text-left p-2.5 text-xs font-bold rounded-xl flex items-center gap-2 ${darkMode ? "hover:bg-gray-700" : "hover:bg-gray-50"}`}
            >
              ℹ️ <span>Acerca de</span>
            </button>
          </div>
        )}
      </div>

      {/* CUERPO CENTRAL */}
      <div className="max-w-xl mx-auto w-full flex-grow">
        <header className="mb-4 flex flex-col items-center">
          <img
            src={logoApp}
            alt="Motolist Logo"
            className="h-20 w-auto object-contain mb-1"
          />
          <p className="text-[11px] text-gray-400 font-medium text-center tracking-wide">
            Buscá un repuesto y encontrá al instante el mejor proveedor.
          </p>
        </header>

        {/* Botón de Carga */}
        {archivosCargados.length < 6 && (
          <div className="mb-4">
            <label className="flex flex-col items-center justify-center w-full h-12 border border-dashed rounded-xl cursor-pointer transition-colors border-blue-300 bg-blue-50/40 active:bg-blue-100/70">
              <span className="text-xs font-bold text-blue-700">
                📂 Cargar carpeta con listas
              </span>
              <input
                type="file"
                webkitdirectory=""
                directory=""
                onChange={leerCarpetaDeExcel}
                className="hidden"
              />
            </label>
          </div>
        )}

        {errores.length > 0 && (
          <div className="bg-red-50 text-red-700 p-3 rounded-xl mb-4 border border-red-100 text-xs font-semibold space-y-1">
            {errores.map((err, i) => (
              <p key={i}>⚠️ {err}</p>
            ))}
          </div>
        )}

        {/* Buscador */}
        {productos.length > 0 && (
          <div className="sticky top-2 z-20 mb-4 shadow-md rounded-xl">
            <input
              type="text"
              placeholder="Ej: 'bujia' o 'pastilla freno'..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              className={`w-full p-3.5 rounded-xl border text-base shadow-inner focus:outline-none focus:ring-2 focus:ring-blue-500 ${
                darkMode
                  ? "bg-gray-800 border-gray-700 text-white placeholder-gray-500"
                  : "bg-white border-gray-300 text-gray-800"
              }`}
            />
          </div>
        )}

        {/* Listado Comparativo */}
        <div className="space-y-3 mb-8">
          {productosAgrupados.slice(0, 50).map((grupo) => (
            <div
              key={grupo.idUnico}
              className={`p-4 rounded-2xl shadow-sm border ${darkMode ? "bg-gray-800 border-gray-700" : "bg-white border-gray-200"}`}
            >
              <h2
                className={`text-base font-black leading-tight mb-3 uppercase ${darkMode ? "text-white" : "text-gray-950"}`}
              >
                {grupo.detalle}
              </h2>

              <div
                className={`rounded-xl border divide-y overflow-hidden ${darkMode ? "bg-gray-900/50 border-gray-700 divide-gray-700" : "bg-gray-50 border-gray-100 divide-gray-100"}`}
              >
                {grupo.opciones.map((opc) => (
                  <div
                    key={`${opc.proveedor}-${opc.codigo}`}
                    className={`p-3 flex items-center justify-between transition-colors ${opc.esElMasBarato ? (darkMode ? "bg-green-950/40" : "bg-green-50/70") : ""}`}
                  >
                    <div className="truncate pr-2">
                      <div className="flex items-center gap-2 truncate">
                        <span
                          className={`text-sm font-bold truncate ${darkMode ? "text-gray-200" : "text-gray-700"}`}
                        >
                          {opc.proveedor}
                        </span>
                        {opc.esElMasBarato && (
                          <span className="bg-green-600 text-white font-extrabold px-1.5 py-0.5 rounded-md text-[9px] uppercase tracking-wider">
                            Recomendado ✅
                          </span>
                        )}
                      </div>
                      {opc.codigo && (
                        <span className="block text-[10px] font-mono text-gray-400 mt-0.5">
                          Cód: {opc.codigo}
                        </span>
                      )}
                    </div>

                    <div className="text-right flex-shrink-0 space-y-0.5">
                      <span
                        className={`text-sm font-bold block ${opc.esElMasBarato ? "text-green-600" : darkMode ? "text-gray-400" : "text-gray-500"}`}
                      >
                        Costo: {formatearMonedaArgentina(opc.precioNum)}
                      </span>

                      {gananciaA > 0 && opc.precioNum !== null && (
                        <span className="block text-xs font-black text-blue-500">
                          PVP A (+{gananciaA}%):{" "}
                          {formatearMonedaArgentina(
                            opc.precioNum * (1 + gananciaA / 100),
                          )}
                        </span>
                      )}
                      {gananciaB > 0 && opc.precioNum !== null && (
                        <span className="block text-xs font-black text-purple-500">
                          PVP B (+{gananciaB}%):{" "}
                          {formatearMonedaArgentina(
                            opc.precioNum * (1 + gananciaB / 100),
                          )}
                        </span>
                      )}

                      {opc.diferencia && (
                        <span className="block text-[10px] font-bold text-amber-500">
                          {opc.diferencia} más caro
                        </span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>

        {/* LISTAS CARGADAS ABAJO */}
        {archivosCargados.length > 0 && (
          <div
            className={`p-4 rounded-2xl shadow-sm border mt-6 ${darkMode ? "bg-gray-800 border-gray-700" : "bg-white border-gray-200"}`}
          >
            <div className="flex justify-between items-center mb-3">
              <span className="text-[10px] font-bold uppercase tracking-wider text-gray-400">
                Listas en memoria ({archivosCargados.length}/6)
              </span>
              <button
                onClick={limpiarTodo}
                className="text-xs font-bold text-red-500 bg-red-50 px-2 py-1 rounded-lg"
              >
                Resetear Todo
              </button>
            </div>
            <div className="grid grid-cols-1 gap-2 mb-3">
              {archivosCargados.map((archivo) => (
                <div
                  key={archivo.id}
                  className={`flex items-center justify-between p-2 rounded-xl border ${darkMode ? "bg-gray-900/50 border-gray-700" : "bg-gray-50 border-gray-150"}`}
                >
                  <div className="truncate pr-2">
                    <p
                      className={`text-xs font-bold truncate ${darkMode ? "text-gray-300" : "text-gray-600"}`}
                    >
                      {archivo.proveedor}
                    </p>
                    <p className="text-[9px] text-gray-400 truncate">
                      {archivo.nombre}
                    </p>
                  </div>
                  <button
                    onClick={() => eliminarListaIndividual(archivo.id)}
                    className="text-xs font-black text-gray-400 hover:text-red-600 px-2 py-1"
                  >
                    ✕
                  </button>
                </div>
              ))}
            </div>
            <div
              className={`border rounded-xl p-2.5 text-center ${darkMode ? "bg-blue-950/20 border-blue-900/50 text-blue-400" : "bg-blue-50/50 border-blue-100 text-blue-800"}`}
            >
              <p className="text-xs font-bold">
                📊 Se cargaron un total de{" "}
                <span className="text-sm font-black">
                  {productos.length.toLocaleString("es-AR")}
                </span>{" "}
                artículos para comparar.
              </p>
            </div>
          </div>
        )}
      </div>

      {/* MODAL: EDITAR DOBLE GANANCIA */}
      {modalGananciaAbierto && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-xs">
          <div
            className={`w-full max-w-xs rounded-2xl p-5 shadow-2xl ${darkMode ? "bg-gray-800 text-white" : "bg-white text-gray-800"}`}
          >
            <h3 className="text-sm font-black uppercase tracking-wide mb-1">
              📈 Ajustar Margen Doble
            </h3>
            <p className="text-[11px] text-gray-400 mb-4">
              Configurá dos porcentajes de recargo distintos para tus repuestos.
            </p>

            <div className="space-y-4 mb-5">
              <div>
                <label className="block text-[10px] font-bold uppercase tracking-wider text-gray-400 mb-1">
                  Ganancia A (PVP Mostrador)
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={inputGananciaA}
                    onChange={(e) => setInputGananciaA(e.target.value)}
                    className={`w-full p-2 rounded-xl text-center font-bold border text-base focus:outline-none ${darkMode ? "bg-gray-700 border-gray-600 text-white" : "bg-gray-50 border-gray-300"}`}
                  />
                  <span className="text-base font-bold text-blue-500">%</span>
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-bold uppercase tracking-wider text-gray-400 mb-1">
                  Ganancia B (PVP Gremio / Taller)
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={inputGananciaB}
                    onChange={(e) => setInputGananciaB(e.target.value)}
                    className={`w-full p-2 rounded-xl text-center font-bold border text-base focus:outline-none ${darkMode ? "bg-gray-700 border-gray-600 text-white" : "bg-gray-50 border-gray-300"}`}
                  />
                  <span className="text-base font-bold text-purple-500">%</span>
                </div>
              </div>
            </div>

            <div className="flex gap-2">
              <button
                onClick={() => setModalGananciaAbierto(false)}
                className={`flex-1 py-2 text-xs font-bold rounded-xl ${darkMode ? "bg-gray-700" : "bg-gray-100"}`}
              >
                Cancelar
              </button>
              <button
                onClick={guardarGanancias}
                className="flex-1 py-2 text-xs font-bold bg-blue-600 text-white rounded-xl"
              >
                Aplicar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: ACERCA DE */}
      {modalAcercaDe && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-xs">
          <div
            className={`w-full max-w-xs rounded-2xl p-5 shadow-2xl text-center ${darkMode ? "bg-gray-800 text-white" : "bg-white text-gray-800"}`}
          >
            <img
              src={logoApp}
              alt="Logo"
              className="h-14 mx-auto mb-2 object-contain"
            />
            <h3 className="text-sm font-black uppercase">
              Motolist Comparador
            </h3>
            <p className="text-[11px] text-gray-400 mt-1 mb-4">
              Versión 2.6 (Anti-Crash)
            </p>
            <p className="text-xs leading-relaxed text-left mb-4">
              Herramienta inteligente de optimización de costos para repuestos
              de motos. Diseñada para unificar catálogos de proveedores en
              segundos y garantizar el mejor precio.
            </p>
            <button
              onClick={() => setModalAcercaDe(false)}
              className="w-full p-2 text-xs font-bold bg-blue-600 text-white rounded-xl"
            >
              Entendido
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
