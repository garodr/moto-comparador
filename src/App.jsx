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

  const leerCarpetaDeExcel = async (e) => {
    const archivosLista = Array.from(e.target.files);
    if (archivosLista.length === 0) return;

    const archivosExcel = archivosLista.filter(
      (archivo) =>
        (archivo.name.endsWith(".xlsx") || archivo.name.endsWith(".xls")) &&
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
            const workbook = XLSX.read(data, { type: "binary" });
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
        reader.readAsBinaryString(archivo);
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
    <div className="min-h-screen bg-gray-50 p-3 md:p-6 text-gray-800 font-sans antialiased flex flex-col justify-between">
      <div className="max-w-xl mx-auto w-full flex-grow">
        {/* Cabecera limpia con Logo Centrado */}
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

        {/* Botón de Carga: Siempre visible arriba para comodidad */}
        {archivosCargados.length < 6 && (
          <div className="mb-4">
            <label className="flex flex-col items-center justify-center w-full h-12 border border-dashed border-blue-300 bg-blue-50/40 rounded-xl cursor-pointer active:bg-blue-100/70 transition-colors">
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

        {/* Mensajes de Alerta */}
        {errores.length > 0 && (
          <div className="bg-red-50 text-red-700 p-3 rounded-xl mb-4 border border-red-100 text-xs font-semibold space-y-1">
            {errores.map((err, i) => (
              <p key={i}>⚠️ {err}</p>
            ))}
          </div>
        )}

        {/* Buscador Pegajoso Superior */}
        {productos.length > 0 && (
          <div className="sticky top-2 z-20 mb-4 shadow-md rounded-xl">
            <input
              type="text"
              placeholder="Ej: 'bujia' o 'pastilla freno'..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              className="w-full p-3.5 rounded-xl border border-gray-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-base shadow-inner"
            />
          </div>
        )}

        {/* Listado Comparativo Principal */}
        <div className="space-y-3 mb-8">
          {productosAgrupados.slice(0, 50).map((grupo) => (
            <div
              key={grupo.idUnico}
              className="bg-white p-4 rounded-2xl shadow-sm border border-gray-200"
            >
              <h2 className="text-base font-black text-gray-950 leading-tight mb-3 uppercase">
                {grupo.detalle}
              </h2>

              <div className="bg-gray-50 rounded-xl border border-gray-100 divide-y divide-gray-100 overflow-hidden">
                {grupo.opciones.map((opc) => (
                  <div
                    key={`${opc.proveedor}-${opc.codigo}`}
                    className={`p-3 flex items-center justify-between transition-colors ${
                      opc.esElMasBarato ? "bg-green-50/70" : ""
                    }`}
                  >
                    <div className="truncate pr-2">
                      <div className="flex items-center gap-2 truncate">
                        <span className="text-sm font-bold text-gray-700 truncate">
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

                    <div className="text-right flex-shrink-0">
                      <span
                        className={`text-base font-black ${
                          opc.esElMasBarato ? "text-green-700" : "text-gray-900"
                        }`}
                      >
                        {formatearMonedaArgentina(opc.precioNum)}
                      </span>

                      {opc.diferencia && (
                        <span className="block text-[10px] font-bold text-amber-600">
                          {opc.diferencia} más caro
                        </span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}

          {productos.length > 0 && productosAgrupados.length === 0 && (
            <div className="text-center text-gray-400 py-10 bg-white rounded-xl border p-4 text-xs font-medium">
              No hay coincidencias para tu búsqueda comparativa.
            </div>
          )}
        </div>

        {/* SECCIÓN TRASLADADA: Listas en memoria abajo de todo */}
        {archivosCargados.length > 0 && (
          <div className="bg-white p-4 rounded-2xl shadow-sm border border-gray-200 mt-6">
            <div className="flex justify-between items-center mb-3">
              <span className="text-[10px] font-bold uppercase tracking-wider text-gray-400">
                Listas cargadas en memoria ({archivosCargados.length}/6)
              </span>
              <button
                onClick={limpiarTodo}
                className="text-xs font-bold text-red-500 active:text-red-700 bg-red-50 px-2 py-1 rounded-lg"
              >
                Resetear Todo
              </button>
            </div>

            <div className="grid grid-cols-1 gap-2">
              {archivosCargados.map((archivo) => (
                <div
                  key={archivo.id}
                  className="flex items-center justify-between bg-gray-50 p-2 rounded-xl border border-gray-150"
                >
                  <div className="truncate pr-2">
                    <p className="text-xs font-bold text-gray-600 truncate">
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
          </div>
        )}
      </div>
    </div>
  );
}
