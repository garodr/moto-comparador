import { useState } from "react";
import * as XLSX from "xlsx";

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
  const stringItem = normalizarTexto(
    JSON.stringify(item) + (item.proveedorOrigen || ""),
  );
  return terminosBusqueda.every((termino) => stringItem.includes(termino));
};

// HEURÍSTICA DE AGRUPACIÓN: Junta los repuestos repetidos entre proveedores
const agruparPorProducto = (itemsFiltrados) => {
  const grupos = {};

  itemsFiltrados.forEach((item) => {
    const detalle = obtenerValorFlexible(item, "DETALLE");
    const codigo = obtenerValorFlexible(item, "CODIGO");

    // Creamos una clave única para agrupar. Priorizamos código, si no tiene, usamos el detalle normalizado.
    const claveGrupo =
      codigo !== "No disponible"
        ? `COD-${normalizarTexto(codigo)}`
        : `DET-${normalizarTexto(detalle)}`;

    // Extraemos el precio final de forma numérica para poder comparar matemáticamente
    const precioRaw = obtenerValorFlexible(item, "PRECIO FINAL");
    const precioFinalNum =
      precioRaw !== "No disponible"
        ? parseFloat(String(precioRaw).replace(/[^0-9.-]+/g, ""))
        : null;

    const ofertaProveedor = {
      proveedor: item.proveedorOrigen || "Desconocido",
      codigo: codigo,
      precioOriginal: precioRaw,
      precioNum: precioFinalNum,
      precioLista: obtenerValorFlexible(item, "PRECIO LISTA"),
      contado: obtenerValorFlexible(item, "CONTADO"),
    };

    if (!grupos[claveGrupo]) {
      grupos[claveGrupo] = {
        detalle:
          detalle !== "No disponible" ? detalle : "Repuesto sin descripción",
        codigoPrincipal: codigo !== "No disponible" ? codigo : null,
        opciones: [],
      };
    }
    grupos[claveGrupo].opciones.push(ofertaProveedor);
  });

  // Procesamos cada grupo para marcar el mejor precio y calcular diferencias
  return Object.values(grupos).map((grupo) => {
    // Filtramos solo las opciones que tengan un precio numérico válido
    const opcionesConPrecio = grupo.opciones.filter(
      (o) => o.precioNum !== null && !isNaN(o.precioNum),
    );

    let precioMinimo = Infinity;
    let mejorProveedor = null;

    if (opcionesConPrecio.length > 0) {
      // Encontramos el precio más bajo
      opcionesConPrecio.forEach((opc) => {
        if (opc.precioNum < precioMinimo) {
          precioMinimo = opc.precioNum;
          mejorProveedor = opc.proveedor;
        }
      });
    }

    // Le inyectamos a cada opción si es la ganadora y su diferencia porcentual
    grupo.opciones = grupo.opciones.map((opc) => {
      const esMejor =
        opc.precioNum === precioMinimo && opc.proveedor === mejorProveedor;
      let diferenciaTexto = "";

      if (opc.precioNum && precioMinimo !== Infinity && !esMejor) {
        // Cálculo matemático de la diferencia: ((PrecioActual - PrecioMinimo) / PrecioMinimo) * 100
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

    // Ordenamos las opciones de la tarjeta para que el más barato aparezca siempre arriba de todo
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

  const leerUnExcel = (e) => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    if (archivosCargados.length >= 5) {
      setErrores([
        "Ya cargaste el máximo de 5 listas. Borrá alguna para sumar otra.",
      ]);
      e.target.value = "";
      return;
    }

    if (archivosCargados.some((a) => a.nombre === archivo.name)) {
      setErrores([`El archivo "${archivo.name}" ya está cargado.`]);
      e.target.value = "";
      return;
    }

    setErrores([]);
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const nombreHoja = workbook.SheetNames[0];
        const hoja = workbook.Sheets[nombreHoja];

        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });

        if (filas.length === 0) {
          setErrores([`El archivo "${archivo.name}" está vacío.`]);
          e.target.value = "";
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

        const filaInicioDatos = indexEncabezados !== -1 ? indexEncabezados : 0;
        const datos = XLSX.utils.sheet_to_json(hoja, {
          range: filaInicioDatos,
        });

        if (datos.length === 0) {
          setErrores([`"${archivo.name}" no tiene filas válidas.`]);
          e.target.value = "";
          return;
        }

        const nombreProveedor = archivo.name.replace(/\.[^/.]+$/, "");
        const idUnico = Date.now().toString();

        const datosConProveedor = datos.map((item) => ({
          ...item,
          archivoId: idUnico,
          proveedorOrigen: nombreProveedor,
        }));

        setProductos((prevProductos) => [
          ...prevProductos,
          ...datosConProveedor,
        ]);
        setArchivosCargados((prevArchivos) => [
          ...prevArchivos,
          { id: idUnico, nombre: archivo.name, proveedor: nombreProveedor },
        ]);

        e.target.value = "";
      } catch (err) {
        console.error(err);
        setErrores([`Error de formato en "${archivo.name}".`]);
        e.target.value = "";
      }
    };

    reader.readAsBinaryString(archivo);
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

  // 1. Filtramos los productos según la barra de búsqueda básica
  const terminosBusqueda = normalizarTexto(busqueda)
    .split(" ")
    .filter((t) => t.length > 0);
  const rawFiltrados = productos.filter((item) =>
    cumpleBusquedaSegura(item, terminosBusqueda),
  );

  // 2. Aplicamos la heurística visual para agrupar repetidos y encontrar el recomendado
  const productosAgrupados = agruparPorProducto(rawFiltrados);

  return (
    <div className="min-h-screen bg-gray-50 p-3 md:p-6 text-gray-800 font-sans antialiased">
      <div className="max-w-xl mx-auto">
        {/* Cabecera */}
        <header className="mb-5 text-center sm:text-left">
          <h1 className="text-2xl font-black tracking-tight text-gray-900">
            📊 Comparador de Precios Inteligente
          </h1>
          <p className="text-xs text-gray-500 mt-1">
            Buscá un repuesto y la app te dirá automáticamente cuál proveedor te
            conviene.
          </p>
        </header>

        {/* Zona de Carga */}
        <div className="bg-white p-4 rounded-2xl shadow-sm mb-4 border border-gray-200">
          <div className="flex justify-between items-center mb-3">
            <span className="text-xs font-bold uppercase tracking-wider text-gray-400">
              Listas en memoria ({archivosCargados.length}/5)
            </span>
            {archivosCargados.length > 0 && (
              <button
                onClick={limpiarTodo}
                className="text-xs font-bold text-red-500 active:text-red-700 bg-red-50 px-2.5 py-1 rounded-lg"
              >
                Borrar todas
              </button>
            )}
          </div>

          {archivosCargados.length < 5 ? (
            <label className="flex flex-col items-center justify-center w-full h-14 border-2 border-dashed border-blue-200 bg-blue-50/50 rounded-xl cursor-pointer active:bg-blue-100 transition-colors">
              <div className="flex items-center gap-2 text-sm font-bold text-blue-700">
                ➕ <span>Sumar lista de proveedor</span>
              </div>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={leerUnExcel}
                className="hidden"
              />
            </label>
          ) : (
            <div className="text-center p-3 bg-amber-50 border border-amber-200 text-amber-800 font-medium rounded-xl text-xs">
              🔒 Límite de 5 listas alcanzado.
            </div>
          )}

          {archivosCargados.length > 0 && (
            <div className="mt-4 pt-3 border-t border-gray-100 space-y-2">
              {archivosCargados.map((archivo) => (
                <div
                  key={archivo.id}
                  className="flex items-center justify-between bg-gray-50 p-2.5 rounded-xl border border-gray-200"
                >
                  <div className="truncate pr-3">
                    <p className="text-xs font-bold text-gray-700 truncate">
                      {archivo.proveedor}
                    </p>
                    <p className="text-[10px] text-gray-400 truncate">
                      {archivo.nombre}
                    </p>
                  </div>
                  <button
                    onClick={() => eliminarListaIndividual(archivo.id)}
                    className="p-1 px-2.5 text-xs font-black text-gray-400 active:text-red-600 rounded-lg"
                  >
                    ✕
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Mensajes de Alerta */}
        {errores.length > 0 && (
          <div className="bg-red-50 text-red-700 p-3.5 rounded-xl mb-4 border border-red-100 text-xs font-semibold space-y-1">
            {errores.map((err, i) => (
              <p key={i}>⚠️ {err}</p>
            ))}
          </div>
        )}

        {/* Buscador Pegajoso */}
        {productos.length > 0 && (
          <div className="sticky top-2 z-20 my-3 shadow-md rounded-xl">
            <input
              type="text"
              placeholder="Ej: 'bujia' o 'pastilla freno'..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              className="w-full p-3.5 rounded-xl border border-gray-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-base shadow-inner"
            />
          </div>
        )}

        {/* Listado Comparativo Visual (Optimizado para Celular) */}
        <div className="space-y-3">
          {productosAgrupados.slice(0, 50).map((grupo, idx) => (
            <div
              key={idx}
              className="bg-white p-4 rounded-2xl shadow-sm border border-gray-200"
            >
              {/* Cabecera del Repuesto Único */}
              {grupo.codigoPrincipal && (
                <span className="inline-block text-[10px] font-mono bg-gray-100 px-1.5 py-0.5 rounded text-gray-500 mb-1">
                  Cód: {grupo.codigoPrincipal}
                </span>
              )}
              <h2 className="text-base font-black text-gray-950 leading-tight mb-3">
                {grupo.detalle}
              </h2>

              {/* Contenedor de la Comparativa de Precios */}
              <div className="bg-gray-50 rounded-xl border border-gray-100 divide-y divide-gray-100 overflow-hidden">
                {grupo.opciones.map((opc, oIdx) => (
                  <div
                    key={oIdx}
                    className={`p-3 flex items-center justify-between transition-colors ${
                      opc.esElMasBarato ? "bg-green-50/70" : ""
                    }`}
                  >
                    {/* Info del Proveedor */}
                    <div className="flex items-center gap-2 truncate pr-2">
                      <span className="text-sm font-bold text-gray-700 truncate">
                        {opc.proveedor}
                      </span>
                      {opc.esElMasBarato && (
                        <span className="text-xs bg-green-600 text-white font-extrabold px-1.5 py-0.5 rounded-md text-[9px] uppercase tracking-wider animate-pulse">
                          Recomendado ✅
                        </span>
                      )}
                    </div>

                    {/* Precios y Diferenciales */}
                    <div className="text-right flex-shrink-0">
                      <span
                        className={`text-base font-black ${
                          opc.esElMasBarato ? "text-green-700" : "text-gray-900"
                        }`}
                      >
                        $
                        {opc.precioOriginal !== "No disponible"
                          ? opc.precioOriginal
                          : "S/D"}
                      </span>

                      {/* Porcentaje de recargo con respecto al más barato */}
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
      </div>
    </div>
  );
}
