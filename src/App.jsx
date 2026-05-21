import { useState } from "react";
import * as XLSX from "xlsx";

// ==========================================
// FUNCIONES AUXILIARES (Heurísticas y Normalización)
// ==========================================

// Limpia el texto eliminando acentos, mayúsculas y espacios innecesarios
const normalizarTexto = (texto) => {
  if (!texto) return "";
  return String(texto)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remueve diacríticos (acentos)
    .trim();
};

// Tu función original mejorada con normalización para encontrar columnas sin importar mayúsculas/acentos
const obtenerValorFlexible = (item, textoBuscar) => {
  if (!item) return "No disponible";
  const terminoLimpio = normalizarTexto(textoBuscar);

  // Busca una clave en el objeto que contenga el texto que queremos mapear
  const claveReales = Object.keys(item).find((key) =>
    normalizarTexto(key).includes(terminoLimpio),
  );

  return claveReales && item[claveReales] !== undefined
    ? item[claveReales]
    : "No disponible";
};

// Filtro inteligente: verifica que cada palabra que tipeó el usuario esté en la fila (en cualquier orden)
const cumpleBusquedaSegura = (item, terminosBusqueda) => {
  if (terminosBusqueda.length === 0) return true;

  // Convertimos toda la fila a un string único normalizado para buscar de forma global
  const stringItem = normalizarTexto(JSON.stringify(item));

  // Heurística: cada pedazo de texto buscado debe existir dentro del objeto
  return terminosBusqueda.every((termino) => stringItem.includes(termino));
};

// ==========================================
// COMPONENTE PRINCIPAL
// ==========================================
export default function App() {
  const [productos, setProductos] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [error, setError] = useState("");
  const [nombreArchivo, setNombreArchivo] = useState("");

  const leerExcel = (e) => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    setNombreArchivo(archivo.name);
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const nombreHoja = workbook.SheetNames[0];
        const hoja = workbook.Sheets[nombreHoja];

        // Convertimos a matriz para validar el contenido e identificar encabezados
        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });

        // Validación 1: Archivo vacío
        if (filas.length === 0) {
          setError("El archivo Excel seleccionado está vacío.");
          return;
        }

        // Heurística para detectar la fila de encabezados (reutiliza tu lógica original pero con más flexibilidad)
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

        // Si la heurística no encuentra palabras clave, por defecto asume la primera fila (0)
        const filaInicioDatos = indexEncabezados !== -1 ? indexEncabezados : 0;

        // Parseo final usando el rango correcto detectado
        const datos = XLSX.utils.sheet_to_json(hoja, {
          range: filaInicioDatos,
        });

        // Validación 2: ¿Pudo procesar filas de datos?
        if (datos.length === 0) {
          setError(
            "No se encontraron filas de repuestos válidas por debajo del encabezado.",
          );
          return;
        }

        setProductos(datos);
        setError(""); // Limpiamos errores previos si todo salió bien
      } catch (err) {
        setError("Ocurrió un error al procesar el Excel. Verifica el formato.");
        console.error(err);
      }
    };

    reader.readAsBinaryString(archivo);
  };

  // Procesamos la barra de búsqueda una sola vez antes de mapear (Mejora el rendimiento en celular)
  const terminosBusqueda = normalizarTexto(busqueda)
    .split(" ")
    .filter((t) => t.length > 0);

  // Filtrado flexible
  const filtrados = productos.filter((item) =>
    cumpleBusquedaSegura(item, terminosBusqueda),
  );

  return (
    <div className="min-h-screen bg-gray-100 p-4 md:p-6 text-gray-800 font-sans">
      <div className="max-w-xl mx-auto">
        {/* Cabecera */}
        <header className="mb-6">
          <h1 className="text-3xl font-black tracking-tight flex items-center gap-2">
            🏍️ Buscador de Repuestos
          </h1>
          <p className="text-sm text-gray-500 mt-1">
            Cargá el Excel del proveedor y buscá al instante desde el taller.
          </p>
        </header>

        {/* Zona de Carga de Archivo */}
        <div className="bg-white p-4 rounded-2xl shadow-sm mb-4 border border-gray-200">
          <label className="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-2">
            Seleccionar Lista de Precios
          </label>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={leerExcel}
            className="w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-bold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer"
          />
          {nombreArchivo && !error && (
            <p className="text-xs text-green-600 font-medium mt-2 flex items-center gap-1">
              ✓ Archivo cargado:{" "}
              <span className="font-semibold">{nombreArchivo}</span> (
              {productos.length} ítems)
            </p>
          )}
        </div>

        {/* Mensajes de Validación */}
        {error && (
          <div className="bg-red-50 text-red-700 p-4 rounded-2xl mb-4 border border-red-200 text-sm font-medium">
            ⚠️ {error}
          </div>
        )}

        {/* Input de Búsqueda Inteligente */}
        {productos.length > 0 && (
          <div className="sticky top-3 z-10 my-4 shadow-md rounded-2xl">
            <input
              type="text"
              placeholder="Ej: 'bujia ngk titan' o 'pastilla freno cg'..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              className="w-full p-4 rounded-2xl border border-gray-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-lg shadow-inner"
            />
          </div>
        )}

        {/* Listado de Tarjetas de Repuestos */}
        <div className="space-y-3">
          {/* El .slice(0, 75) limita las tarjetas renderizadas simultáneamente para evitar tildar el navegador del celular */}
          {filtrados.slice(0, 75).map((item, index) => {
            const detalle = obtenerValorFlexible(item, "DETALLE");
            const codigo = obtenerValorFlexible(item, "CODIGO");
            const precioLista = obtenerValorFlexible(item, "PRECIO LISTA");
            const precioFinal = obtenerValorFlexible(item, "PRECIO FINAL");
            const contado = obtenerValorFlexible(item, "CONTADO");

            return (
              <div
                key={index}
                className="bg-white p-4 rounded-2xl shadow-sm border border-gray-200 active:scale-[0.99] transition-transform"
              >
                {/* Código de Artículo */}
                {codigo !== "No disponible" && (
                  <span className="inline-block text-xs font-mono bg-gray-100 px-2 py-0.5 rounded text-gray-600 mb-1">
                    Cód: {codigo}
                  </span>
                )}

                {/* Detalle o Descripción */}
                <h2 className="text-lg font-bold text-gray-950 leading-tight">
                  {detalle !== "No disponible"
                    ? detalle
                    : "Repuesto sin descripción"}
                </h2>

                {/* Precios secundarios (Lista y Contado) */}
                <div className="grid grid-cols-2 gap-2 text-sm mt-3 pt-2 border-t border-gray-100">
                  {precioLista !== "No disponible" && (
                    <div>
                      <span className="text-gray-400 block text-xs">
                        Precio Lista
                      </span>
                      <span className="font-semibold text-gray-700">
                        ${precioLista}
                      </span>
                    </div>
                  )}
                  {contado !== "No disponible" && (
                    <div>
                      <span className="text-blue-600 block text-xs font-medium">
                        Contado / Efvo
                      </span>
                      <span className="font-bold text-blue-700">
                        ${contado}
                      </span>
                    </div>
                  )}
                </div>

                {/* Precio Principal Destacado (Precio Final) */}
                {precioFinal !== "No disponible" && (
                  <div className="mt-3 bg-green-50 p-3 rounded-xl flex justify-between items-center border border-green-100">
                    <span className="text-green-800 text-xs font-bold uppercase tracking-wider">
                      Precio Final:
                    </span>
                    <span className="text-2xl font-black text-green-600">
                      ${precioFinal}
                    </span>
                  </div>
                )}
              </div>
            );
          })}

          {/* Feedback cuando no hay coincidencias */}
          {productos.length > 0 && filtrados.length === 0 && (
            <div className="text-center text-gray-500 py-12 bg-white rounded-2xl border p-4">
              <p className="text-lg font-medium">
                No encontramos ese repuesto.
              </p>
              <p className="text-sm text-gray-400 mt-1">
                Probá escribiendo menos palabras o revisando la ortografía.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
