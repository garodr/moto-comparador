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

  // Incluimos el nombre del proveedor en el string global de búsqueda para poder filtrar por proveedor también
  const stringItem = normalizarTexto(
    JSON.stringify(item) + (item.proveedorOrigen || ""),
  );

  return terminosBusqueda.every((termino) => stringItem.includes(termino));
};

// ==========================================
// COMPONENTE PRINCIPAL
// ==========================================
export default function App() {
  const [productos, setProductos] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [errores, setErrores] = useState([]);
  const [archivosCargados, setArchivosCargados] = useState([]);

  // Procesar un archivo individual (Retorna una Promesa para poder usarse en lote)
  const procesarArchivoIndividual = (archivo) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (evt) => {
        try {
          const data = evt.target.result;
          const workbook = XLSX.read(data, { type: "binary" });
          const nombreHoja = workbook.SheetNames[0];
          const hoja = workbook.Sheets[nombreHoja];

          const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });

          if (filas.length === 0) {
            reject(`El archivo "${archivo.name}" está vacío.`);
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
            reject(`"${archivo.name}" no tiene filas válidas de repuestos.`);
            return;
          }

          // Formateamos el nombre del archivo para usarlo como nombre de Proveedor
          const nombreProveedor = archivo.name.replace(/\.[^/.]+$/, "");

          // Inyectamos el nombre del proveedor a cada fila de este Excel
          const datosConProveedor = datos.map((item) => ({
            ...item,
            proveedorOrigen: nombreProveedor,
          }));

          resolve({ datos: datosConProveedor, nombre: archivo.name });
        } catch (err) {
          console.error(err); // Usado para que ESLint no proteste
          reject(`Error de formato en el archivo "${archivo.name}".`);
        }
      };

      reader.onerror = () =>
        reject(`No se pudo leer el archivo "${archivo.name}".`);
      reader.readAsBinaryString(archivo);
    });
  };

  // Manejador de la carga múltiple
  const leerMultiplesExcel = async (e) => {
    const archivos = Array.from(e.target.files);
    if (archivos.length === 0) return;

    // Validación: limitar a 5 listas para cuidar la memoria del celular
    if (archivos.length > 5) {
      setErrores(["Puedes cargar un máximo de 5 listas en simultáneo."]);
      return;
    }

    setErrores([]);

    // Procesamos todos los archivos en paralelo usando Promesas
    const promesas = archivos.map((archivo) =>
      procesarArchivoIndividual(archivo),
    );

    try {
      const resultados = await Promise.allSettled(promesas);

      let productosAcumulados = [];
      let nombresExitosos = [];
      let mensajesError = [];

      resultados.forEach((res) => {
        if (res.status === "fulfilled") {
          productosAcumulados = [...productosAcumulados, ...res.value.datos];
          nombresExitosos.push(res.value.nombre);
        } else {
          mensajesError.push(res.reason);
        }
      });

      // Guardamos todo el lote unificado en el estado
      setProductos(productosAcumulados);
      setArchivosCargados(nombresExitosos);
      if (mensajesError.length > 0) setErrores(mensajesError);
    } catch (err) {
      console.error(err); // Usado para que ESLint no proteste
      setErrores([
        "Ocurrió un error inesperado al procesar el lote de archivos.",
      ]);
    }
  };

  // Limpiar el comparador para volver a empezar
  const limpiarListas = () => {
    setProductos([]);
    setArchivosCargados([]);
    setErrores([]);
    setBusqueda("");
  };

  const terminosBusqueda = normalizarTexto(busqueda)
    .split(" ")
    .filter((t) => t.length > 0);
  const filtrados = productos.filter((item) =>
    cumpleBusquedaSegura(item, terminosBusqueda),
  );

  return (
    <div className="min-h-screen bg-gray-100 p-4 md:p-6 text-gray-800 font-sans">
      <div className="max-w-xl mx-auto">
        {/* Cabecera */}
        <header className="mb-6">
          <h1 className="text-3xl font-black tracking-tight flex items-center gap-2">
            📊 Comparador de Listas Multi-Proveedor
          </h1>
          <p className="text-sm text-gray-500 mt-1">
            Seleccioná hasta 5 archivos Excel juntos para cotejar precios al
            toque.
          </p>
        </header>

        {/* Zona de Carga Múltiple */}
        <div className="bg-white p-4 rounded-2xl shadow-sm mb-4 border border-gray-200">
          <label className="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-2">
            Seleccionar de 1 a 5 Archivos Excel:
          </label>
          <input
            type="file"
            accept=".xlsx, .xls"
            multiple
            onChange={leerMultiplesExcel}
            className="w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-bold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer"
          />

          {archivosCargados.length > 0 && (
            <div className="mt-3 pt-3 border-t border-gray-100 flex justify-between items-end">
              <div>
                <p className="text-xs text-gray-500 font-bold uppercase tracking-wider">
                  Listas activas:
                </p>
                <ul className="text-xs text-green-700 font-medium list-disc pl-4 mt-1 space-y-0.5">
                  {archivosCargados.map((n, i) => (
                    <li key={i}>{n}</li>
                  ))}
                </ul>
                <p className="text-xs text-gray-400 mt-1 font-semibold">
                  Total unificado: {productos.length} ítems
                </p>
              </div>
              <button
                onClick={limpiarListas}
                className="text-xs font-bold text-red-600 bg-red-50 px-3 py-1.5 rounded-xl hover:bg-red-100"
              >
                Borrar todo
              </button>
            </div>
          )}
        </div>

        {/* Mensajes de Validación/Errores */}
        {errores.length > 0 && (
          <div className="bg-red-50 text-red-700 p-4 rounded-2xl mb-4 border border-red-200 text-sm space-y-1">
            {errores.map((err, i) => (
              <p key={i}>⚠️ {err}</p>
            ))}
          </div>
        )}

        {/* Buscador Inteligente */}
        {productos.length > 0 && (
          <div className="sticky top-3 z-10 my-4 shadow-md rounded-2xl">
            <input
              type="text"
              placeholder="Ej: 'bujia' (coteja precios de todas las listas)..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              className="w-full p-4 rounded-2xl border border-gray-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-lg shadow-inner"
            />
          </div>
        )}

        {/* Listado de Tarjetas de Repuestos */}
        <div className="space-y-3">
          {filtrados.slice(0, 100).map((item, index) => {
            const detalle = obtenerValorFlexible(item, "DETALLE");
            const codigo = obtenerValorFlexible(item, "CODIGO");
            const precioLista = obtenerValorFlexible(item, "PRECIO LISTA");
            const precioFinal = obtenerValorFlexible(item, "PRECIO FINAL");
            const contado = obtenerValorFlexible(item, "CONTADO");
            const proveedor = item.proveedorOrigen || "Desconocido";

            return (
              <div
                key={index}
                className="bg-white p-4 rounded-2xl shadow-sm border border-gray-200 active:scale-[0.99] transition-transform relative overflow-hidden"
              >
                {/* Etiqueta flotante del Proveedor */}
                <span className="absolute top-3 right-3 text-[10px] font-black uppercase tracking-wider px-2 py-0.5 bg-blue-50 text-blue-700 rounded-md max-w-[150px] truncate">
                  📋 {proveedor}
                </span>

                {/* Código de Artículo */}
                {codigo !== "No disponible" && (
                  <span className="inline-block text-xs font-mono bg-gray-100 px-2 py-0.5 rounded text-gray-600 mb-1">
                    Cód: {codigo}
                  </span>
                )}

                {/* Detalle */}
                <h2 className="text-lg font-bold text-gray-950 leading-tight pr-24">
                  {detalle !== "No disponible"
                    ? detalle
                    : "Repuesto sin descripción"}
                </h2>

                {/* Precios secundarios */}
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

                {/* Precio Final Principal */}
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
                No se encontraron coincidencias.
              </p>
              <p className="text-sm text-gray-400 mt-1">
                Intentá con un término más genérico.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
