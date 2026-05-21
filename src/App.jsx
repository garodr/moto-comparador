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

  // Incluimos el nombre del proveedor en el string global de búsqueda
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
  const [archivosCargados, setArchivosCargados] = useState([]); // Guardará objetos: { id, nombre, proveedor }

  // Manejador para cargar los Excel de a uno y acumularlos
  const leerUnExcel = (e) => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    // Validación: Máximo de 5 listas acumuladas
    if (archivosCargados.length >= 5) {
      setErrores([
        "Ya cargaste el máximo de 5 listas. Borrá alguna para sumar otra.",
      ]);
      e.target.value = "";
      return;
    }

    // Validación: Evitar duplicados por nombre
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

        // Formateamos el nombre del proveedor sacando la extensión (.xlsx)
        const nombreProveedor = archivo.name.replace(/\.[^/.]+$/, "");
        const idUnico = Date.now().toString(); // ID para poder borrarlo individualmente después

        // Inyectamos el ID y nombre del proveedor a cada fila de este Excel
        const datosConProveedor = datos.map((item) => ({
          ...item,
          archivoId: idUnico,
          proveedorOrigen: nombreProveedor,
        }));

        // Acumulamos en los estados correspondientes
        setProductos((prevProductos) => [
          ...prevProductos,
          ...datosConProveedor,
        ]);
        setArchivosCargados((prevArchivos) => [
          ...prevArchivos,
          { id: idUnico, nombre: archivo.name, proveedor: nombreProveedor },
        ]);

        e.target.value = ""; // Limpiamos el input para el próximo toque
      } catch (err) {
        console.error(err);
        setErrores([`Error de formato en "${archivo.name}".`]);
        e.target.value = "";
      }
    };

    reader.readAsBinaryString(archivo);
  };

  // FUNCIÓN CLAVE PARA CELULARES: Permite sacar una lista individual de la memoria sin resetear todo
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

  const terminosBusqueda = normalizarTexto(busqueda)
    .split(" ")
    .filter((t) => t.length > 0);
  const filtrados = productos.filter((item) =>
    cumpleBusquedaSegura(item, terminosBusqueda),
  );

  return (
    <div className="min-h-screen bg-gray-50 p-3 md:p-6 text-gray-800 font-sans antialiased">
      <div className="max-w-xl mx-auto">
        {/* Cabecera */}
        <header className="mb-5 text-center sm:text-left">
          <h1 className="text-2xl font-black tracking-tight text-gray-900">
            🏍️ Buscador Multi-Listas
          </h1>
          <p className="text-xs text-gray-500 mt-1">
            Cargá los Excel de a uno. Buscá y compará precios al instante.
          </p>
        </header>

        {/* Contenedor de Carga (Diseño optimizado para dedos/pantallas táctiles) */}
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

          {/* Botón de carga grande para el dedo */}
          {archivosCargados.length < 5 ? (
            <label className="flex flex-col items-center justify-center w-full h-14 border-2 border-dashed border-blue-200 bg-blue-50/50 rounded-xl cursor-pointer active:bg-blue-100 transition-colors">
              <div className="flex items-center gap-2 text-sm font-bold text-blue-700">
                ➕ <span>Toca acá para sumar un Excel</span>
              </div>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={leerUnExcel}
                className="hidden" // Escondemos el input feo nativo
              />
            </label>
          ) : (
            <div className="text-center p-3 bg-amber-50 border border-amber-200 text-amber-800 font-medium rounded-xl text-xs">
              🔒 Límite de 5 listas alcanzado.
            </div>
          )}

          {/* Panel de Gestión de Listas Activas (Tarjetitas individuales con botón de borrado) */}
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
                  {/* Botón X bien grande y fácil de presionar con el pulgar */}
                  <button
                    onClick={() => eliminarListaIndividual(archivo.id)}
                    className="p-1 px-2.5 text-xs font-black text-gray-400 hover:text-red-600 active:bg-red-50 rounded-lg transition-colors"
                    title="Eliminar esta lista"
                  >
                    ✕
                  </button>
                </div>
              ))}
              <p className="text-[11px] text-gray-400 text-right font-medium mt-1">
                Total acumulado para buscar:{" "}
                <span className="font-bold text-gray-700">
                  {productos.length}
                </span>{" "}
                repuestos.
              </p>
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

        {/* Buscador Inteligente Pegajoso (Se queda fijo arriba al hacer scroll) */}
        {productos.length > 0 && (
          <div className="sticky top-2 z-20 my-3 shadow-md rounded-xl">
            <input
              type="text"
              placeholder="Escribí marca, modelo o código..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              className="w-full p-3.5 rounded-xl border border-gray-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-base shadow-inner"
            />
          </div>
        )}

        {/* Tarjetas de Repuestos */}
        <div className="space-y-2.5">
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
                className="bg-white p-3.5 rounded-xl shadow-sm border border-gray-200 active:bg-gray-50 active:scale-[0.99] transition-all relative overflow-hidden"
              >
                {/* Nombre de la Lista / Proveedor */}
                <span className="absolute top-3 right-3 text-[9px] font-black uppercase tracking-wider px-2 py-0.5 bg-blue-50 text-blue-700 rounded-md max-w-[130px] truncate">
                  📋 {proveedor}
                </span>

                {/* Código */}
                {codigo !== "No disponible" && (
                  <span className="inline-block text-[10px] font-mono bg-gray-100 px-1.5 py-0.5 rounded text-gray-500 mb-1">
                    Cód: {codigo}
                  </span>
                )}

                {/* Detalle */}
                <h2 className="text-base font-bold text-gray-900 leading-tight pr-24">
                  {detalle !== "No disponible"
                    ? detalle
                    : "Repuesto sin descripción"}
                </h2>

                {/* Precios secundarios */}
                <div className="grid grid-cols-2 gap-2 text-xs mt-2.5 pt-2 border-t border-gray-100">
                  {precioLista !== "No disponible" && (
                    <div>
                      <span className="text-gray-400 block text-[10px]">
                        Precio Lista
                      </span>
                      <span className="font-semibold text-gray-600">
                        ${precioLista}
                      </span>
                    </div>
                  )}
                  {contado !== "No disponible" && (
                    <div>
                      <span className="text-blue-600 block text-[10px] font-medium">
                        Contado / Efvo
                      </span>
                      <span className="font-bold text-blue-700">
                        ${contado}
                      </span>
                    </div>
                  )}
                </div>

                {/* Precio Principal */}
                {precioFinal !== "No disponible" && (
                  <div className="mt-2.5 bg-green-50 p-2.5 rounded-xl flex justify-between items-center border border-green-100">
                    <span className="text-green-800 text-[10px] font-bold uppercase tracking-wider">
                      Precio Final:
                    </span>
                    <span className="text-xl font-black text-green-600">
                      ${precioFinal}
                    </span>
                  </div>
                )}
              </div>
            );
          })}

          {/* Estado de lista vacía */}
          {productos.length > 0 && filtrados.length === 0 && (
            <div className="text-center text-gray-400 py-10 bg-white rounded-xl border p-4 text-xs font-medium">
              No hay coincidencias para tu búsqueda.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
