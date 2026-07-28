# Cotizador FMI — Panorama Outsourcing

App de escritorio para convertir cotizaciones de proveedores en tu cotización FMI
(formato PAN-FO-FM-02), con Sustento de Gasto, comparativo de proveedores y bitácora.

## ¿Qué hace?
1. Cargas la cotización del proveedor (PDF con texto se extrae automático; si es
   escaneada/foto, la app usa **OCR** para leerla igual — sin instalar nada extra).
2. La app aplica el **margen** configurable (por defecto 15% "sobre venta" = costo ÷ 0.85).
3. Agregas líneas estándar (SCTR, EPPS, traslado) con un clic.
4. Genera con un botón:
   - **Cotización FMI** en Excel y PDF (tu plantilla PAN-FO-FM-02).
   - **Sustento de Gasto** (Excel y PDF) con utilidad.
   - **Comparativo** de varios proveedores (opcional).
   - **Bitácora** (registro acumulado de todas tus cotizaciones).
5. El **N° correlativo** avanza solo (FMI - 0246-26, 0247-26, ...).

## Dónde está la app instalada
**`Documentos\CotizadorFMI\`** — doble clic en **`CotizadorFMI.exe`**
(o el acceso directo **"Cotizador FMI"** del Escritorio).

La carpeta contiene:
- `CotizadorFMI.exe` — la app.
- `_internal\` — sus componentes (no borrar).
- `salidas\` — las cotizaciones que generas (una subcarpeta por número).
- `datos\` — tu configuración (margen, correlativo, firma) y la bitácora.

> **No borres `_internal` ni muevas el .exe solo.** Si querés cambiar la app de
> lugar o llevarla a otra PC, copiá la **carpeta completa**.

## Llevarla a otra PC
1. Copiá **toda la carpeta** `CotizadorFMI` (por USB o red).
2. Pegala en la otra PC, en Documentos o Escritorio.
3. Doble clic en `CotizadorFMI.exe`.

> Si el antivirus la bloquea (falso positivo): restaurá `CotizadorFMI.exe` desde
> el antivirus (Seguridad de Windows → Historial de protección → Restaurar) o
> agregá la carpeta como exclusión. El formato carpeta reduce mucho este problema.

## Cómo usarla (sin instalar nada, si tienes Python)
Doble clic en **`Iniciar_Cotizador.bat`**.

## Cómo crear el ejecutable (.exe)
1. Instala PyInstaller una sola vez:  `pip install pyinstaller`
2. Doble clic en **`build.bat`**.
3. El ejecutable queda en `dist\CotizadorFMI.exe`.
   Mantén la carpeta **`datos`** junto al `.exe` (guarda tu configuración y bitácora).

## Flujo en la ventana
1. **Datos generales**: cliente, fecha (auto), asunto, N° ticket, sede.
2. **Proveedor principal**: nombre, RUC, N° de su cotización. Botón *Cargar PDF*.
3. **Líneas**: revisa/edita descripción, U.M., cantidad y **costo del proveedor**.
   La columna *Venta* muestra el precio con margen en vivo.
   - Marca **Gasto** en las líneas del proveedor (cuentan para la utilidad).
   - Desmarca *Gasto* en líneas propias (SCTR/EPPS) que no son costo de proveedor.
4. **Margen**: cambia el % o el modo (÷ sobre venta / × sobre costo).
5. **Generar**: elige documentos y presiona **GENERAR DOCUMENTOS**.

## Comparativo de proveedores
Para comparar varios: llena las líneas del proveedor A → *"＋ Proveedor al comparativo"*.
Cambia el nombre/precios al proveedor B → agrégalo también. Marca *Comparativo* y genera.
El Excel resalta en verde el más económico por ítem y el total menor.

## Logo y firma (🖼)
Botón **"🖼 Logo y firma"** arriba a la derecha. Tiene dos opciones independientes:

- **Logo de empresa** — aparece en la cabecera de la Cotización y del Sustento.
- **Logo de firma** — la firma escaneada que va al pie, sobre el nombre.

De cada una podés *Cargar imagen* (PNG/JPG) o *Usar la original* para volver a la de
las plantillas. Muestra una vista previa y la imagen se ajusta sola al espacio,
manteniendo su proporción. Mientras no cargues nada, se usan las originales.

## Configuración (⚙)
Cambia margen por defecto, IGV, prefijo/año/próximo número del correlativo, moneda,
razón social y datos de firma. Se guardan en `datos/config.json`.

## Dónde quedan los archivos
En `salidas/<N° cotización>/`. La bitácora en `datos/Bitacora_Cotizaciones_FMI.xlsx`.

## Reglas de cálculo (validadas con tu caso real 0245-26)
- Venta unitaria = costo ÷ (1 − margen)  →  755 ÷ 0.85 = **888.24**
- IGV 18% sobre el costo directo.
- Utilidad = Venta (sin IGV) − Gasto del proveedor (sin IGV).
