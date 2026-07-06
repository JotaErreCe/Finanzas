# Apps Script — Sistema de Finanzas Familiares v3

## Qué hay aquí

| Archivo | Estado | Qué hace |
|---|---|---|
| `Code.gs` | **Ya está en tu proyecto** (copia de referencia, no hay que pegarlo) | v2: procesa correos de consumo de tarjeta, alertas, resumen semanal |
| `Transferencias.gs` | **NUEVO — pegar** | Procesa correos de transferencia ACH; separa internas de gastos reales |
| `MetodosPago.gs` | **NUEVO — pegar** | Hoja "Métodos de Pago": gasto por tarjeta/cuenta, últimos 6 períodos |
| `ControlMaestro.gs` | **NUEVO — pegar** | Catálogo "Fijos" + checklist mensual "Control Pagos" en el Control Maestro |
| `API.gs` | **NUEVO — pegar** | Web App: la app de GitHub lee el dashboard y escribe gastos/pagos |
| `Setup.gs` | **NUEVO — pegar** | `setupV3()`: crea hojas, token y triggers de un solo clic |

**El código v2 no se toca.** Todo lo nuevo son archivos independientes que
conviven en el mismo proyecto (Apps Script comparte funciones entre archivos).
`procesarCorreos` y sus triggers siguen igual.

## Instalación (una sola vez, ~5 minutos)

1. Abrir el sheet de seguimiento → **Extensiones → Apps Script**.
2. Por cada archivo NUEVO: **＋ → Secuencia de comandos**, nombrarlo igual
   (ej. `Transferencias`) y pegar el contenido. Guardar.
3. Seleccionar la función **`setupV3`** en la barra superior → **Ejecutar**.
   - Autorizar los permisos nuevos (acceso al Control Maestro).
   - En los logs aparece el **TOKEN de la API** — copiarlo.
4. **Implementar → Nueva implementación → Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
   - Copiar la **URL** que termina en `/exec`.
5. Abrir la app → ⚙ Ajustes → pegar URL y token. Listo (tu esposa hace lo
   mismo una vez en su teléfono, o abre el enlace de configuración que
   puedes compartirle desde Ajustes).

## Después de instalar

- **Revisar la hoja "Cuentas Propias"** en el sheet de seguimiento y agregar
  los números de cuenta que faltan (BI, Banrural, cuentas de tu esposa).
  Toda transferencia hacia esas cuentas se registra en "Transferencias
  Internas" y **no cuenta como gasto**.
- **Revisar la hoja "Fijos"** en el Control Maestro: montos, cuáles siguen
  activos, y opcionalmente el día de pago.
- Las hojas viejas del Control Maestro (Pagos 2026, Control mensual…) quedan
  intactas; cuando confíes en el flujo nuevo puedes archivarlas.

## Funciones de prueba

- `probarParserTransferencia()` — valida el parser con los correos reales.
- `probarSistema()` / `probarParserGyT()` — las de siempre de v2.
- `mostrarToken()` — reimprime el token si lo perdiste.
