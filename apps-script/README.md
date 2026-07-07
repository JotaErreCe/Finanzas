# Apps Script — Sistema de Finanzas Familiares v3

## Estructura

- **`proyecto/`** — espejo del proyecto real de Apps Script (gestionado con
  clasp; `clasp pull` / `clasp push` desde esa carpeta):
  - `Código.js` — v2 original: correos de consumo de tarjeta, alertas,
    resumen semanal. **No se modifica.**
  - `FinanzasV3.js` — extensiones v3 en un solo archivo: transferencias ACH
    (excluye cuentas propias), hoja "Métodos de Pago", fijos del Control
    Maestro ("Fijos" + "Control Pagos"), API web para la app y `setupV3()`.
  - `appsscript.json` — manifest (zona horaria GT + config de Web App).
- `Code.gs` / `FinanzasV3.gs` — copias de referencia legibles en la raíz
  (misma fuente que `proyecto/`).

## Activación (ya casi todo hecho)

1. ~~Subir el código al proyecto~~ ✓ (con clasp)
2. ~~Desplegar como Aplicación web~~ ✓ (con clasp)
3. **Ejecutar `setupV3` una vez desde el editor** (esto lo hace el dueño de
   la cuenta): abrir `FinanzasV3.gs`, elegir `setupV3` en el desplegable de
   funciones, Ejecutar, y autorizar los permisos. Crea:
   - Hojas "Cuentas Propias", "Transferencias Internas", "Métodos de Pago"
     y "API Config" (con el token) en el sheet de seguimiento.
   - Hojas "Fijos" y "Control Pagos" en el Control Maestro.
   - Triggers: transferencias c/15 min, métodos c/hora, checklist mensual.
4. Configurar la app (⚙ Ajustes → URL + token, o el enlace `#cfg=` que
   genera Claude leyendo la hoja "API Config").

## Después de activar

- Revisar **"Cuentas Propias"**: agregar cuentas BI/Banrural/de la esposa.
  Las transferencias hacia esas cuentas NO cuentan como gasto.
- Revisar **"Fijos"** en el Control Maestro (montos, activos, día de pago).
- Las hojas viejas del Control Maestro quedan intactas.

## Funciones útiles

- `probarParserTransferencia()` — valida el parser con los correos reales.
- `mostrarToken()` — reimprime el token en los logs.
- `generarPagosDelMes()` — regenera el checklist del mes a mano.
