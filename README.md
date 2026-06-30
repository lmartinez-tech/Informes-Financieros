# Informes financieros

Aplicacion web en Streamlit para preparar dos informes iniciales:

- **Comparativo mensual**: basado en el flujo existente de TLG/mensualizados. Permite cargar saldo inicial, balances mensuales, revisar indicadores, analizar terceros y descargar Excel/PDF.
- **Estados financieros bajo NIIF**: genera un borrador con estado de situacion financiera, resultado integral, flujo de efectivo preliminar, cambios en patrimonio, notas y control de preparacion.

## Instalacion

```bash
pip install -r requirements.txt
```

## Ejecucion

```bash
streamlit run app.py
```

## Archivos de entrada

Los dos modulos trabajan con balances de prueba en Excel (`.xlsx`) que contengan, como minimo:

- Codigo de cuenta.
- Nombre de cuenta.
- Saldo final.

Cuando el archivo incluye saldo inicial, movimientos debito/credito, tercero, identificacion y sucursal, la aplicacion conserva ese detalle para validaciones y analisis.

## Alcance NIIF

El modulo NIIF deja preparada la estructura requerida para revision profesional: estados principales, notas y checklist de control. Antes de emitir estados financieros definitivos se deben completar politicas contables, revelaciones, validaciones con auxiliares, firmas y autorizaciones.
