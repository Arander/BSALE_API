# BSALE_API
Macro excel para obtener datos de precios, stock y costos a traves de BSALE API y pegarlos en la planilla

La macro busca el SKU de cada uno de los productos, en un rango determinado, y pega los siguientes valores en columnas determinadas:

- Stock en unidades
- Último Costo (por algún motivo, si el stock es cero, esta información no es obtenible)
- Precio neto según Lista de Precio determinada.

Actualmente funcionando en Microsoft Excel 2010 Version 14.0 (64-bit)

Necesita instalar el conversor JSON-VBA en https://github.com/VBA-tools/VBA-JSON
