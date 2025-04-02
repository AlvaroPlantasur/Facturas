import os
import psycopg2
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys

def main():
    # 1. Obtener credenciales y la ruta del archivo base
    db_name = os.environ.get('DB_NAME')
    db_user = os.environ.get('DB_USER')
    db_password = os.environ.get('DB_PASSWORD')
    db_host = os.environ.get('DB_HOST')
    db_port = os.environ.get('DB_PORT')
    # Se espera que EXCEL_FILE_PATH se configure en GitHub Secrets; de lo contrario se usa el valor por defecto.
    file_path = os.environ.get('EXCEL_FILE_PATH')
    
    db_params = {
        'dbname': db_name,
        'user': db_user,
        'password': db_password,
        'host': db_host,
        'port': db_port
    }
    
    # 2. Calcular el rango de fechas:
    # Desde el primer día del mes de hace dos meses hasta el día actual.
    end_date = datetime.now()
    start_date = (end_date - relativedelta(months=2)).replace(day=1)
    end_date_str = end_date.strftime('%Y-%m-%d')
    start_date_str = start_date.strftime('%Y-%m-%d')
    
    # 3. Consulta SQL (utilizando DISTINCT ON (sm.id))
    query = f"""
    SELECT DISTINCT ON (sm.id)
        sm.invoice_id as "ID FACTURA",
        rp.id as "ID CLIENTE",
        'S-' || rp.id AS "ID BBSeeds",
        sp.name as "DESCRIPCIÓN",
        sp.internal_number as "CÓDIGO FACTURA",
        sp.number as "NÚMERO DEL ASIENTO",
        to_char(sp.date_invoice, 'DD/MM/YYYY') as "FECHA FACTURA",
        sp.origin as "DOCUMENTO ORIGEN",
        pp.default_code as "REFERENCIA PRODUCTO", 
        pp.name as "NOMBRE", 
        COALESCE(pm.name,'') as "MARCA",
        s.name AS "SECCION", 
        f.name as "FAMILIA", 
        sf.name as "SUBFAMILIA",
        rc.name as "COMPAÑÍA",
        ssp.name as "SEDE",
        stp.date_preparado_app as "FECHA PREPARADO APP",
        (CASE WHEN stp.directo_cliente = true THEN 'Sí' ELSE 'No' END) AS "CAMIÓN DIRECTO",
        stp.number_of_packages as "NUMERO DE BULTOS",
        stp.num_pales as "NUMERO DE PALES",
        sp.portes as "PORTES",
        sp.portes_cubiertos as "PORTES CUBIERTOS",
        rp.nombre_comercial as "CLIENTE",
        rp.vat as "CIF CLIENTE",
        (CASE WHEN rpa.prov_id IS NOT NULL THEN (SELECT name FROM res_country_provincia WHERE id = rpa.prov_id) 
              ELSE rpa.state_id_2 END) as "PROVINCIA",
        rpa.city as "CIUDAD",
        (CASE WHEN rpa.cautonoma_id IS NOT NULL THEN (SELECT upper(name) FROM res_country_ca WHERE id = rpa.cautonoma_id) 
              ELSE '' END) as "COMUNIDAD",
        c.name as "PAÍS",
        to_char(sp.date_invoice, 'MM') as "MES",
        to_char(sp.date_invoice, 'DD') as "DÍA",
        spp.name as "PREPARADOR",
        sm.peso_arancel as "PESO",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.cantidad_pedida
                 WHEN sp.type = 'out_refund' THEN -sm.cantidad_pedida END) as "UNIDADES VENTA",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.price_subtotal
                 WHEN sp.type = 'out_refund' THEN -sm.price_subtotal END) as "BASE VENTA TOTAL",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.margen
                 WHEN sp.type = 'out_refund' THEN -sm.margen END) as "MARGEN EUROS",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.cantidad_pedida * sm.cost_price_real
                 WHEN sp.type = 'out_refund' THEN -sm.cantidad_pedida * sm.cost_price_real END) as "COSTE VENTA TOTAL"
    FROM account_invoice_line sm
    INNER JOIN account_invoice sp ON sp.id = sm.invoice_id
    INNER JOIN product_product pp ON sm.product_id = pp.id
    INNER JOIN res_partner_address rpa ON sp.address_invoice_id = rpa.id
    INNER JOIN res_country c ON c.id = rpa.pais_id
    LEFT OUTER JOIN stock_picking_invoice_rel spir ON spir.invoice_id = sp.id
    LEFT OUTER JOIN stock_picking stp ON stp.id = spir.pick_id
    LEFT OUTER JOIN stock_picking_preparador spp ON spp.id = stp.preparador
    LEFT OUTER JOIN res_company rc ON rc.id = sp.company_id
    LEFT OUTER JOIN res_partner rp ON rp.id = sp.partner_id
    LEFT OUTER JOIN stock_sede_ps ssp ON ssp.id = sp.sede_id
    LEFT OUTER JOIN product_category s ON s.id = pp.seccion
    LEFT OUTER JOIN product_category f ON f.id = pp.familia
    LEFT OUTER JOIN product_category sf ON sf.id = pp.subfamilia
    LEFT OUTER JOIN product_marca pm ON pm.id = pp.marca
    WHERE sp.type in ('out_invoice','out_refund')
      AND sp.state in ('open','paid')
      AND sp.journal_id != 11
      AND sp.anticipo = false
      AND pp.default_code NOT LIKE 'XXX%'
      AND sp.obsolescencia = false
      AND rp.nombre_comercial NOT LIKE '%PLANTASUR TRADING%'
      AND rp.nombre_comercial NOT LIKE '%PLANTADUCH%'
      AND sp.date_invoice >= '{start_date_str}' 
      AND sp.date_invoice <= '{end_date_str}'
    GROUP BY 
        sm.id,
        rp.id,
        sm.company_id,
        sp.sede_id,
        sp.date_invoice,
        pp.seccion,
        pp.familia,
        pp.subfamilia,
        pp.default_code,
        pp.id,
        sm.cantidad_pedida,
        sp.partner_id,
        rpa.prov_id,
        rpa.state_id_2,
        rpa.cautonoma_id,
        c.name,
        sp.address_invoice_id,
        pp.proveedor_id1,
        sm.price_subtotal,
        sm.margen,
        pp.tarifa5,
        sp.directo_cliente,
        sp.obsolescencia,
        sp.anticipo,
        sp.name,
        s.name,
        f.name,
        sf.name,
        rc.name,
        ssp.name,
        stp.date_preparado_app,
        stp.directo_cliente,
        stp.number_of_packages,
        stp.num_pales,
        sp.portes,
        sp.portes_cubiertos,
        rp.nombre_comercial,
        spp.name,
        sp.internal_number,
        sp.origin,
        sp.number,
        rp.vat,
        pm.name,
        rpa.city
    ORDER BY 
        sm.id, stp.number_of_packages DESC;
    """
    
    # 4. Ejecutar la consulta y obtener los datos
    try:
        with psycopg2.connect(**db_params) as conn:
            with conn.cursor() as cur:
                cur.execute(query)
                resultados = cur.fetchall()
                headers = [desc[0] for desc in cur.description]
    except Exception as e:
        print(f"Error al conectar o ejecutar la consulta: {e}")
        sys.exit(1)
    
    if not resultados:
        print("No se obtuvieron resultados de la consulta.")
        return
    else:
        print(f"Se obtuvieron {len(resultados)} filas de la consulta.")
    
    # 5. Abrir el archivo base "Master_Facturas_Desglosadas_2025.xlsx" (ubicado en la raíz del repositorio)
    try:
        book = load_workbook(file_path)
        sheet = book.active
    except FileNotFoundError:
        print(f"No se encontró el archivo base '{file_path}'. Se aborta para no perder el formato.")
        return
    
    # 6. Evitar duplicados (asumiendo que la primera columna es "ID FACTURA")
    existing_ids = {row[0] for row in sheet.iter_rows(min_row=2, values_only=True)}
    for row in resultados:
        if row[0] not in existing_ids:
            sheet.append(row)
    
    # 7. Actualizar la referencia de la tabla existente (la tabla se llama "Lineas2025")
    if "Lineas2025" in sheet.tables:
        tabla = sheet.tables["Lineas2025"]
        max_row = sheet.max_row
        max_col = sheet.max_column
        last_col_letter = get_column_letter(max_col)
        new_ref = f"A1:{last_col_letter}{max_row}"
        tabla.ref = new_ref
        print(f"Tabla 'Lineas2025' actualizada a rango: {new_ref}")
    else:
        print("No se encontró la tabla 'Lineas2025'. Se conservará el formato actual, pero no se actualizará la referencia de la tabla.")
    
    book.save(file_path)
    print(f"Archivo guardado con los datos actualizados en '{file_path}'.")

if __name__ == '__main__':
    main()
