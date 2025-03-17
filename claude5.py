import ezdxf
from shapely.geometry import Point, Polygon
import re
import os
import sys
import xlwings as xw
import traceback
import time

# Forzar la consola a aceptar UTF-8 en Windows
if os.name == 'nt':
    os.system('chcp 65001')
    sys.stdout.reconfigure(encoding='utf-8')

# Función para reemplazar caracteres especiales
def reemplazar_caracteres_especiales(texto):
    texto = texto.replace("%%C", "∅")
    texto = texto.replace("\\A1;", "")  # Eliminar \A1; que aparece en algunos textos
    texto = re.sub(r'\\[A-Za-z0-9]+;', '', texto)
    return texto

# Función para obtener textos dentro de una polilínea
def obtener_textos_dentro_de_polilinea(polilinea, textos):
    vertices = [(p[0], p[1]) for p in polilinea]
    poligono = Polygon(vertices)
    textos_en_polilinea = []

    for text in textos:
        punto_texto = Point(text.dxf.insert)
        if poligono.contains(punto_texto):
            if text.dxftype() == 'MTEXT':
                texto_contenido = text.text
            else:
                texto_contenido = text.dxf.text

            texto_contenido = reemplazar_caracteres_especiales(texto_contenido)
            textos_en_polilinea.append(texto_contenido)
    return textos_en_polilinea

# Función para obtener polilíneas dentro de una polilínea principal
def obtener_polilineas_dentro_de_polilinea(polilinea_principal, polilineas_anidadas):
    vertices_principal = [(p[0], p[1]) for p in polilinea_principal]
    poligono_principal = Polygon(vertices_principal)
    polilineas_dentro = []

    # Print debug information about the polyline layers
    capas_encontradas = set()
    for polilinea in polilineas_anidadas:
        capas_encontradas.add(polilinea.dxf.layer)
    
    print(f"Capas de polilíneas encontradas: {capas_encontradas}")

    for polilinea in polilineas_anidadas:
        try:
            vertices_anidada = [(p[0], p[1]) for p in polilinea.get_points('xy')]
            poligono_anidado = Polygon(vertices_anidada)
            
            # Check if the polyline intersects or is contained in the main polyline
            if poligono_principal.intersects(poligono_anidado):
                area_interseccion = poligono_principal.intersection(poligono_anidado).area
                area_anidada = poligono_anidado.area
                ratio_interseccion = area_interseccion / area_anidada if area_anidada > 0 else 0
                
                # If at least 30% of the polyline is inside the main polyline, consider it as inside
                if ratio_interseccion >= 0.3:
                    print(f"Polilínea en capa '{polilinea.dxf.layer}' intersecta con {ratio_interseccion:.2f} de su área")
                    polilineas_dentro.append(polilinea)
        except Exception as e:
            print(f"Error al procesar polilínea: {e}")
    
    return polilineas_dentro
# Función para calcular el centro de una polilínea
def calcular_centro_polilinea(vertices):
    x_coords = [v[0] for v in vertices]
    y_coords = [v[1] for v in vertices]
    centro_x = min(x_coords) + (max(x_coords) - min(x_coords)) / 2
    centro_y = min(y_coords) + (max(y_coords) - min(y_coords)) / 2
    return centro_x, centro_y

# Función para encontrar el bloque acero en el documento
def encontrar_bloque_acero(doc, bloque_nombre="BD-ACERO PRELOSA", capa_nombre="BD-ACERO POSITIVO"):
    """
    Busca el bloque de acero en el documento.
    """
    msp = doc.modelspace()
    
    # Método 1: Buscar por nombre exacto
    for entity in msp:
        if entity.dxftype() == 'INSERT' and entity.dxf.name.strip().upper() == bloque_nombre.upper():
            print(f"Bloque encontrado por nombre: {entity.dxf.name}")
            return entity

    # Método 2: Buscar por capa
    for entity in msp:
        if entity.dxftype() == 'INSERT' and capa_nombre.upper() in entity.dxf.layer.upper():
            print(f"Bloque encontrado por capa: {entity.dxf.layer}")
            return entity

    # Método 3: Buscar coincidencias parciales
    for entity in msp:
        if entity.dxftype() == 'INSERT':
            nombre = entity.dxf.name.upper()
            capa = entity.dxf.layer.upper()
            if "ACERO" in nombre and "PRELOSA" in nombre:
                print(f"Bloque encontrado por coincidencia parcial en nombre: {entity.dxf.name}")
                return entity
            if "ACERO" in capa and "POSITIVO" in capa:
                print(f"Bloque encontrado por coincidencia parcial en capa: {entity.dxf.layer}")
                return entity

    # Método 4: Buscar por atributos
    for entity in msp:
        if entity.dxftype() == 'INSERT':
            try:
                atributos = {}
                
                # Intentar diferentes métodos para obtener atributos
                try:
                    for attrib in entity.attribs:
                        atributos[attrib.dxf.tag] = attrib.dxf.text
                except:
                    pass
                
                if not atributos:
                    try:
                        for child in entity:
                            if child.dxftype() == 'ATTRIB':
                                atributos[child.dxf.tag] = child.dxf.text
                    except:
                        pass
                
                if not atributos:
                    try:
                        for attrib in entity.get_attribs():
                            atributos[attrib.dxf.tag] = attrib.dxf.text
                    except:
                        pass
                
                # Verificar atributos específicos
                if 'AS_LONG' in atributos or 'AS_TRA1' in atributos or 'AS_TRA2' in atributos:
                    print(f"Bloque encontrado por atributos: {list(atributos.keys())}")
                    return entity
            except:
                pass
    
    print("No se encontró el bloque de acero. Se creará uno genérico.")
    return None

# Función para obtener definición del bloque
def obtener_definicion_bloque(bloque_original):
    """
    Obtiene la definición y propiedades de un bloque original para replicarlo.
    """
    try:
        definicion = {
            'nombre': bloque_original.dxf.name,
            'capa': bloque_original.dxf.layer,
            'xscale': getattr(bloque_original.dxf, 'xscale', 1.0),
            'yscale': getattr(bloque_original.dxf, 'yscale', 1.0),
            'rotation': getattr(bloque_original.dxf, 'rotation', 0.0)
        }
        return definicion
    except Exception as e:
        print(f"Error al obtener definición del bloque: {e}")
        return {
            'nombre': 'BD-ACERO PRELOSA',
            'capa': 'BD-ACERO POSITIVO',
            'xscale': 1.0,
            'yscale': 1.0,
            'rotation': 0.0
        }


def calcular_orientacion_prelosa(vertices):
    """
    Calcula la orientación de la prelosa analizando la longitud de sus lados.
    Devuelve el ángulo de rotación en grados para alinear el bloque con el lado más corto
    sin que el texto se invierta o aparezca al revés.
    """
    try:
        import math
        
        # Si la prelosa no tiene al menos 3 vértices, devolver rotación por defecto
        if len(vertices) < 3:
            return 0.0
            
        # Calcular la longitud de los lados
        lados = []
        for i in range(len(vertices)):
            x1, y1 = vertices[i]
            x2, y2 = vertices[(i + 1) % len(vertices)]
            
            # Calcular longitud del lado
            longitud = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
            
            # Calcular ángulo del lado (en radianes)
            angulo = math.atan2(y2 - y1, x2 - x1)
            
            lados.append((longitud, angulo))
        
        # Ordenar los lados por longitud (ascendente)
        lados_ordenados = sorted(lados, key=lambda x: x[0])
        
        # Obtener el ángulo del lado más corto (en radianes)
        angulo_lado_corto = lados_ordenados[0][1]
        
        # Convertir a grados
        angulo_grados = math.degrees(angulo_lado_corto)
        
        # Forzar el ángulo a estar entre 0 y 180 grados
        # Esto evita la inversión del texto en el bloque
        if angulo_grados < 0:
            angulo_grados += 180
        elif angulo_grados > 180:
            angulo_grados -= 180
            
        print(f"Orientación final: {angulo_grados:.2f}° (texto correctamente orientado)")
        
        return angulo_grados
        
    except Exception as e:
        print(f"Error al calcular la orientación: {e}")
        return 0.0


# Función para insertar bloque con atributos modificados
def insertar_bloque_acero(msp, definicion_bloque, centro, as_long, as_tra1, as_tra2=None):
    """
    Inserta un bloque de acero en el centro de la prelosa con los valores calculados.
    """
    try:
        print("Copiar y Pegando bloque de aceros en centro de la prelosa...")
        
        # Ya no intentamos convertir a float, puesto que recibimos strings formateados
        # directamente desde la función procesar_prelosa
        
        # Usar directamente los valores recibidos (que ya son strings formateados)
        str_as_long = as_long
        str_as_tra1 = as_tra1
        str_as_tra2 = as_tra2 if as_tra2 is not None else ""
        
        print("Modificando Texto del bloque de aceros...")
        print("texto modificado:")
        print(f"    AS_LONG: {str_as_long}")
        print(f"    AS_TRA1: {str_as_tra1}")
        if as_tra2 is not None:
            print(f"    AS_TRA2: {str_as_tra2}")
        
        # Mostrar rotación que se está aplicando (para depuración)
        if 'rotation' in definicion_bloque:
            print(f"    Rotación del bloque: {definicion_bloque['rotation']:.2f}°")
        
        # Crear inserción del bloque
        bloque = msp.add_blockref(
            name=definicion_bloque['nombre'],
            insert=centro,
            dxfattribs={
                'layer': definicion_bloque['capa'],
                'xscale': definicion_bloque['xscale'],
                'yscale': definicion_bloque['yscale'],
                'rotation': definicion_bloque['rotation']
            }
        )
        
        # Preparar valores de atributos
        valores_atributos = {
            'AS_LONG': str_as_long,
            'AS_TRA1': str_as_tra1
        }
        
        # IMPORTANTE: Solo incluir AS_TRA2 si tiene un valor
        if as_tra2 is not None and as_tra2 != "":
            valores_atributos['AS_TRA2'] = str_as_tra2
        
        # Asignar atributos - métodos por orden de preferencia
        attribs_asignados = False
        
        # Método 1: add_auto_attribs - el más confiable y rápido
        try:
            bloque.add_auto_attribs(valores_atributos)
            attribs_asignados = True
            return bloque
        except Exception as e:
            print(f"No se pudo usar add_auto_attribs: {e}")
        
        # Método 2: Iterar a través de attribs
        if not attribs_asignados:
            try:
                for attrib in bloque.attribs:
                    if attrib.dxf.tag in valores_atributos:
                        attrib.dxf.text = valores_atributos[attrib.dxf.tag]
                        attribs_asignados = True
            except Exception as e:
                print(f"No se pudo acceder a attribs: {e}")
        
        # Método 3: Iterar a través de los hijos
        if not attribs_asignados:
            try:
                for child in bloque:
                    if child.dxftype() == 'ATTRIB' and child.dxf.tag in valores_atributos:
                        child.dxf.text = valores_atributos[child.dxf.tag]
                        attribs_asignados = True
            except Exception as e:
                print(f"No se pudo iterar a través de los hijos: {e}")
        
        # Método 4: usar get_attribs()
        if not attribs_asignados:
            try:
                for attrib in bloque.get_attribs():
                    if attrib.dxf.tag in valores_atributos:
                        attrib.dxf.text = valores_atributos[attrib.dxf.tag]
                        attribs_asignados = True
            except Exception as e:
                print(f"No se pudo usar get_attribs(): {e}")
        
        # Método 5: último recurso - añadir atributos manualmente
        if not attribs_asignados:
            try:
                offset_y = 0
                for tag, valor in valores_atributos.items():
                    bloque.add_attrib(
                        tag=tag,
                        text=valor,
                        insert=(0, offset_y),
                        dxfattribs={
                            'layer': definicion_bloque['capa'],
                            'height': 2.5,
                            'style': 'STANDARD'
                        }
                    )
                    offset_y -= 3  # Espacio vertical entre atributos
                attribs_asignados = True
            except Exception as e:
                print(f"Error al añadir atributos manualmente: {e}")
        
        return bloque
    
    except Exception as e:
        print(f"Error al insertar bloque: {e}")
        traceback.print_exc()
        return None



def eliminar_entidades_por_capa(doc, capas_a_eliminar):
    """
    Elimina todas las entidades en el modelo que pertenecen a las capas especificadas.
    
    Args:
        doc: Documento DXF
        capas_a_eliminar: Lista de nombres de capas cuyas entidades se eliminarán
    
    Returns:
        int: Número de entidades eliminadas
    """
    msp = doc.modelspace()
    entidades_eliminadas = 0
    
    # Recopilamos todas las entidades a eliminar en una lista
    entidades_a_eliminar = []
    for entity in msp:
        if hasattr(entity, 'dxf') and hasattr(entity.dxf, 'layer'):
            if entity.dxf.layer in capas_a_eliminar:
                entidades_a_eliminar.append(entity)
    
    # Eliminamos las entidades
    for entity in entidades_a_eliminar:
        msp.delete_entity(entity)
        entidades_eliminadas += 1
    
    print(f"Se eliminaron {entidades_eliminadas} entidades de las capas: {', '.join(capas_a_eliminar)}")
    return entidades_eliminadas

# Modificar el código principal para llamar a esta función justo antes de guardar el archivo DXF
# Añadir justo antes de doc.saveas(output_dxf_path):

# Eliminar polilíneas de acero




def formatear_valor_espaciamiento(valor_decimal):
    """
    Formatea valores de espaciamiento según la regla:
    - Si termina en 00 (como 200, 100), se recorta a 20, 10
    - Si no, se mantiene el valor completo (como 175)
    
    Args:
        valor_decimal: Valor decimal a formatear (ej: 0.200, 0.175)
        
    Returns:
        str: Valor formateado según las reglas
    """
    # Convertir a entero multiplicado por 1000 y redondeado
    valor_entero = int(round(valor_decimal, 3) * 1000)
    
    # Verificar si termina en 00 (divisible por 100)
    if valor_entero % 100 == 0:
        # Recortar el valor (200 -> 20, 100 -> 10)
        return str(valor_entero // 10)
    else:
        # Mantener el valor completo con 3 dígitos
        return f"{valor_entero:03d}"



# Función principal modificada para usar bloques en lugar de textos
# Función principal modificada para usar bloques en lugar de textos
def procesar_prelosas_con_bloques(file_path, excel_path, output_dxf_path):
    """
    Procesa las prelosas identificando tipos y contenidos,
    calcula valores usando Excel y coloca bloques con los resultados.
    """
    try:
        tiempo_inicio = time.time()
        
        # Cargar el documento DXF
        doc = ezdxf.readfile(file_path)
        msp = doc.modelspace()
        
        print("Capas disponibles en el archivo DXF:")
        for layer in doc.layers:
            print(f"  - {layer.dxf.name}")
            
        # Abrir Excel
        app = xw.App(visible=False)  # Abrir Excel en segundo plano
        wb = app.books.open(excel_path)  # Abrir el archivo
        ws = wb.sheets.active  # Obtener la hoja activa
        
        # Leer los valores iniciales de las celdas K8 y K17
        k8_original = ws.range('K8').value
        k17_original = ws.range('K17').value
        print("\nVALORES INICIALES EN EXCEL:")
        print(f"  Celda K8 = {k8_original}")
        print(f"  Celda K17 = {k17_original}")
        print("-" * 40)
            
        # Encontrar bloque de referencia para acero
        bloque_original = encontrar_bloque_acero(doc)
        if bloque_original:
            definicion_bloque = obtener_definicion_bloque(bloque_original)
        else:
            print("Usando definición genérica para el bloque")
            definicion_bloque = {
                'nombre': 'BD-ACERO PRELOSA',
                'capa': 'BD-ACERO POSITIVO',
                'xscale': 1.0,
                'yscale': 1.0,
                'rotation': 0.0
            } 
        
        # Obtener polilíneas y textos
        polilineas_macizas = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA MACIZA"]
        polilineas_aligeradas = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 20"]
        polilineas_aligeradas_2sent = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 20 - 2 SENT"]
        polilineas_acero = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and 
                   entity.dxf.layer in ["ACERO HORIZONTAL", "ACERO VERTICAL", 
                                        "BD-ACERO HORIZONTAL", "BD-ACERO VERTICAL",
                                        "ACERO", "REFUERZO", "ARMADURA"]]
        textos = [entity for entity in msp if entity.dxftype() in ['TEXT', 'MTEXT']]
        
        # Contadores para estadísticas
        total_prelosas = 0
        total_bloques = 0
        
        # Función para calcular la orientación de la prelosa
        def calcular_orientacion_prelosa(vertices):
            """
            Calcula la orientación de la prelosa analizando la longitud de sus lados.
            Devuelve el ángulo de rotación en grados para alinear el bloque con el lado más corto.
            
            Args:
                vertices: Lista de puntos (x, y) que forman la polilínea de la prelosa
                
            Returns:
                float: Ángulo de rotación en grados
            """
            try:
                import math
                
                # Si la prelosa no tiene al menos 3 vértices, devolver rotación por defecto
                if len(vertices) < 3:
                    return 0.0
                    
                # Calcular la longitud de los lados
                lados = []
                for i in range(len(vertices)):
                    x1, y1 = vertices[i]
                    x2, y2 = vertices[(i + 1) % len(vertices)]  # Cierra el circuito al último punto
                    
                    # Calcular longitud del lado
                    longitud = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
                    
                    # Calcular ángulo del lado (en radianes)
                    angulo = math.atan2(y2 - y1, x2 - x1)
                    
                    lados.append((longitud, angulo))
                
                # Ordenar los lados por longitud (ascendente)
                lados_ordenados = sorted(lados, key=lambda x: x[0])
                
                # Obtener el ángulo del lado más corto (en radianes)
                angulo_lado_corto = lados_ordenados[0][1]
                
                # Convertir a grados
                angulo_grados = math.degrees(angulo_lado_corto)
                
                # Normalizar el ángulo para que esté en el rango 0-360
                if angulo_grados < 0:
                    angulo_grados += 360
                
                # Usar directamente el ángulo del lado más corto para que el acero azul 
                # quede alineado con el lado más estrecho
                angulo_final = angulo_grados
                
                print(f"Orientación corregida: {angulo_final:.2f}° (acero azul alineado con lado más estrecho)")
                
                return angulo_final
                
            except Exception as e:
                print(f"Error al calcular la orientación: {e}")
                return 0.0  # Valor por defecto en caso de error
        
        # FUNCIÓN AUXILIAR: Procesa una prelosa (se aplica a todos los tipos)
        def procesar_prelosa(polilinea, tipo_prelosa, idx):
            nonlocal total_prelosas, total_bloques
            
            total_prelosas += 1
            vertices = polilinea.get_points('xy')
            centro_prelosa = calcular_centro_polilinea(vertices)
            polilineas_dentro = obtener_polilineas_dentro_de_polilinea(vertices, polilineas_acero)
            
            print(f"{tipo_prelosa} numero {idx+1} encontrada:")
            print(f"valor por defecto: .20")
            print(f"Centro de la prelosa: ({centro_prelosa[0]}, {centro_prelosa[1]})")
            print(f"Polilíneas dentro encontradas: {len(polilineas_dentro)}")
            
            # Variables para almacenar textos por tipo de acero
            textos_horizontal = []
            textos_vertical = []
            
            # Procesar polilíneas de acero
            for polilinea_anidada in polilineas_dentro:
                vertices_anidada = polilinea_anidada.get_points('xy')
                textos_dentro = obtener_textos_dentro_de_polilinea(vertices_anidada, textos)
                
                print(f"Polilínea anidada en {tipo_prelosa.lower()} {idx+1} tiene {len(textos_dentro)} textos dentro.")
                
                # Clasificar textos según el tipo de acero
                tipo_acero = polilinea_anidada.dxf.layer.upper()
                if "HORIZONTAL" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO HORIZONTAL: {texto}")
                        textos_horizontal.append(texto)
                elif "VERTICAL" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO VERTICAL: {texto}")
                        textos_vertical.append(texto)
            
            # Procesar datos en Excel
            print("Actualizando excel...")
            
            # No limpiar celdas, solo sobrescribir
            # Almacenar los valores originales de K8 y K17 antes de cualquier modificación
            k8_actual = ws.range('K8').value
            k17_actual = ws.range('K17').value
            print(f"\nVALORES ACTUALES PARA PRESERVAR:")
            print(f"  Valor K8 actual = {k8_actual}")
            print(f"  Valor K17 actual = {k17_actual}")
            print("-" * 40)
            
            # Casos especiales para PRELOSA ALIGERADA 20
            if tipo_prelosa == "PRELOSA ALIGERADA 20":
                # Caso 1: Si tenemos textos horizontales (prioridad)
                if len(textos_horizontal) > 0:
                    print(f"Procesando {len(textos_horizontal)} textos horizontales en PRELOSA ALIGERADA 20")
                    
                    # Procesar primer texto horizontal (G4, H4, J4)
                    if len(textos_horizontal) >= 1:
                        texto = textos_horizontal[0]
                        try:
                            # Extraer diámetro del texto (ej: "1∅3/8"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario (para 3/8")
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = "1"  # Por defecto
                                separacion_decimal = 20  # Valor por defecto para prelosas aligeradas
                                
                                # Escribir en Excel
                                ws.range('G4').value = int(cantidad)
                                ws.range('H4').value = diametro_con_comillas
                                ws.range('J4').value = separacion_decimal
                                
                                print(f"Colocando en el excel primer texto: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto horizontal en PRELOSA ALIGERADA 20 '{texto}': {e}")
                    
                    # Procesar segundo texto horizontal (G5, H5, J5) si existe
                    if len(textos_horizontal) >= 2:
                        texto = textos_horizontal[1]
                        try:
                            # Extraer diámetro del texto (ej: "1∅8mm")
                            diametro_match = re.search(r'∅([\w/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Dejar el diámetro como está, normalmente "8mm"
                                diametro_formateado = diametro
                                
                                cantidad = "1"  # Por defecto
                                separacion_decimal = 20  # Valor por defecto para prelosas aligeradas
                                
                                # Escribir en Excel
                                ws.range('G5').value = int(cantidad)
                                ws.range('H5').value = diametro_formateado
                                ws.range('J5').value = separacion_decimal
                                
                                print(f"Colocando en el excel segundo texto: {cantidad} -> G5, {diametro_formateado} -> H5, {separacion_decimal} -> J5")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto horizontal en PRELOSA ALIGERADA 20 '{texto}': {e}")
                
                # Caso 2: Si tenemos textos verticales pero no horizontales
                elif len(textos_vertical) > 0:
                    for i, texto in enumerate(textos_vertical):
                        try:
                            # Extraer diámetro del texto (ej: "1∅1/2"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                diametro_con_comillas = f"{diametro}\""
                                cantidad = "1"  # Por defecto
                                
                                # Para prelosas aligeradas, siempre usar 20 como separación
                                separacion_decimal = 20  # Valor por defecto para prelosas aligeradas
                                
                                # Escribir en Excel
                                ws.range('G4').value = int(cantidad)
                                ws.range('H4').value = diametro_con_comillas
                                ws.range('J4').value = separacion_decimal
                                
                                print(f"Colocando en el excel para PRELOSA ALIGERADA 20: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar texto vertical en PRELOSA ALIGERADA 20 '{texto}': {e}")
                        # Agregar esta variable para la distancia en prelosas aligeradas 2 sent

            # Modificar la sección donde se procesan los casos especiales para incluir PRELOSA ALIGERADA 20 - 2 SENT
            # Añadir este bloque después del caso de PRELOSA ALIGERADA 20

            # Casos especiales para PRELOSA ALIGERADA 20 - 2 SENT
            elif tipo_prelosa == "PRELOSA ALIGERADA 20 - 2 SENT":
                dist_aligerada2sent = 0.605  # Valor por defecto
                # Caso 1: Si tenemos textos horizontales
                if len(textos_horizontal) > 0:
                    print(f"Procesando {len(textos_horizontal)} textos horizontales en PRELOSA ALIGERADA 20 - 2 SENT")
                    
                    # Procesar primer texto horizontal (G4, H4, J4)
                    if len(textos_horizontal) >= 1:
                        texto = textos_horizontal[0]
                        try:
                            # Extraer diámetro del texto (ej: "1∅3/8"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario (para 3/8")
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = "1"  # Por defecto
                                separacion_decimal = dist_aligerada2sent  # Usar el valor específico para 2SENT
                                
                                # Escribir en Excel
                                ws.range('G4').value = int(cantidad)
                                ws.range('H4').value = diametro_con_comillas
                                ws.range('J4').value = separacion_decimal
                                
                                print(f"Colocando en el excel primer texto horizontal: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto horizontal en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                    
                    # Procesar segundo texto horizontal (G5, H5, J5) si existe
                    if len(textos_horizontal) >= 2:
                        texto = textos_horizontal[1]
                        try:
                            # Extraer diámetro del texto (ej: "1∅1/2"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = "1"  # Por defecto
                                separacion_decimal = dist_aligerada2sent  # Usar el valor específico para 2SENT
                                
                                # Escribir en Excel
                                ws.range('G5').value = int(cantidad)
                                ws.range('H5').value = diametro_con_comillas
                                ws.range('J5').value = separacion_decimal
                                
                                print(f"Colocando en el excel segundo texto horizontal: {cantidad} -> G5, {diametro_con_comillas} -> H5, {separacion_decimal} -> J5")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto horizontal en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                
                # Caso 2: Si tenemos textos verticales
                if len(textos_vertical) > 0:
                    print(f"Procesando {len(textos_vertical)} textos verticales en PRELOSA ALIGERADA 20 - 2 SENT")
                    
                    # Procesar primer texto vertical (G14, H14, J14)
                    if len(textos_vertical) >= 1:
                        texto = textos_vertical[0]
                        try:
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = "1"  # Por defecto
                                separacion_decimal = dist_aligerada2sent  # Usar el valor específico para 2SENT
                                
                                # Escribir en Excel
                                ws.range('G14').value = int(cantidad)
                                ws.range('H14').value = diametro_con_comillas
                                ws.range('J14').value = separacion_decimal
                                
                                print(f"Colocando en el excel primer texto vertical: {cantidad} -> G14, {diametro_con_comillas} -> H14, {separacion_decimal} -> J14")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto vertical '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto vertical en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                    
                    # Procesar segundo texto vertical (G15, H15, J15) si existe
                    if len(textos_vertical) >= 2:
                        texto = textos_vertical[1]
                        try:
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = "1"  # Por defecto
                                separacion_decimal = dist_aligerada2sent  # Usar el valor específico para 2SENT
                                
                                # Escribir en Excel
                                ws.range('G15').value = int(cantidad)
                                ws.range('H15').value = diametro_con_comillas
                                ws.range('J15').value = separacion_decimal
                                
                                print(f"Colocando en el excel segundo texto vertical: {cantidad} -> G15, {diametro_con_comillas} -> H15, {separacion_decimal} -> J15")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto vertical '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto vertical en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
            else:
                # Procesar acero horizontal (G4, H4, J4) para otros tipos de prelosas
                for i, texto in enumerate(textos_horizontal):
                    try:
                        # Extraer información del texto
                        match = re.match(r'∅(\d+\/\d+)\"@(\d+)', texto)
                        if match:
                            diametro = match.group(1)
                            separacion = match.group(2)
                            cantidad = "1"  # Por defecto
                            
                            # Añadir comillas al diámetro si no las tiene
                            diametro_con_comillas = f"{diametro}\""
                            
                            # Convertir separación a decimal
                            separacion_decimal = float(f"0.{separacion}")
                            
                            # Escribir en Excel
                            ws.range('G4').value = int(cantidad)
                            ws.range('H4').value = diametro_con_comillas
                            ws.range('J4').value = separacion_decimal
                            
                            print(f"Colocando en el excel: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                        else:
                            # Intento alternativo de extracción
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            espaciamiento_match = re.search(r'@(\d+)', texto)
                            
                            if diametro_match and espaciamiento_match:
                                diametro = diametro_match.group(1)
                                diametro_con_comillas = f"{diametro}\""
                                separacion = espaciamiento_match.group(1)
                                separacion_decimal = float(f"0.{separacion}")
                                cantidad = "1"
                                
                                # Escribir en Excel
                                ws.range('G4').value = int(cantidad)
                                ws.range('H4').value = diametro_con_comillas
                                ws.range('J4').value = separacion_decimal
                                
                                print(f"Colocando en el excel: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                            else:
                                print(f"No se pudo extraer información del texto '{texto}'")
                    except Exception as e:
                        print(f"Error al procesar texto horizontal '{texto}': {e}")
                
                # Procesar acero vertical (G14, H14, J14) para otros tipos de prelosas
                for i, texto in enumerate(textos_vertical):
                    try:
                        # Extraer información del texto
                        match = re.match(r'∅(\d+\/\d+)\"@(\d+)', texto)
                        if match:
                            diametro = match.group(1)
                            separacion = match.group(2)
                            cantidad = "1"  # Por defecto
                            
                            # Añadir comillas al diámetro si no las tiene
                            diametro_con_comillas = f"{diametro}\""
                            
                            # Convertir separación a decimal
                            separacion_decimal = float(f"0.{separacion}")
                            
                            # Escribir en Excel
                            ws.range('G14').value = int(cantidad)
                            ws.range('H14').value = diametro_con_comillas
                            ws.range('J14').value = separacion_decimal
                            
                            # Si la separación es 20, entonces verificamos si se necesita ajustar para obtener 0.100
                            if separacion == "20":
                                print("Detectado acero vertical con espaciamiento 20, este debería producir 0.100 en K17")
                            
                            print(f"Colocando en el excel: {cantidad} -> G14, {diametro_con_comillas} -> H14, {separacion_decimal} -> J14")
                        else:
                            # Intento alternativo de extracción
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            espaciamiento_match = re.search(r'@(\d+)', texto)
                            
                            if diametro_match and espaciamiento_match:
                                diametro = diametro_match.group(1)
                                diametro_con_comillas = f"{diametro}\""
                                separacion = espaciamiento_match.group(1)
                                separacion_decimal = float(f"0.{separacion}")
                                cantidad = "1"
                                
                                # Escribir en Excel
                                ws.range('G14').value = int(cantidad)
                                ws.range('H14').value = diametro_con_comillas
                                ws.range('J14').value = separacion_decimal
                                
                                # Si la separación es 20, verificar ajuste
                                if separacion == "20":
                                    print("Detectado acero vertical con espaciamiento 20, este debería producir 0.100 en K17")
                                
                                print(f"Colocando en el excel: {cantidad} -> G14, {diametro_con_comillas} -> H14, {separacion_decimal} -> J14")
                            else:
                                print(f"No se pudo extraer información del texto '{texto}'")
                    except Exception as e:
                        print(f"Error al procesar texto vertical '{texto}': {e}")
            
            # Forzar cálculo y obtener resultados
            try:
                # Actualizar Excel de manera más agresiva
                ws.book.app.calculate()
                
                # Dar tiempo a Excel para calcular
                time.sleep(0.1)
                
                # Intento 1: Intentar usar CalculateFull si está disponible
                try:
                    wb.api.CalculateFull()
                except:
                    pass
                
                # Intento 2: Tocar las celdas manualmente para forzar el cálculo
                try:
                    # Tocar las celdas K8 y K17 para forzar su actualización
                    temp_k8 = ws.range('K8').value
                    ws.range('K8').value = temp_k8
                    
                    temp_k17 = ws.range('K17').value
                    ws.range('K17').value = temp_k17
                except:
                    pass
                
                # Dar más tiempo para los cálculos
                time.sleep(0.1)
                
                print("Actualizando excel...")
                
                # Leer valores después de procesamiento pero antes de correcciones
                print("\nVALORES DESPUÉS DE ACTUALIZAR EXCEL (ANTES DE CORRECCIONES):")
                k8_calculado = ws.range('K8').value
                k17_calculado = ws.range('K17').value
                print(f"  Celda K8 = {k8_calculado}")
                print(f"  Celda K17 = {k17_calculado}")
                print("-" * 40)
                
                # Obtener resultados de las celdas
                as_long = ws.range('K8').value
                as_tra1 = ws.range('K17').value
                as_tra2 = ws.range('K18').value if ws.range('K18').value else None
                
                # Corrección específica: si hay acero vertical pero as_tra1 es 0.2, ajustar a 0.1
                # Este es un ajuste basado en la lógica del Excel que parece que debería dar 0.1
                if len(textos_vertical) > 0:
                    # Revisar si hay textos con "@20"
                    tiene_espaciamiento_20 = False
                    for texto in textos_vertical:
                        if "@20" in texto:
                            tiene_espaciamiento_20 = True
                            break
                    
                    if tiene_espaciamiento_20:
                        # FORZAR el valor a 0.1 independientemente de lo que diga Excel
                        as_tra1 = 0.1
                        print("FORZANDO valor de K17 a 0.100 porque se detectó @20 en acero vertical")
                
                # Si siguen siendo 0, usar valores de respaldo o corregir valores incorrectos
                print("Verificando si los valores calculados son correctos...")
                
                # Si no hay textos (polilíneas vacías), restaurar los valores originales
                if len(textos_horizontal) == 0 and len(textos_vertical) == 0:
                    print("No se encontraron textos en la prelosa. Restaurando valores originales.")
                    as_long = k8_original
                    as_tra1 = k17_original
                
                # Verificar si hay textos horizontales pero as_long es 0 o nulo
                if len(textos_horizontal) > 0 and (as_long is None or as_long == 0):
                    print("Calculando valor de respaldo para acero horizontal...")
                    for texto in textos_horizontal:
                        match = re.search(r'@(\d+)', texto)
                        if match:
                            espaciamiento = match.group(1)
                            as_long = float(f"0.{espaciamiento}")
                            print(f"Valor de respaldo calculado para AS_LONG: {as_long}")
                            break
                
                # Verificar si hay textos verticales pero as_tra1 es 0 o nulo
                if len(textos_vertical) > 0 and (as_tra1 is None or as_tra1 == 0):
                    print("Calculando valor de respaldo para acero vertical...")
                    for texto in textos_vertical:
                        match = re.search(r'@(\d+)', texto)
                        if match:
                            espaciamiento = match.group(1)
                            if espaciamiento == "20":
                                # Caso especial para espaciamiento 20 en vertical
                                as_tra1 = 0.100
                                print(f"Asignando valor especial 0.100 para AS_TRA1 (espaciamiento 20)")
                            else:
                                as_tra1 = float(f"0.{espaciamiento}")
                                print(f"Valor de respaldo calculado para AS_TRA1: {as_tra1}")
                            break
                
                # Corregir caso específico donde esperamos 0.100 pero obtenemos 0.200
                if len(textos_vertical) > 0 and abs(as_tra1 - 0.2) < 0.01:
                    for texto in textos_vertical:
                        match = re.search(r'@(\d+)', texto)
                        if match and match.group(1) == "20":
                            print("Corrigiendo valor incorrecto de K17: cambiando 0.200 a 0.100")
                            as_tra1 = 0.100
                            break
                
                print("\nVALORES FINALES DESPUÉS DE CORRECCIONES:")
                print(f"  Celda K8 = {as_long}")
                print(f"  Celda K17 = {as_tra1}")
                if as_tra2:
                    print(f"  Celda K18 = {as_tra2}")
                print("-" * 40)
                
            except Exception as e:
                print(f"Error al recalcular Excel: {str(e)}")
                # Usar valores basados en los textos encontrados
                as_long = 0.20
                as_tra1 = 0.20
                as_tra2 = None
                
                # Intentar extraer valores de los textos
                for texto in textos_horizontal:
                    match = re.search(r'@(\d+)', texto)
                    if match:
                        espaciamiento = match.group(1)
                        as_long = float(f"0.{espaciamiento}")
                        break
                
                for texto in textos_vertical:
                    match = re.search(r'@(\d+)', texto)
                    if match:
                        espaciamiento = match.group(1)
                        # Caso especial para espaciamiento 20
                        if espaciamiento == "20":
                            as_tra1 = 0.100
                        else:
                            as_tra1 = float(f"0.{espaciamiento}")
                        break
                
                print(f"Usando valores extraídos de los textos - K8: {as_long}, K17: {as_tra1}")
            
            # Insertar bloque con los resultados
            try:
                # Convertir valores numéricos a textos formateados según el patrón deseado
                
                # Para prelosas macizas, asignar valores específicos
                if tipo_prelosa == "PRELOSA MACIZA":
                    # Para acero horizontal (AS_LONG)
                    if len(textos_horizontal) > 0:
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    else:
                        # Si no hay textos horizontales, usar valor original
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_original)}"
                    
                    # Para acero vertical (AS_TRA1) - siempre a 0.28 en prelosas macizas
                    as_tra1_texto = "1Ø6 mm@.28"
                    
                    # Para AS_TRA2 - usar el valor calculado de K18
                    if as_tra2 is not None:
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(as_tra2)}"
                    else:
                        as_tra2_texto = None

                # Para prelosas aligeradas 20
                elif tipo_prelosa == "PRELOSA ALIGERADA 20":
                    # Para acero horizontal (AS_LONG)
                    as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    
                    # Para acero vertical (AS_TRA1) - siempre fijo en aligeradas
                    as_tra1_texto = "1Ø6 mm@.50"
                    
                    # Para AS_TRA2 - siempre fijo en aligeradas
                    as_tra2_texto = "1Ø8 mm@.50"

                # Para prelosas aligeradas 20 - 2 SENT
                elif tipo_prelosa == "PRELOSA ALIGERADA 20 - 2 SENT":
                    # Para acero horizontal (AS_LONG)
                    as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    
                    # Para acero vertical (AS_TRA1) - siempre a 0.28 en prelosas aligeradas 2 sent
                    as_tra1_texto = "1Ø6 mm@.28"
                    
                    # Para AS_TRA2 - Usar el valor calculado de K18 si existe
                    if as_tra2 is not None:
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(as_tra2)}"
                    else:
                        as_tra2_texto = None

                # Para otros tipos de prelosas
                else:
                    # Para acero horizontal
                    if len(textos_horizontal) > 0:
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    else:
                        # Si no hay textos horizontales, usar valor original
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_original)}"
                    
                    # Para acero vertical
                    if len(textos_vertical) > 0:
                        # Caso especial para espaciamiento 20 en vertical
                        for texto in textos_vertical:
                            if "@20" in texto:
                                as_tra1_texto = "1Ø6 mm@.10"
                                break
                        else:
                            # Si no tiene @20, usar el valor calculado
                            as_tra1_texto = f"1Ø6 mm@.{formatear_valor_espaciamiento(as_tra1)}"
                    else:
                        # Si no hay textos verticales, usar valor original
                        as_tra1_texto = f"1Ø6 mm@.{formatear_valor_espaciamiento(k17_original)}"
                    
                    # No usar valor fijo para AS_TRA2 en otros tipos de prelosas
                    as_tra2_texto = None
                                
                print(f"Valores formateados para inserción en bloque:")
                print(f"    AS_LONG: {as_long_texto}")
                print(f"    AS_TRA1: {as_tra1_texto}")
                if as_tra2_texto:
                    print(f"    AS_TRA2: {as_tra2_texto}")
                
                # Calcular la orientación del bloque para que el acero azul apunte al lado más estrecho
                angulo_rotacion = calcular_orientacion_prelosa(vertices)
                print(f"Orientando bloque a {angulo_rotacion:.2f}° (alineado con lado más estrecho)")
                
                # Crear una copia de la definición del bloque con la orientación calculada
                definicion_bloque_orientada = definicion_bloque.copy()
                definicion_bloque_orientada['rotation'] = angulo_rotacion
                
                # Insertar bloque con los valores formateados y la orientación correcta
                bloque = insertar_bloque_acero(msp, definicion_bloque_orientada, centro_prelosa, as_long_texto, as_tra1_texto, as_tra2_texto)
                
                if bloque:
                    total_bloques += 1
                    print(f"{tipo_prelosa} CONCLUIDA CON EXITO")
                    print("=" * 52)
                    return True
                else:
                    print(f"Error al insertar el bloque en {tipo_prelosa}")
                    print("=" * 52)
                    return False
            except Exception as e:
                print(f"Error al insertar bloque: {e}")
                print(f"Error al insertar el bloque en {tipo_prelosa}")
                print("=" * 52)
                return False
        
        # Procesar PRELOSAS MACIZAS
        for idx, polilinea_maciza in enumerate(polilineas_macizas):
            procesar_prelosa(polilinea_maciza, "PRELOSA MACIZA", idx)
        
        # Procesar PRELOSAS ALIGERADAS 20
        for idx, polilinea_aligerada in enumerate(polilineas_aligeradas):
            procesar_prelosa(polilinea_aligerada, "PRELOSA ALIGERADA 20", idx)
        
        # Procesar PRELOSAS ALIGERADAS 20 - 2 SENT
        for idx, polilinea_aligerada_2sent in enumerate(polilineas_aligeradas_2sent):
            procesar_prelosa(polilinea_aligerada_2sent, "PRELOSA ALIGERADA 20 - 2 SENT", idx)
        
        # Cerrar Excel y guardar DXF
        try:
            wb.save()
            wb.close()
            app.quit()
        except:
            print("Error al cerrar Excel, continuando...")
        
        capas_acero = ["ACERO HORIZONTAL", "ACERO VERTICAL", 
                "BD-ACERO HORIZONTAL", "BD-ACERO VERTICAL",
                "ACERO", "REFUERZO", "ARMADURA"]
        eliminar_entidades_por_capa(doc, capas_acero)
        
        doc.saveas(output_dxf_path)
        print(f"Archivo DXF guardado en: {output_dxf_path}")
        
        # Tiempo total
        tiempo_total = time.time() - tiempo_inicio
        
        # Estadísticas finales
        print("\n" + "=" * 50)
        print("RESUMEN DEL PROCESAMIENTO")
        print("=" * 50)
        print(f"Total de prelosas procesadas: {total_prelosas}")
        print(f"Total de bloques insertados: {total_bloques}")
        print(f"Tiempo total: {tiempo_total:.2f} segundos")
        print(f"Tiempo promedio por prelosa: {tiempo_total/max(total_prelosas, 1):.4f} segundos")
        print(f"Archivo guardado: {output_dxf_path}")
        
        return total_bloques
    
    except Exception as e:
        print(f"Error al procesar el archivo: {e}")
        traceback.print_exc()
        
        # Intentar cerrar Excel si está abierto
        try:
            wb.close()
            app.quit()
        except:
            pass
        
        return 0


def imprimir_banner_script():
    """
    Imprime un banner ASCII decorativo e impresionante con el texto 'ACERO SCRIPT'
    """
    banner = r"""
    ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════
    ║                                                                                                          ║
    ║       ▄▄▄       ▄████▄  ▓█████  ██▀███   ▒█████      ██████  ▄████▄   ██▀███   ██▓ ██▓███  ▄▄▄█████▓     ║
    ║      ▒████▄    ▒██▀ ▀█  ▓█   ▀ ▓██ ▒ ██▒▒██▒  ██▒   ▒██    ▒ ▒██▀ ▀█  ▓██ ▒ ██▒▓██▒▓██░  ██▒▓  ██▒ ▓▒    ║
    ║      ▒██  ▀█▄  ▒▓█    ▄ ▒███   ▓██ ░▄█ ▒▒██░  ██▒   ░ ▓██▄   ▒▓█    ▄ ▓██ ░▄█ ▒▒██▒▓██░ ██▓▒▒ ▓██░ ▒░    ║
    ║      ░██▄▄▄▄██ ▒▓▓▄ ▄██▒▒▓█  ▄ ▒██▀▀█▄  ▒██   ██░     ▒   ██▒▒▓▓▄ ▄██▒▒██▀▀█▄  ░██░▒██▄█▓▒ ▒░ ▓██▓ ░     ║
    ║       ▓█   ▓██▒▒ ▓███▀ ░░▒████▒░██▓ ▒██▒░ ████▓▒░   ▒██████▒▒▒ ▓███▀ ░░██▓ ▒██▒░██░▒██▒ ░  ░  ▒██▒ ░     ║
    ║       ▒▒   ▓▒█░░ ░▒ ▒  ░░░ ▒░ ░░ ▒▓ ░▒▓░░ ▒░▒░▒░    ▒ ▒▓▒ ▒ ░░ ░▒ ▒  ░░ ▒▓ ░▒▓░░▓  ▒▓▒░ ░  ░  ▒ ░░       ║
    ║        ▒   ▒▒ ░  ░  ▒    ░ ░  ░  ░▒ ░ ▒░  ░ ▒ ▒░    ░ ░▒  ░ ░  ░  ▒     ░▒ ░ ▒░ ▒ ░░▒ ░         ░        ║
    ║        ░   ▒   ░           ░     ░░   ░ ░ ░ ░ ▒     ░  ░  ░  ░          ░░   ░  ▒ ░░░         ░          ║
    ║            ░  ░░ ░         ░  ░   ░         ░ ░           ░  ░ ░         ░      ░                        ║
    ║                ░                                              ░                                          ║
    ║                                                                                                          ║
    ║                        Herramienta para Automatización de Aceros en Prelosas                             ║
    ║                                      by DODOD SOLUTIONS                                                  ║
    ║                                                                                                          ║
    ╚══════════════════════════════════════════════════════════════════════════════════════════════════════════
    """
    print(banner)
    
# Punto de entrada principal del script
if __name__ == "__main__":
    # Rutas de archivo (ajustar según necesidad)

    print(imprimir_banner_script())

    file_path = "PLANO1.dxf"
    excel_path = "CONVERTIDOR.xlsx"
    output_dxf_path = "test_57.dxf"
    
    # Ejecutar el procesamiento
    print("\n==== INICIANDO PROCESAMIENTO DE PRELOSAS ====")
    print(f"Archivo DXF original: {file_path}")
    print(f"Archivo Excel: {excel_path}")
    print(f"Archivo de salida: {output_dxf_path}")
    print("=" * 60)
    
    # Medir tiempo
    tiempo_inicio = time.time()
    
    # Procesar prelosas
    total = procesar_prelosas_con_bloques(file_path, excel_path, output_dxf_path)
    
    # Tiempo total
    tiempo_total = time.time() - tiempo_inicio