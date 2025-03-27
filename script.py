import ezdxf
from shapely.geometry import Point, Polygon
import re
import os
import sys
import xlwings as xw
import traceback
import time
import random


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

    def validar_formato_texto(texto):
        # Reemplazar caracteres especiales básicos
        texto = reemplazar_caracteres_especiales(texto)
        
        # Cambiar comillas simples '' por comillas rectas "
        texto = texto.replace("''", '"')
        
        # Validación de diámetro con unidades (Ej: 1Ø8mm@.20)
        # Agregar espacio entre número y unidad si no existe
        texto = re.sub(r'(\d+)Ø(\d+)mm', r'\1Ø \2 mm', texto)
        
        # Validación de comillas
        # Asegurar que las comillas sean rectas " en lugar de curvas
        texto = texto.replace('"', '"').replace('"', '"')
        
        return texto

    for text in textos:
        # Procesar textos normales (TEXT, MTEXT)
        if text.dxftype() in ['TEXT', 'MTEXT']:
            punto_texto = Point(text.dxf.insert)
            if poligono.contains(punto_texto):
                if text.dxftype() == 'MTEXT':
                    texto_contenido = text.text
                else:
                    texto_contenido = text.dxf.text

                # Aplicar validaciones de formato
                texto_formateado = validar_formato_texto(texto_contenido)
                
                # Solo añadir si el texto no está vacío después de formatear
                if texto_formateado.strip():
                    textos_en_polilinea.append(texto_formateado)
        
        # Procesar bloques con atributos (INSERT)
        elif text.dxftype() == 'INSERT':
            punto_bloque = Point(text.dxf.insert)
            if poligono.contains(punto_bloque):
                try:
                    # Método 1: Obtener atributos usando .attribs
                    if hasattr(text, 'attribs'):
                        for attrib in text.attribs:
                            if hasattr(attrib.dxf, 'tag') and hasattr(attrib.dxf, 'text'):
                                # Filtrar atributos relevantes
                                if attrib.dxf.tag in ['ACERO', 'AS_LONG', 'AS_TRA1', 'AS_TRA2']:
                                    texto_contenido = attrib.dxf.text
                                    texto_formateado = validar_formato_texto(texto_contenido)
                                    if texto_formateado.strip():
                                        print(f"Encontrado atributo en bloque: {attrib.dxf.tag} = {texto_formateado}")
                                        textos_en_polilinea.append(texto_formateado)
                except:
                    pass
                
                try:
                    # Método 2: Obtener atributos iterando sobre los hijos
                    for child in text:
                        if child.dxftype() == 'ATTRIB':
                            if hasattr(child.dxf, 'tag') and hasattr(child.dxf, 'text'):
                                if child.dxf.tag in ['ACERO', 'AS_LONG', 'AS_TRA1', 'AS_TRA2']:
                                    texto_contenido = child.dxf.text
                                    texto_formateado = validar_formato_texto(texto_contenido)
                                    if texto_formateado.strip():
                                        print(f"Encontrado atributo (método 2): {child.dxf.tag} = {texto_formateado}")
                                        textos_en_polilinea.append(texto_formateado)
                except:
                    pass
                
                try:
                    # Método 3: Usar get_attribs si está disponible
                    if hasattr(text, 'get_attribs'):
                        for attrib in text.get_attribs():
                            if hasattr(attrib.dxf, 'tag') and hasattr(attrib.dxf, 'text'):
                                if attrib.dxf.tag in ['ACERO', 'AS_LONG', 'AS_TRA1', 'AS_TRA2']:
                                    texto_contenido = attrib.dxf.text
                                    texto_formateado = validar_formato_texto(texto_contenido)
                                    if texto_formateado.strip():
                                        print(f"Encontrado atributo (método 3): {attrib.dxf.tag} = {texto_formateado}")
                                        textos_en_polilinea.append(texto_formateado)
                except:
                    pass

    return textos_en_polilinea

# Función para obtener polilíneas dentro de una polilínea principal
def obtener_polilineas_dentro_de_polilinea(polilinea_principal, polilineas_anidadas):
    """
    Obtiene las polilíneas que están dentro o intersectan con una polilínea principal.
    """
    # Capas válidas de acero (case-insensitive)
    capas_acero_validas = [
        "ACERO LONGITUDINAL", 
        "ACERO TRANSVERSAL", 
        "BD-ACERO LONGITUDINAL", 
        "BD-ACERO TRANSVERSAL",
        "ACERO LONG ADI",
        "ACERO TRA ADI"
    ]

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
            # Filtrar por capas de acero (case-insensitive)
            capa_polilinea = polilinea.dxf.layer.upper()
            if not any(capa_acero.upper() in capa_polilinea for capa_acero in capas_acero_validas):
                continue

            vertices_anidada = [(p[0], p[1]) for p in polilinea.get_points('xy')]
            poligono_anidado = Polygon(vertices_anidada)
            
            # Verificar intersección o contención
            if poligono_principal.intersects(poligono_anidado) or poligono_principal.contains(poligono_anidado):
                area_interseccion = poligono_principal.intersection(poligono_anidado).area
                area_anidada = poligono_anidado.area
                
                # Calcular ratio de intersección
                ratio_interseccion = area_interseccion / area_anidada if area_anidada > 0 else 0
                
                # Si al menos el 20% de la polilínea está dentro, considerarla
                if ratio_interseccion >= 0.2:
                    print(f"Polilínea en capa '{polilinea.dxf.layer}' intersecta con {ratio_interseccion:.2f} de su área")
                    polilineas_dentro.append(polilinea)
        except Exception as e:
            print(f"Error al procesar polilínea: {e}")
    
    # Información de depuración adicional
    print(f"Total de polilíneas de acero encontradas dentro de la prelosa: {len(polilineas_dentro)}")
    
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

# Función para calcular el espaciamiento de acero
# def calcular_orientacion_prelosa(vertices):
#     """
#     Calcula la orientación de la prelosa analizando la longitud de sus lados.
#     Devuelve el ángulo de rotación en grados para alinear el bloque con el lado más corto
#     sin que el texto se invierta o aparezca al revés.
#     """
#     try:
#         import math
        
#         # Si la prelosa no tiene al menos 3 vértices, devolver rotación por defecto
#         if len(vertices) < 3:
#             return 0.0
            
#         # Calcular la longitud de los lados
#         lados = []
#         for i in range(len(vertices)):
#             x1, y1 = vertices[i]
#             x2, y2 = vertices[(i + 1) % len(vertices)]
            
#             # Calcular longitud del lado
#             longitud = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
            
#             # Calcular ángulo del lado (en radianes)
#             angulo = math.atan2(y2 - y1, x2 - x1)
            
#             lados.append((longitud, angulo))
        
#         # Ordenar los lados por longitud (ascendente)
#         lados_ordenados = sorted(lados, key=lambda x: x[0])
        
#         # Obtener el ángulo del lado más corto (en radianes)
#         angulo_lado_corto = lados_ordenados[0][1]
        
#         # Convertir a grados
#         angulo_grados = math.degrees(angulo_lado_corto)
        
#         # Forzar el ángulo a estar entre 0 y 180 grados
#         # Esto evita la inversión del texto en el bloque
#         if angulo_grados < 0:
#             angulo_grados += 180
#         elif angulo_grados > 180:
#             angulo_grados -= 180
            
#         print(f"Orientación final: {angulo_grados:.2f}° (texto correctamente orientado)")
        
#         return angulo_grados
        
#     except Exception as e:
#         print(f"Error al calcular la orientación: {e}")
#         return 0.0

def insertar_bloque_acero(msp, definicion_bloque, centro, as_long, as_tra1, as_tra2=None):
    """
    Inserta un bloque de acero en el centro de la prelosa con los valores calculados.
    Asegura que los atributos sean visibles y editables.
    """
    try:
        print("Copiar y Pegando bloque de aceros en centro de la prelosa...")
        
        # Usar directamente los valores recibidos
        str_as_long = as_long
        str_as_tra1 = as_tra1
        str_as_tra2 = as_tra2 if as_tra2 is not None else ""
        
        print("Modificando Texto del bloque de aceros...")
        print("texto modificado:")
        print(f"    AS_LONG: {str_as_long}")
        print(f"    AS_TRA1: {str_as_tra1}")
        if as_tra2 is not None:
            print(f"    AS_TRA2: {str_as_tra2}")
        
        # Verificar y desbloquear la capa si es necesario
        capa_destino = definicion_bloque.get('capa', 'BD-ACERO POSITIVO')
        
        # Verificar si la capa existe y desbloquearla
        doc = msp.doc
        if capa_destino in doc.layers:
            layer = doc.layers.get(capa_destino)
            if hasattr(layer.dxf, 'lock') and layer.dxf.lock:
                layer.dxf.lock = False
                print(f"Desbloqueando capa '{capa_destino}' para permitir edición")
        
        # Obtener la rotación original
        rotation = definicion_bloque.get('rotation', 0.0)
        print(f"    Rotación original del bloque: {rotation:.2f}°")
        
        # Normalizar a 0-360
        rotation = rotation % 360
        
        # Determinar el cuadrante y ajustar
        corrected_rotation = rotation
        
        # Horizontales (lado derecho hacia abajo)
        if 150 <= rotation <= 210:
            corrected_rotation = (rotation + 180) % 360
            print(f"    Ajustando ángulo horizontal invertido de {rotation:.2f}° a {corrected_rotation:.2f}°")
        
        # Verticales (de cabeza)
        elif 240 <= rotation <= 300:
            corrected_rotation = (rotation + 180) % 360
            print(f"    Ajustando ángulo vertical invertido de {rotation:.2f}° a {corrected_rotation:.2f}°")
        
        # Configurar escalas
        xscale = definicion_bloque.get('xscale', 1.0)
        yscale = definicion_bloque.get('yscale', 1.0)
        
        # Usar escalas positivas
        xscale = abs(xscale)
        yscale = abs(yscale)
        
        # CAMBIO IMPORTANTE: Método alternativo de inserción para preservar atributos
        try:
            # Intentar usar el método insert_block que trata mejor los atributos
            blockref = msp.insert_block(
                insert=centro,
                name=definicion_bloque['nombre'],
                dxfattribs={
                    'layer': capa_destino,
                    'xscale': xscale,
                    'yscale': yscale,
                    'rotation': corrected_rotation
                }
            )
            
            # Preparar valores de atributos
            valores_atributos = {
                'AS_LONG': str_as_long,
                'AS_TRA1': str_as_tra1
            }
            
            if as_tra2 is not None and as_tra2 != "":
                valores_atributos['AS_TRA2'] = str_as_tra2
            
            # Asignar valores a los atributos y configurarlos como visibles y editables
            for attrib in blockref.attribs:
                if attrib.dxf.tag in valores_atributos:
                    attrib.dxf.text = valores_atributos[attrib.dxf.tag]
                    
                    # Configurar flags del atributo para visualización y edición
                    if hasattr(attrib.dxf, 'invisible'):
                        attrib.dxf.invisible = 0  # 0 = visible
                    if hasattr(attrib.dxf, 'constant'):
                        attrib.dxf.constant = 0  # 0 = editable
                    if hasattr(attrib.dxf, 'verify'):
                        attrib.dxf.verify = 1    # 1 = verificar al editar
                    if hasattr(attrib.dxf, 'preset'):
                        attrib.dxf.preset = 0    # 0 = no preestablecido
            
            print("Bloque insertado usando insert_block con atributos visibles y editables")
            return blockref
        
        except Exception as e:
            print(f"Error al usar insert_block: {e}")
            print("Cayendo al método tradicional add_blockref...")
        
        # MÉTODO TRADICIONAL (RESPALDO)
        # Crear inserción del bloque con la rotación corregida
        bloque = msp.add_blockref(
            name=definicion_bloque['nombre'],
            insert=centro,
            dxfattribs={
                'layer': capa_destino,
                'xscale': xscale,
                'yscale': yscale,
                'rotation': corrected_rotation
            }
        )
        
        # Preparar valores de atributos
        valores_atributos = {
            'AS_LONG': str_as_long,
            'AS_TRA1': str_as_tra1
        }
        
        if as_tra2 is not None and as_tra2 != "":
            valores_atributos['AS_TRA2'] = str_as_tra2
        
        # Método 1: add_auto_attribs es el más confiable para asignar atributos
        try:
            bloque.add_auto_attribs(valores_atributos)
            
            # Después de asignar, configurar cada atributo para hacerlo visible y editable
            for attrib in bloque.attribs:
                # Configurar flags del atributo
                attrib.dxf.invisible = 0  # 0 = visible
                # Más configuraciones si están disponibles
                if hasattr(attrib.dxf, 'constant'):
                    attrib.dxf.constant = 0
                if hasattr(attrib.dxf, 'verify'):
                    attrib.dxf.verify = 1
                if hasattr(attrib.dxf, 'preset'):
                    attrib.dxf.preset = 0
            
            print("Atributos asignados y configurados como visibles usando add_auto_attribs")
            return bloque
        except Exception as e:
            print(f"No se pudo usar add_auto_attribs: {e}")
        
        # Método 2: Asignar atributos uno por uno
        try:
            for attrib in bloque.attribs:
                if attrib.dxf.tag in valores_atributos:
                    attrib.dxf.text = valores_atributos[attrib.dxf.tag]
                    attrib.dxf.invisible = 0
                    
                    # Más configuraciones si están disponibles
                    if hasattr(attrib.dxf, 'constant'):
                        attrib.dxf.constant = 0
                    if hasattr(attrib.dxf, 'verify'):
                        attrib.dxf.verify = 1
                    if hasattr(attrib.dxf, 'preset'):
                        attrib.dxf.preset = 0
            
            print("Atributos asignados y configurados manualmente")
            return bloque
        except Exception as e:
            print(f"No se pudo asignar atributos manualmente: {e}")
        
        # Si llegamos aquí, intenta un último recurso
        try:
            # Intenta acceder a los atributos usando .get_attribs()
            for attrib in bloque.get_attribs():
                if attrib.dxf.tag in valores_atributos:
                    attrib.dxf.text = valores_atributos[attrib.dxf.tag]
                    attrib.dxf.invisible = 0
            
            print("Atributos asignados usando get_attribs")
            return bloque
        except Exception as e:
            print(f"Error en último intento: {e}")
            return bloque  # Devolver el bloque aunque haya errores
    
    except Exception as e:
        print(f"Error global al insertar bloque: {e}")
        traceback.print_exc()
        return None

# Función para eliminar entidades por capa
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



def formatear_valor_espaciamiento(valor):
    """
    Formatea un valor decimal para mostrarlo correctamente en el bloque.
    Si el valor termina en 0, se elimina el último dígito.
    
    Args:
        valor (float): Valor decimal (por ejemplo 0.200, 0.350, 0.175)
        
    Returns:
        str: Valor formateado (por ejemplo "20", "35", "175")
    """
    if valor is None:
        return "20"  # Valor por defecto
    
    # Multiplicar por 1000 y redondear para evitar errores de punto flotante
    valor_entero = round(float(valor) * 1000)
    
    # Convertir a string
    valor_str = str(valor_entero)
    
    # Si termina en 0, eliminar el último dígito
    if valor_str.endswith("0"):
        valor_str = valor_str[:-1]
    
    return valor_str


# Función principal modificada para usar bloques en lugar de textos
# Función principal modificada para usar bloques en lugar de textos
def procesar_prelosas_con_bloques(file_path, excel_path, output_dxf_path, valores_predeterminados=None):
    """
    Procesa las prelosas identificando tipos y contenidos,
    calcula valores usando Excel y coloca bloques con los resultados.
    
    Args:
        file_path (str): Ruta del archivo DXF de entrada
        excel_path (str): Ruta del archivo Excel
        output_dxf_path (str): Ruta del archivo DXF de salida
        valores_predeterminados (dict, optional): Valores predeterminados para cada tipo de prelosa
    """

    banner = """
        ╔════════════════════════════════════════════════════════════════════════════════════════════════╗
        ║                                                                                                ║
        ║     █████╗  ██████╗███████╗██████╗  ██████╗     ███████╗ ██████╗██████╗ ██╗██████╗████████╗    ║
        ║    ██╔══██╗██╔════╝██╔════╝██╔══██╗██╔═══██╗    ██╔════╝██╔════╝██╔══██╗██║██╔══██╗╚══██╔══╝   ║
        ║    ███████║██║     █████╗  ██████╔╝██║   ██║    ███████╗██║     ██████╔╝██║██████╔╝   ██║      ║
        ║    ██╔══██║██║     ██╔══╝  ██╔══██╗██║   ██║    ╚════██║██║     ██╔══██╗██║██╔═══╝    ██║      ║
        ║    ██║  ██║╚██████╗███████╗██║  ██║╚██████╔╝    ███████║╚██████╗██║  ██║██║██║        ██║      ║
        ║    ╚═╝  ╚═╝ ╚═════╝╚══════╝╚═╝  ╚═╝ ╚═════╝     ╚══════╝ ╚═════╝╚═╝  ╚═╝╚═╝╚═╝        ╚═╝      ║
        ║                                                                                                ║
        ║                    Herramienta para Automatización de Aceros en Prelosas                       ║
        ║                                     by DODOD SOLUTIONS                                         ║
        ║                                       v1.0.0 (2025)                                            ║
        ║                                                                                                ║
        ╚════════════════════════════════════════════════════════════════════════════════════════════════╝

    """
    print(banner)


    # Valores predeterminados por defecto
    default_valores = {
        'PRELOSA MACIZA': {
            'espaciamiento': '0.20'
        },
        'PRELOSA ALIGERADA 20': {
            'espaciamiento': '0.20'
        },
        'PRELOSA ALIGERADA 20 - 2 SENT': {
            'espaciamiento': '0.605'
        }
    }

    # Combinar valores predeterminados proporcionados con los predeterminados
    if valores_predeterminados:
        for tipo, valores in valores_predeterminados.items():
            if tipo in default_valores:
                default_valores[tipo].update(valores)

    print("Valores predeterminados combinados:")
    print(default_valores)
    """
    Procesa las prelosas identificando tipos y contenidos,
    calcula valores usando Excel y coloca bloques con los resultados.
    """
    try:
        tiempo_inicio = time.time()
        
        # Cargar el documento DXF
        doc = ezdxf.readfile(file_path)
        msp = doc.modelspace()
        
        # Abrir Excel
        app = xw.App(visible=False)  # Abrir Excel en segundo plano
        wb = app.books.open(excel_path)  # Abrir el archivo
        ws = wb.sheets.active  # Obtener la hoja activa
        
        # NUEVO: Limpiar celdas antes de empezar
        print("Limpiando celdas en Excel para evitar interferencias...")
        try:
            # Limpiar celdas de acero horizontal - valores por defecto

            
            # Limpiar segundas filas horizontales
            ws.range('G5').value = 0

            
            # Limpiar celdas de acero vertical - valores por defecto

            
            # Limpiar segundas filas verticales
            ws.range('G15').value = 0

            # Forzar cálculo para actualizar con estos valores por defecto
            wb.app.calculate()
            
            print("Celdas limpiadas y valores por defecto establecidos.")
        except Exception as e:
            print(f"Error al limpiar celdas: {e}")
        
        # Leer los valores iniciales de las celdas K8 y K17
        k8_original = ws.range('K8').value
        k17_original = ws.range('K17').value
        
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
             entity.dxf.layer in ["ACERO LONGITUDINAL", "ACERO TRANSVERSAL", "ACERO LONG ADI", "ACERO TRA ADI",
                                  "BD-ACERO LONGITUDINAL", "BD-ACERO TRANSVERSAL",
                                  "ACERO", "REFUERZO", "ARMADURA"]]
        textos = [entity for entity in msp if entity.dxftype() in ['TEXT', 'MTEXT', 'INSERT']]
        
        # Contadores para estadísticas
        total_prelosas = 0
        total_bloques = 0
        
# Print all layer names in the DXF file

        def calcular_orientacion_prelosa(vertices, polilineas_longitudinal=None, polilineas_long_adi=None):
            """
            Calcula la orientación del bloque con la siguiente prioridad:
            1. Dirección de ACERO LONGITUDINAL
            2. Dirección de ACERO LONG ADI
            3. Orientación hacia el lado más ESTRECHO de la prelosa
            
            Args:
                vertices: Lista de puntos (x, y) que forman la polilínea de la prelosa
                polilineas_longitudinal: Lista de polilíneas de ACERO LONGITUDINAL
                polilineas_long_adi: Lista de polilíneas de ACERO LONG ADI
                
            Returns:
                float: Ángulo de rotación en grados
            """
            try:
                import math
                import numpy as np
                
                # 1. Primero intentar con polilíneas longitudinales regulares
                if polilineas_longitudinal and len(polilineas_longitudinal) > 0:
                    print("Usando orientación de ACERO LONGITUDINAL para el bloque")
                    
                    # Obtener la primera polilínea longitudinal
                    polilinea_long = polilineas_longitudinal[0]
                    vertices_long = polilinea_long.get_points('xy')
                    
                    # Necesitamos al menos 2 puntos para determinar una dirección
                    if len(vertices_long) >= 2:
                        # Calcular la dirección principal de la polilínea longitudinal
                        vertices_long_array = np.array(vertices_long)
                        
                        # Calcular el ángulo usando los primeros dos puntos
                        x1, y1 = vertices_long_array[0]
                        x2, y2 = vertices_long_array[1]
                        
                        # Calcular la dirección de la línea
                        dx = x2 - x1
                        dy = y2 - y1
                        
                        # Calcular ángulo en grados
                        angulo = math.degrees(math.atan2(dy, dx))
                        
                        # Normalizar a 0-180 grados (para evitar texto al revés)
                        angulo_final = angulo % 180
                        
                        print(f"Dirección de ACERO LONGITUDINAL detectada: {angulo_final:.2f}°")
                        return angulo_final
                
                # 2. Si no hay ACERO LONGITUDINAL, intentar con ACERO LONG ADI
                if polilineas_long_adi and len(polilineas_long_adi) > 0:
                    print("No se encontró ACERO LONGITUDINAL. Usando orientación de ACERO LONG ADI para el bloque")
                    
                    # Obtener la primera polilínea de acero adicional
                    polilinea_long_adi = polilineas_long_adi[0]
                    vertices_long_adi = polilinea_long_adi.get_points('xy')
                    
                    # Necesitamos al menos 2 puntos para determinar una dirección
                    if len(vertices_long_adi) >= 2:
                        # Calcular la dirección principal de la polilínea longitudinal adicional
                        vertices_long_adi_array = np.array(vertices_long_adi)
                        
                        # Calcular el ángulo usando los primeros dos puntos
                        x1, y1 = vertices_long_adi_array[0]
                        x2, y2 = vertices_long_adi_array[1]
                        
                        # Calcular la dirección de la línea
                        dx = x2 - x1
                        dy = y2 - y1
                        
                        # Calcular ángulo en grados
                        angulo = math.degrees(math.atan2(dy, dx))
                        
                        # Normalizar a 0-180 grados (para evitar texto al revés)
                        angulo_final = angulo % 180
                        
                        print(f"Dirección de ACERO LONG ADI detectada: {angulo_final:.2f}°")
                        return angulo_final
                
                # 3. Si no se pudo determinar orientación con polilíneas, usar el método de caja contenedora
                print("No se pudo determinar orientación por ACERO LONGITUDINAL ni ACERO LONG ADI, usando método de caja")
                
                # Convertir vértices a array NumPy
                vertices_array = np.array(vertices)
                
                # Calcular caja contenedora
                min_x = np.min(vertices_array[:, 0])
                max_x = np.max(vertices_array[:, 0])
                min_y = np.min(vertices_array[:, 1])
                max_y = np.max(vertices_array[:, 1])
                
                ancho = max_x - min_x
                alto = max_y - min_y
                
                # CAMBIO: Orientar para que la línea azul apunte al lado más ESTRECHO (como en la versión original)
                if ancho < alto:  # Si el ancho es menor que el alto
                    # Prelosa más alta que ancha -> línea azul horizontal (0°)
                    angulo_final = 0.0
                    print(f"Prelosa vertical (más alta que ancha). Orientando bloque horizontalmente: {angulo_final}°")
                else:
                    # Prelosa más ancha que alta -> línea azul vertical (90°)
                    angulo_final = 90.0
                    print(f"Prelosa horizontal (más ancha que alta). Orientando bloque verticalmente: {angulo_final}°")
                
                return angulo_final
                    
            except Exception as e:
                print(f"Error al calcular la orientación: {e}")
                traceback.print_exc()
                return 0.0  # Valor por defecto en caso de error
        # FUNCIÓN AUXILIAR: Procesa una prelosa (se aplica a todos los tipos)
        def procesar_prelosa(polilinea, tipo_prelosa, idx):
            nonlocal total_prelosas, total_bloques
            
            total_prelosas += 1
            vertices = polilinea.get_points('xy')
            centro_prelosa = calcular_centro_polilinea(vertices)
            polilineas_dentro = obtener_polilineas_dentro_de_polilinea(vertices, polilineas_acero)
            
            print(f"{tipo_prelosa} numero {idx+1} encontrada:")
            print(f"Centro de la prelosa: ({centro_prelosa[0]}, {centro_prelosa[1]})")
            print(f"Polilíneas dentro encontradas: {len(polilineas_dentro)}")
            
            # Variables para almacenar textos por tipo de acero
            textos_longitudinal = []
            textos_transversal = []
            textos_adicionales = []
            textos_long_adi = []   # Nuevo
            textos_tra_adi = []   # Nuevo
            
            # Procesar polilíneas de acero
            for polilinea_anidada in polilineas_dentro:
                vertices_anidada = polilinea_anidada.get_points('xy')
                textos_dentro = obtener_textos_dentro_de_polilinea(vertices_anidada, textos)
                
                print(f"Polilínea anidada en {tipo_prelosa.lower()} {idx+1} tiene {len(textos_dentro)} textos dentro.")
                
                # Clasificar textos según el tipo de acero
                tipo_acero = polilinea_anidada.dxf.layer.upper()
                if "LONGITUDINAL" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO LONGITUDINAL: {texto}")
                        textos_longitudinal.append(texto)
                elif "TRANSVERSAL" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO TRANSVERSAL: {texto}")
                        textos_transversal.append(texto)
                elif "ACERO LONG ADI" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO LONG ADI: {texto}")
                        textos_long_adi.append(texto)
                elif "ACERO TRA ADI" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO TRA ADI: {texto}")
                        textos_tra_adi.append(texto)
                elif "ADICIONAL" in tipo_acero:
                    for texto in textos_dentro:
                        print(f"Texto encontrado en ACERO ADICIONAL: {texto}")
                        textos_adicionales.append(texto)
            
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
                print("----------------------------------------")
                print("INICIANDO PROCESAMIENTO DE PRELOSA ALIGERADA 20")
                print("----------------------------------------")
                
                # Usar el espaciamiento de los valores predeterminados
                espaciamiento_aligerada = float(default_valores.get('PRELOSA ALIGERADA 20', {}).get('espaciamiento', 0.605))
                print(f"Usando espaciamiento predeterminado para PRELOSA ALIGERADA 20: {espaciamiento_aligerada}")
                
                # Imprimir todos los textos encontrados para depuración
                print("TEXTOS ENCONTRADOS PARA DEPURACIÓN:")
                print(f"Textos transversales ({len(textos_transversal)}): {textos_transversal}")
                print(f"Textos longitudinales ({len(textos_longitudinal)}): {textos_longitudinal}")
                
                # Combinar textos verticales y horizontales para procesar
                # Primero los verticales y luego los horizontales (si hay)
                textos_a_procesar = textos_transversal + textos_longitudinal
                print(f"Total textos a procesar (vertical + horizontal): {len(textos_a_procesar)}")
                
                # Procesar los textos (independientemente si son verticales u horizontales)
                if len(textos_a_procesar) > 0:
                    print(f"Procesando {len(textos_a_procesar)} textos en PRELOSA ALIGERADA 20")
                    
                    # Procesar primer texto (G4, H4, J4)
                    if len(textos_a_procesar) >= 1:
                        #limpiar celda g5
                        ws.range('G5').value = 0
                        texto = textos_a_procesar[0]
                        print(f"Procesando primer texto: '{texto}'")
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            print(f"Cantidad extraída: {cantidad}")
                            
                            # Extraer diámetro del texto con manejo mejorado para mm
                            if "mm" in texto:
                                # Caso específico para milímetros
                                diametro_match = re.search(r'∅(\d+)mm', texto)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    diametro_con_comillas = f"{diametro}mm"
                            else:
                                # Caso para fraccionales con o sin comillas
                                diametro_match = re.search(r'∅([\d/]+)', texto)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    # Asegurarnos de añadir comillas si es necesario
                                    if "\"" not in diametro and "/" in diametro:
                                        diametro_con_comillas = f"{diametro}\""
                                    else:
                                        diametro_con_comillas = diametro
                            
                            if diametro_match:
                                print(f"Diámetro extraído: {diametro_con_comillas}")
                                
                                cantidad = int(cantidad)  # Convertir a entero
                                
                                # NUEVO: Verificar si el texto incluye información de espaciamiento
# Verificar si el texto incluye información de espaciamiento (funciona para @30 y @.30)
                                espaciamiento_match = re.search(r'@\.?(\d+)', texto)
                                if espaciamiento_match:
                                    # Si encuentra espaciamiento en el texto, lo usa
                                    separacion = espaciamiento_match.group(1)
                                    separacion_decimal = float(f"0.{separacion}")
                                    print(f"Espaciamiento encontrado en texto: @{separacion} -> {separacion_decimal}")
                                else:
                                    # Si no encuentra espaciamiento, usa el valor predeterminado
                                    separacion_decimal = espaciamiento_aligerada
                                    print(f"No se encontró espaciamiento en texto, usando valor predeterminado: {espaciamiento_aligerada}")
                                # Escribir en Excel
                                print(f"Escribiendo en Excel: G4={cantidad}, H4={diametro_con_comillas}, J4={separacion_decimal}")
                                ws.range('G4').value = cantidad
                                ws.range('H4').value = diametro_con_comillas
                                ws.range('J4').value = separacion_decimal
                                
                                print(f"Colocación en Excel exitosa para primer texto")
                            else:
                                print(f"ERROR: No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"ERROR al procesar primer texto en PRELOSA ALIGERADA 20 '{texto}': {e}")
                                        
                    # Procesar segundo texto (G5, H5, J5) si existe
                    if len(textos_a_procesar) >= 2:
                        texto = textos_a_procesar[1]
                        print(f"Procesando segundo texto: '{texto}'")
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            print(f"Cantidad extraída: {cantidad}")
                            
                            # Extraer diámetro del texto con manejo mejorado para mm
                            if "mm" in texto:
                                # Caso específico para milímetros
                                diametro_match = re.search(r'∅(\d+)mm', texto)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    diametro_con_comillas = f"{diametro}mm"
                            else:
                                # Caso para fraccionales con o sin comillas
                                diametro_match = re.search(r'∅([\d/]+)', texto)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    # Asegurarnos de añadir comillas si es necesario
                                    if "\"" not in diametro and "/" in diametro:
                                        diametro_con_comillas = f"{diametro}\""
                                    else:
                                        diametro_con_comillas = diametro
                            
                            if diametro_match:
                                print(f"Diámetro extraído: {diametro_con_comillas}")
                                
                                cantidad = int(cantidad)  # Convertir a entero
                                
                                # NUEVO: Verificar si el texto incluye información de espaciamiento
                                espaciamiento_match = re.search(r'@\.?(\d+)', texto)
                                if espaciamiento_match:
                                    # Si encuentra espaciamiento en el texto, lo usa
                                    separacion = espaciamiento_match.group(1)
                                    separacion_decimal = float(f"0.{separacion}")
                                    print(f"Espaciamiento encontrado en texto: @{separacion} -> {separacion_decimal}")
                                else:
                                    # Si no encuentra espaciamiento, usa el valor predeterminado
                                    separacion_decimal = espaciamiento_aligerada
                                    print(f"No se encontró espaciamiento en texto, usando valor predeterminado: {espaciamiento_aligerada}")
                                
                                # Escribir en Excel
                                print(f"Escribiendo en Excel: G5={cantidad}, H5={diametro_con_comillas}, J5={separacion_decimal}")
                                ws.range('G5').value = cantidad
                                ws.range('H5').value = diametro_con_comillas
                                ws.range('J5').value = separacion_decimal
                                
                                print(f"Colocación en Excel exitosa para segundo texto")
                            else:
                                print(f"ERROR: No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"ERROR al procesar segundo texto en PRELOSA ALIGERADA 20 '{texto}': {e}")
                else:
                    print("ADVERTENCIA: No se encontraron textos (ni verticales ni horizontales) para PRELOSA ALIGERADA 20")
                
                # Verificamos antes de recalcular los valores actuales
                print("VALORES ANTES DE RECALCULAR:")
                print(f"  Celda K8 = {ws.range('K8').value}")
                print(f"  Celda K17 = {ws.range('K17').value}")
                
                # IMPORTANTE: Forzar recálculo y GUARDAR los valores calculados antes de cualquier limpieza
                print("Forzando recálculo de Excel...")
                ws.book.app.calculate()
                time.sleep(0.1)
                
                # GUARDAR los valores calculados en variables locales
                k8_valor = ws.range('K8').value
                k17_valor = ws.range('K17').value
                k18_valor = ws.range('K18').value
                
                print("VALORES FINALES CALCULADOS POR EXCEL (GUARDADOS):")
                print(f"  Celda K8 = {k8_valor}")
                print(f"  Celda K17 = {k17_valor}")
                print(f"  Celda K18 = {k18_valor}")
                
                # MODIFICAR las variables globales as_long, as_tra1, as_tra2 para que usen los valores guardados
                # Formatear correctamente el valor K8 (convertir de decimal a entero si es posible)
                if k8_valor is not None:
                    # Verificar si es un número
                    try:
                        # Multiplicar por 100 y redondear para evitar errores de punto flotante
                        k8_int = round(float(k8_valor) * 100)
                        # Convertir a entero si es posible (si termina en .0)
                        if k8_int % 100 == 0:  # Si es un número entero
                            k8_formateado = str(k8_int // 100)
                        else:  # Si tiene decimales
                            k8_formateado = str(k8_int / 100).rstrip('0').rstrip('.')
                    except (ValueError, TypeError):
                        k8_formateado = str(k8_valor)
                else:
                    k8_formateado = "20"  # Valor por defecto si K8 es None
                
                # Crear las cadenas finales para el bloque
                as_long = f"1Ø3/8\"@.{k8_formateado}"
                as_tra1 = "1Ø6 mm@.50"  # Valor fijo para PRELOSA ALIGERADA 20
                as_tra2 = "1Ø8 mm@.50"  # Valor fijo para PRELOSA ALIGERADA 20
                
                # Guardar estos valores finales en variables globales que no se pueden modificar
                # Esta es la parte crítica - asegurar que estos valores no cambien después
                global as_long_final, as_tra1_final, as_tra2_final
                as_long_final = as_long
                as_tra1_final = as_tra1
                as_tra2_final = as_tra2
                
                # Para seguridad, volvemos a imprimir los valores que se usarán
                print("VALORES FINALES QUE SE USARÁN PARA EL BLOQUE (NO SE MODIFICARÁN):")
                print(f"  AS_LONG: {as_long_final}")
                print(f"  AS_TRA1: {as_tra1_final} (valor fijo)")
                print(f"  AS_TRA2: {as_tra2_final} (valor fijo)")
                
                print("----------------------------------------")
                print("PROCESAMIENTO DE PRELOSA ALIGERADA 20 FINALIZADO")
                print("----------------------------------------")
                
                # IMPORTANTE: Asegurarnos de que estos valores se usen para el bloque
                as_long = as_long_final
                as_tra1 = as_tra1_final 
                as_tra2 = as_tra2_final

                        # Casos especiales para PRELOSA ALIGERADA 20 - 2 SENT
            elif tipo_prelosa == "PRELOSA ALIGERADA 20 - 2 SENT":
                # Usar el espaciamiento de los valores predeterminados (viene del tkinter)
                dist_aligerada2sent = float(default_valores.get('PRELOSA ALIGERADA 20 - 2 SENT', {}).get('espaciamiento', 0.605))
                print(f"Usando espaciamiento predeterminado para PRELOSA ALIGERADA 20 - 2 SENT: {dist_aligerada2sent}")
                
                # Caso 1: Si tenemos textos horizontales
                if len(textos_longitudinal) > 0:
                    print(f"Procesando {len(textos_longitudinal)} textos horizontales en PRELOSA ALIGERADA 20 - 2 SENT")
                    
                    # Procesar primer texto horizontal (G4, H4, J4)
                    if len(textos_longitudinal) >= 1:
                        texto = textos_longitudinal[0]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto (ej: "1∅3/8"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario (para 3/8")
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = int(cantidad)  # Convertir a entero
                                # Usar el valor predeterminado para el espaciamiento
                                separacion_decimal = dist_aligerada2sent
                                
                                # Escribir en Excel
                                ws.range('G4').value = cantidad
                                ws.range('H4').value = diametro_con_comillas
                                ws.range('J4').value = separacion_decimal
                                
                                print(f"Colocando en el excel primer texto horizontal: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                                print(f"NOTA: Usando espaciamiento predeterminado porque no viene en el texto")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto horizontal en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                    
                    # Procesar segundo texto horizontal (G5, H5, J5) si existe
                    if len(textos_longitudinal) >= 2:
                        texto = textos_longitudinal[1]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto (ej: "1∅1/2"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = int(cantidad)  # Convertir a entero
                                # Usar el valor predeterminado para el espaciamiento
                                separacion_decimal = dist_aligerada2sent
                                
                                # Escribir en Excel
                                ws.range('G5').value = cantidad
                                ws.range('H5').value = diametro_con_comillas
                                ws.range('J5').value = separacion_decimal
                                
                                print(f"Colocando en el excel segundo texto horizontal: {cantidad} -> G5, {diametro_con_comillas} -> H5, {separacion_decimal} -> J5")
                                print(f"NOTA: Usando espaciamiento predeterminado porque no viene en el texto")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto horizontal en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                
                # Caso 2: Si tenemos textos verticales
                if len(textos_transversal) > 0:
                    print(f"Procesando {len(textos_transversal)} textos verticales en PRELOSA ALIGERADA 20 - 2 SENT")
                    
                    # Procesar primer texto vertical (G14, H14, J14)
                    if len(textos_transversal) >= 1:
                        texto = textos_transversal[0]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = int(cantidad)  # Convertir a entero
                                # Usar el valor predeterminado para el espaciamiento
                                separacion_decimal = dist_aligerada2sent
                                
                                # Escribir en Excel
                                ws.range('G14').value = cantidad
                                ws.range('H14').value = diametro_con_comillas
                                ws.range('J14').value = separacion_decimal
                                
                                print(f"Colocando en el excel primer texto vertical: {cantidad} -> G14, {diametro_con_comillas} -> H14, {separacion_decimal} -> J14")
                                print(f"NOTA: Usando espaciamiento predeterminado porque no viene en el texto")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto vertical '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto vertical en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                    
                    # Procesar segundo texto vertical (G15, H15, J15) si existe
                    if len(textos_transversal) >= 2:
                        texto = textos_transversal[1]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                cantidad = int(cantidad)  # Convertir a entero
                                # Usar el valor predeterminado para el espaciamiento
                                separacion_decimal = dist_aligerada2sent
                                
                                # Escribir en Excel
                                ws.range('G15').value = cantidad
                                ws.range('H15').value = diametro_con_comillas
                                ws.range('J15').value = separacion_decimal
                                
                                print(f"Colocando en el excel segundo texto vertical: {cantidad} -> G15, {diametro_con_comillas} -> H15, {separacion_decimal} -> J15")
                                print(f"NOTA: Usando espaciamiento predeterminado porque no viene en el texto")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto vertical '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto vertical en PRELOSA ALIGERADA 20 - 2 SENT '{texto}': {e}")
                
                # IMPORTANTE: Forzar recálculo y GUARDAR los valores calculados antes de cualquier limpieza
                print("Forzando recálculo de Excel...")
                ws.book.app.calculate()
                time.sleep(0.1)
                
                # GUARDAR los valores calculados en variables locales
                valores_calculados = {
                    "k8": ws.range('K8').value,
                    "k17": ws.range('K17').value,
                    "k18": ws.range('K18').value
                }
                
                print("VALORES FINALES CALCULADOS POR EXCEL (GUARDADOS):")
                print(f"  Celda K8 = {valores_calculados['k8']}")
                print(f"  Celda K17 = {valores_calculados['k17']}")
                print(f"  Celda K18 = {valores_calculados['k18']}")
                
                # AHORA limpiar las celdas (esto no afectará los valores ya guardados)
           
                
                # MODIFICAR las variables globales as_long, as_tra1, as_tra2 para que usen los valores guardados
                # Esto es para que el código posterior use estos valores, no los recalculados
                as_long = valores_calculados["k8"]
                as_tra1 = 0.28  # Valor fijo para PRELOSA ALIGERADA 20 - 2 SENT
                as_tra2 = valores_calculados["k18"]
                
                # Para seguridad, volvemos a imprimir los valores que se usarán
                print("VALORES QUE SE USARÁN PARA EL BLOQUE (NO SE RECALCULARÁN):")
                print(f"  AS_LONG: {as_long}")
                print(f"  AS_TRA1: {as_tra1} (valor fijo)")
                print(f"  AS_TRA2: {as_tra2}")
            elif tipo_prelosa == "PRELOSA MACIZA":
                # Procesar textos horizontales
                if len(textos_longitudinal) > 0:
                    print(f"Procesando {len(textos_longitudinal)} textos longitudinal en PRELOSA MACIZA")
                    
                    # Procesar primer texto horizontal (G4, H4, J4)
                    if len(textos_longitudinal) >= 1:
                        texto = textos_longitudinal[0]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto (ej: "3/8"")
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario (para 3/8")
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                # Extraer espaciamiento del texto
                                espaciamiento_match = re.search(r'@(\d+)', texto)
                                if espaciamiento_match:
                                    separacion = int(espaciamiento_match.group(1))
                                    # Convertir a metros (dividir por 100)
                                    separacion_decimal = separacion / 100
                                else:
                                    print(f"ADVERTENCIA: No se encontró espaciamiento en el texto '{texto}'. No se usará valor por defecto.")
                                    separacion_decimal = None
                                
                                # Solo escribir en Excel si tenemos todos los datos
                                if separacion_decimal is not None:
                                    ws.range('G4').value = int(cantidad)
                                    ws.range('H4').value = diametro_con_comillas
                                    ws.range('J4').value = separacion_decimal
                                    
                                    print(f"Colocando en el excel primer texto horizontal: {cantidad} -> G4, {diametro_con_comillas} -> H4, {separacion_decimal} -> J4")
                                else:
                                    print(f"No se pudo extraer la separación del texto '{texto}', no se escribirá en Excel")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto horizontal en PRELOSA MACIZA '{texto}': {e}")
                    
                    # Procesar segundo texto horizontal (G5, H5, J5) si existe
                    if len(textos_longitudinal) >= 2:
                        texto = textos_longitudinal[1]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                # Extraer espaciamiento del texto
                                espaciamiento_match = re.search(r'@(\d+)', texto)
                                if espaciamiento_match:
                                    separacion = int(espaciamiento_match.group(1))
                                    # Convertir a metros (dividir por 100)
                                    separacion_decimal = separacion / 100
                                else:
                                    print(f"ADVERTENCIA: No se encontró espaciamiento en el texto '{texto}'. No se usará valor por defecto.")
                                    separacion_decimal = None
                                
                                # Solo escribir en Excel si tenemos todos los datos
                                if separacion_decimal is not None:
                                    ws.range('G5').value = int(cantidad)
                                    ws.range('H5').value = diametro_con_comillas
                                    ws.range('J5').value = separacion_decimal
                                    
                                    print(f"Colocando en el excel segundo texto horizontal: {cantidad} -> G5, {diametro_con_comillas} -> H5, {separacion_decimal} -> J5")
                                else:
                                    print(f"No se pudo extraer la separación del texto '{texto}', no se escribirá en Excel")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto horizontal en PRELOSA MACIZA '{texto}': {e}")
                
                # Procesar textos verticales
                if len(textos_transversal) > 0:
                    print(f"Procesando {len(textos_transversal)} textos verticales en PRELOSA MACIZA")
                    
                    # Procesar primer texto vertical (G14, H14, J14)
                    if len(textos_transversal) >= 1:
                        texto = textos_transversal[0]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                # Extraer espaciamiento del texto
                                espaciamiento_match = re.search(r'@(\d+)', texto)
                                if espaciamiento_match:
                                    separacion = int(espaciamiento_match.group(1))
                                    # Convertir a metros (dividir por 100)
                                    separacion_decimal = separacion / 100
                                else:
                                    print(f"ADVERTENCIA: No se encontró espaciamiento en el texto '{texto}'. No se usará valor por defecto.")
                                    separacion_decimal = None
                                
                                # Solo escribir en Excel si tenemos todos los datos
                                if separacion_decimal is not None:
                                    ws.range('G14').value = int(cantidad)
                                    ws.range('H14').value = diametro_con_comillas
                                    ws.range('J14').value = separacion_decimal
                                    
                                    print(f"Colocando en el excel primer texto vertical: {cantidad} -> G14, {diametro_con_comillas} -> H14, {separacion_decimal} -> J14")
                                else:
                                    print(f"No se pudo extraer la separación del texto '{texto}', no se escribirá en Excel")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto vertical '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar primer texto vertical en PRELOSA MACIZA '{texto}': {e}")
                    
                    # Procesar segundo texto vertical (G15, H15, J15) si existe
                    if len(textos_transversal) >= 2:
                        texto = textos_transversal[1]
                        try:
                            # Extraer cantidad (número antes de ∅)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                            else:
                                cantidad = "1"  # Si no hay número antes de ∅, la cantidad es 1
                            
                            # Extraer diámetro del texto
                            diametro_match = re.search(r'∅([\d/]+)', texto)
                            if diametro_match:
                                diametro = diametro_match.group(1)
                                # Asegurarnos de añadir comillas si es necesario
                                if "\"" not in diametro and "/" in diametro:
                                    diametro_con_comillas = f"{diametro}\""
                                else:
                                    diametro_con_comillas = diametro
                                
                                # Extraer espaciamiento del texto
                                espaciamiento_match = re.search(r'@(\d+)', texto)
                                if espaciamiento_match:
                                    separacion = int(espaciamiento_match.group(1))
                                    # Convertir a metros (dividir por 100)
                                    separacion_decimal = separacion / 100
                                else:
                                    print(f"ADVERTENCIA: No se encontró espaciamiento en el texto '{texto}'. No se usará valor por defecto.")
                                    separacion_decimal = None
                                
                                # Solo escribir en Excel si tenemos todos los datos
                                if separacion_decimal is not None:
                                    ws.range('G15').value = int(cantidad)
                                    ws.range('H15').value = diametro_con_comillas
                                    ws.range('J15').value = separacion_decimal
                                    
                                    print(f"Colocando en el excel segundo texto vertical: {cantidad} -> G15, {diametro_con_comillas} -> H15, {separacion_decimal} -> J15")
                                else:
                                    print(f"No se pudo extraer la separación del texto '{texto}', no se escribirá en Excel")
                            else:
                                print(f"No se pudo extraer información del diámetro en el texto vertical '{texto}'")
                        except Exception as e:
                            print(f"Error al procesar segundo texto vertical en PRELOSA MACIZA '{texto}': {e}")
                

                if len(textos_long_adi) > 0:
                    print("=" * 60)
                    print(f"PROCESANDO {len(textos_long_adi)} TEXTOS LONG ADI EN PRELOSA MACIZA ESPECIAL")
                    print("=" * 60)
                    
                    # Obtener el espaciamiento por defecto de los valores de tkinter
                    espaciamiento_macizas_adi = float(default_valores.get('PRELOSA MACIZA', {}).get('espaciamiento', 0.20))
                    print(f"Espaciamiento predeterminado para PRELOSA MACIZA ADI: {espaciamiento_macizas_adi}")

                    # Analizar el primer texto para ver si contiene información de espaciamiento
                    espaciamiento_primera_fila = espaciamiento_macizas_adi  # Valor por defecto
                    print(f"Inicializando espaciamiento primera fila con valor predeterminado: {espaciamiento_primera_fila}")

                    if len(textos_long_adi) > 0:
                        primer_texto = textos_long_adi[0]
                        print(f"Analizando primer texto para extraer espaciamiento: '{primer_texto}'")
                        
                        # Extraer espaciamiento del primer texto si existe
                        espaciamiento_match = re.search(r'@\.?(\d+)', texto)
                        if espaciamiento_match:
                            separacion = int(espaciamiento_match.group(1))
                            # Convertir a metros (dividir por 100)
                            separacion_decimal = separacion / 100
                        else:
                            print(f"ADVERTENCIA: No se encontró espaciamiento en el texto '{texto}'. No se usará valor por defecto.")
                            separacion_decimal = None

                    print("\nCOLOCANDO VALORES EN PRIMERA FILA (G4, H4, J4):")
                    print(f"  - Celda G4 = 1")
                    print(f"  - Celda H4 = 3/8\"")
                    print(f"  - Celda J4 = {espaciamiento_macizas_adi}")
                    
                    # Colocar los valores en la primera fila
                    ws.range('G4').value = 1
                    ws.range('H4').value = "3/8\""
                    ws.range('J4').value = espaciamiento_macizas_adi
                    
                    # NUEVO: Colocar los mismos valores en G14, H14, J14
                    print("\nCOLOCANDO VALORES EN PRIMERA FILA VERTICAL (G14, H14, J14):")
                    print(f"  - Celda G14 = 1")
                    print(f"  - Celda H14 = 3/8\"")
                    print(f"  - Celda J14 = {espaciamiento_macizas_adi}")
                    
                    ws.range('G14').value = 1
                    ws.range('H14').value = "3/8\""
                    ws.range('J14').value = espaciamiento_macizas_adi
                    
                    print(f"Valores default colocados exitosamente en filas 4 y 14")
                    
                    # Procesar los textos de acero long adi
                    datos_textos = []
                    print("\nPROCESANDO TEXTOS INDIVIDUALES DE ACERO LONG ADI:")
                    
                    for i, texto in enumerate(textos_long_adi):
                        print("-" * 50)
                        print(f"TEXTO #{i+1}: '{texto}'")
                        try:
                            # Extraer cantidad (número antes de ∅, si no hay se asume 1)
                            cantidad_match = re.search(r'^(\d+)∅', texto)
                            if cantidad_match:
                                cantidad = cantidad_match.group(1)
                                print(f"  ✓ Cantidad extraída del texto: {cantidad}")
                            else:
                                cantidad = "1"
                                print(f"  ✓ No se encontró cantidad explícita, asumiendo: {cantidad}")
                            
                            # Extraer diámetro del texto con manejo especial para mm
                            diametro_con_comillas = None
                            
                            # Primero intentar detectar formato específico mm
                            if "mm" in texto:
                                mm_match = re.search(r'∅(\d+)mm', texto)
                                if mm_match:
                                    diametro = mm_match.group(1)
                                    diametro_con_comillas = f"{diametro}mm"
                                    print(f"  ✓ Detectado diámetro en milímetros: {diametro_con_comillas}")
                            
                            # Si no se encontró formato mm específico, intentar formato general
                            if diametro_con_comillas is None:
                                diametro_match = re.search(r'∅([\d/]+)', texto)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    # Si es un número simple y el texto menciona mm, aplicar formato mm
                                    if "mm" in texto and "/" not in diametro:
                                        diametro_con_comillas = f"{diametro}mm"
                                        print(f"  ✓ Detectado número simple con indicación mm: {diametro_con_comillas}")
                                    # Si es fraccional, añadir comillas si es necesario
                                    elif "/" in diametro and "\"" not in diametro:
                                        diametro_con_comillas = f"{diametro}\""
                                        print(f"  ✓ Diámetro extraído: {diametro} -> añadiendo comillas: {diametro_con_comillas}")
                                    else:
                                        diametro_con_comillas = diametro
                                        print(f"  ✓ Diámetro extraído (ya con formato correcto): {diametro_con_comillas}")
                            
                            # Continuar solo si se extrajo un diámetro
                            if diametro_con_comillas:
                                # Extraer espaciamiento del texto
                                espaciamiento_match = re.search(r'@\.?(\d+)', texto)
                                if espaciamiento_match:
                                    separacion = espaciamiento_match.group(1)
                                    # Convertir a formato decimal (ej: 30 -> 0.30)
                                    separacion_decimal = float(f"0.{separacion}")
                                    print(f"  ✓ Espaciamiento extraído: @{separacion} -> {separacion_decimal}")
                                else:
                                    # Si no hay espaciamiento, usar el valor predeterminado
                                    separacion_decimal = float(espaciamiento_macizas_adi)
                                    print(f"  ✓ No se encontró espaciamiento en texto, usando valor predeterminado: {espaciamiento_macizas_adi}")
                                
                                # Guardar los datos procesados
                                datos_textos.append([int(cantidad), diametro_con_comillas, separacion_decimal])
                                print(f"  ✓ DATOS PROCESADOS Y GUARDADOS: cantidad={cantidad}, diámetro={diametro_con_comillas}, separación={separacion_decimal}")
                            else:
                                print(f"  ✗ ERROR: No se pudo extraer información del diámetro en el texto '{texto}'")
                        except Exception as e:
                            print(f"  ✗ ERROR al procesar texto: {e}")
                            traceback.print_exc()

                    print("\nCOLOCANDO VALORES EN FILAS ADICIONALES:")
                    # Colocar los valores extraídos en las filas adicionales (G5, H5, J5, etc.)
                    for i, datos in enumerate(datos_textos):
                        fila = 5 + i  # Comienza en fila 5
                        cantidad, diametro, separacion = datos
                        
                        print(f"FILA #{i+1} (G{fila}, H{fila}, J{fila}):")
                        print(f"  - Celda G{fila} = {cantidad}")
                        print(f"  - Celda H{fila} = {diametro}")
                        print(f"  - Celda J{fila} = {separacion}")
                        
                        ws.range(f'G{fila}').value = cantidad
                        ws.range(f'H{fila}').value = diametro
                        ws.range(f'J{fila}').value = separacion
                        
                        print(f"  ✓ Valores colocados exitosamente en fila {fila}")
                    
                    # NUEVO: Procesar aceros transversales adicionales
                    if len(textos_tra_adi) > 0:
                        print("\n" + "=" * 60)
                        print(f"PROCESANDO {len(textos_tra_adi)} TEXTOS TRANSVERSALES ADI EN PRELOSA MACIZA ESPECIAL")
                        print("=" * 60)
                        
                        # Procesar los textos de acero transversal adi
                        datos_textos_tra = []
                        print("\nPROCESANDO TEXTOS INDIVIDUALES DE ACERO TRANSVERSAL ADI:")
                        
                        for i, texto in enumerate(textos_tra_adi):
                            print("-" * 50)
                            print(f"TEXTO TRANSVERSAL #{i+1}: '{texto}'")
                            try:
                                # Extraer cantidad (número antes de ∅, si no hay se asume 1)
                                cantidad_match = re.search(r'^(\d+)∅', texto)
                                if cantidad_match:
                                    cantidad = cantidad_match.group(1)
                                    print(f"  ✓ Cantidad extraída del texto: {cantidad}")
                                else:
                                    cantidad = "1"
                                    print(f"  ✓ No se encontró cantidad explícita, asumiendo: {cantidad}")
                                
                                # Verificar si el texto tiene formato de milímetros
                                if "mm" in texto:
                                    # Caso específico para milímetros
                                    diametro_match = re.search(r'∅(\d+)mm', texto)
                                    if diametro_match:
                                        diametro = diametro_match.group(1)
                                        diametro_con_comillas = f"{diametro}mm"
                                        print(f"  ✓ Diámetro extraído (milímetros): {diametro_con_comillas}")
                                    else:
                                        # Si no pudo extraer con formato exacto, intentar el método genérico
                                        diametro_match = re.search(r'∅([\d/]+)', texto)
                                        if diametro_match:
                                            diametro = diametro_match.group(1)
                                            diametro_con_comillas = f"{diametro}mm"  # Forzar formato mm porque sabemos que está en texto
                                            print(f"  ✓ Diámetro extraído (milímetros, método alternativo): {diametro_con_comillas}")
                                        else:
                                            diametro_con_comillas = None
                                            print(f"  ✗ ERROR: No se pudo extraer diámetro del texto '{texto}'")
                                else:
                                    # Caso para fraccionales con o sin comillas
                                    diametro_match = re.search(r'∅([\d/]+)', texto)
                                    if diametro_match:
                                        diametro = diametro_match.group(1)
                                        # Asegurarnos de añadir comillas si es necesario
                                        if "\"" not in diametro and "/" in diametro:
                                            diametro_con_comillas = f"{diametro}\""
                                            print(f"  ✓ Diámetro extraído: {diametro} -> añadiendo comillas: {diametro_con_comillas}")
                                        else:
                                            diametro_con_comillas = diametro
                                            print(f"  ✓ Diámetro extraído (ya con formato correcto): {diametro_con_comillas}")
                                    else:
                                        diametro_con_comillas = None
                                        print(f"  ✗ ERROR: No se pudo extraer diámetro del texto '{texto}'")
                                
                                # Continuar solo si se extrajo un diámetro
                                if diametro_con_comillas:
                                    # Extraer espaciamiento del texto
                                    espaciamiento_match = re.search(r'@\.?(\d+)', texto)
                                    if espaciamiento_match:
                                        separacion = espaciamiento_match.group(1)
                                        # Convertir a formato decimal (ej: 30 -> 0.30)
                                        separacion_decimal = float(f"0.{separacion}")
                                        print(f"  ✓ Espaciamiento extraído: @{separacion} -> {separacion_decimal}")
                                    else:
                                        # Si no hay espaciamiento, usar el valor predeterminado
                                        separacion_decimal = float(espaciamiento_macizas_adi)
                                        print(f"  ✓ No se encontró espaciamiento en texto, usando valor predeterminado: {espaciamiento_macizas_adi}")
                                    
                                    # Guardar los datos procesados
                                    datos_textos_tra.append([int(cantidad), diametro_con_comillas, separacion_decimal])
                                    print(f"  ✓ DATOS PROCESADOS Y GUARDADOS: cantidad={cantidad}, diámetro={diametro_con_comillas}, separación={separacion_decimal}")
                                else:
                                    print(f"  ✗ ERROR: No se pudo procesar completo el texto '{texto}' por falta de diámetro")
                            except Exception as e:
                                print(f"  ✗ ERROR al procesar texto transversal: {e}")
                                traceback.print_exc()
                                                    
                        print("\nCOLOCANDO VALORES TRANSVERSALES EN FILAS ADICIONALES:")
                        # Colocar los valores extraídos en las filas adicionales (G15, H15, J15, etc.)
                        for i, datos in enumerate(datos_textos_tra):
                            fila = 15 + i  # Comienza en fila 15
                            cantidad, diametro, separacion = datos
                            
                            print(f"FILA TRANSVERSAL #{i+1} (G{fila}, H{fila}, J{fila}):")
                            print(f"  - Celda G{fila} = {cantidad}")
                            print(f"  - Celda H{fila} = {diametro}")
                            print(f"  - Celda J{fila} = {separacion}")
                            
                            ws.range(f'G{fila}').value = cantidad
                            ws.range(f'H{fila}').value = diametro
                            ws.range(f'J{fila}').value = separacion
                            
                            print(f"  ✓ Valores transversales colocados exitosamente en fila {fila}")
                    
                    # IMPORTANTE: Forzar recálculo y GUARDAR los valores calculados
                    print("\nFORZANDO RECÁLCULO DE EXCEL...")
                    ws.book.app.calculate()
                    time.sleep(0.1)
                    
                    # GUARDAR los valores calculados
                    k8_valor = ws.range('K8').value
                    k17_valor = ws.range('K17').value
                    k18_valor = ws.range('K18').value if ws.range('K18').value else None
                    
                    print(f"RESULTADOS CALCULADOS EN EXCEL:")
                    print(f"  ★ Celda K8 = {k8_valor}")
                    print(f"  ★ Celda K17 = {k17_valor}")
                    if k18_valor is not None:
                        print(f"  ★ Celda K18 = {k18_valor}")
                    
                    # Formatear valores para el bloque
                    k8_formateado = formatear_valor_espaciamiento(k8_valor)
                    as_long_texto = f"1Ø3/8\"@.{k8_formateado}"
                    as_tra1_texto = "1Ø6 mm@.28"  # Valor fijo para prelosas macizas
                    
                    # NUEVO: Formatear AS_TRA2 usando el valor de K18 si hay textos_tra_adi
                    if len(textos_tra_adi) > 0 and k18_valor is not None:
                        k18_formateado = formatear_valor_espaciamiento(k18_valor)
                        as_tra2_texto = f"1Ø8 mm@.{k18_formateado}"
                        print(f"  ★ Valor para AS_TRA2 calculado de K18: {as_tra2_texto}")
                    else:
                        as_tra2_texto = None
                    
                    print("\nVALORES FINALES FORMATEADOS PARA BLOQUE:")
                    print(f"  ★ AS_LONG: {as_long_texto} (de K8={k8_valor} -> formato .{k8_formateado})")
                    print(f"  ★ AS_TRA1: {as_tra1_texto} (valor fijo para prelosas macizas)")
                    if as_tra2_texto:
                        print(f"  ★ AS_TRA2: {as_tra2_texto} (de K18={k18_valor} -> formato .{formatear_valor_espaciamiento(k18_valor)})")
                    print("=" * 60)
#
                if len(textos_longitudinal) == 0 and len(textos_transversal) == 0 and len(textos_long_adi) == 0 and len(textos_tra_adi) == 0:
                    print("=" * 60)
                    print("PRELOSA MACIZA SIN NINGÚN TIPO DE ACERO DETECTADO")
                    print("=" * 60)
                    
                    # Obtener el espaciamiento por defecto de los valores de tkinter
                    espaciamiento_macizas_adi = float(default_valores.get('PRELOSA MACIZA', {}).get('espaciamiento', 0.20))
                    print(f"Usando espaciamiento predeterminado del tkinter: {espaciamiento_macizas_adi}")
                    
                    print("\nCOLOCANDO VALORES POR DEFECTO EN EXCEL:")
                    print(f"  - Celda G4 = 1")
                    print(f"  - Celda H4 = 3/8\"")
                    print(f"  - Celda J4 = {espaciamiento_macizas_adi}")
                    
                    # Colocar valores por defecto en Excel
                    ws.range('G4').value = 1
                    ws.range('H4').value = "3/8\""
                    ws.range('J4').value = espaciamiento_macizas_adi
                    
                    # Limpiar otras celdas para evitar interferencias
                    ws.range('G5').value = 0

                    ws.range('G15').value = 0
                    
                    # Forzar recálculo
                    print("\nFORZANDO RECÁLCULO DE EXCEL...")
                    ws.book.app.calculate()
                    time.sleep(0.1)
                    
                    # Guardar resultado para el bloque
                    k8_valor = ws.range('K8').value
                    print(f"RESULTADO CALCULADO EN EXCEL:")
                    print(f"  ★ Celda K8 = {k8_valor}")
                    
                    # Formatear para bloque
                    as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                    as_tra1_texto = "1Ø6 mm@.28"  # Valor fijo para prelosas macizas
                    as_tra2_texto = None
                    
                    print("\nVALORES FINALES PARA BLOQUE (CASO SIN ACEROS):")
                    print(f"  ★ AS_LONG: {as_long_texto}")
                    print(f"  ★ AS_TRA1: {as_tra1_texto}")
                    print("=" * 60)
             

                # Después de actualizar los valores en J, forzar recálculo


            else:
                # Procesar acero horizontal (G4, H4, J4)
                for i, texto in textos_longitudinal:
                    try:
                        # Extraer información del texto
                        match = re.match(r'∅(\d+\/\d+)\"@(\d+)', texto)
                        if match:
                            diametro = match.group(1)
                            separacion = match.group(2)
                            cantidad = "1"  # Por defecto
                            
                            # Añadir comillas al diámetro si no las tiene
                            diametro_con_comillas = f"{diametro}\""
                            
                            # Convertir separación a decimal directamente del texto
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
                
                # Procesar acero vertical (G14, H14, J14)
                for i, texto in textos_transversal:
                    try:
                        # Extraer información del texto
                        match = re.match(r'∅(\d+\/\d+)\"@(\d+)', texto)
                        if match:
                            diametro = match.group(1)
                            separacion = match.group(2)
                            cantidad = "1"  # Por defecto
                            
                            # Añadir comillas al diámetro si no las tiene
                            diametro_con_comillas = f"{diametro}\""
                            
                            # Convertir separación a decimal directamente del texto
                            separacion_decimal = float(f"0.{separacion}")
                            
                            # Escribir en Excel
                            ws.range('G14').value = int(cantidad)
                            ws.range('H14').value = diametro_con_comillas
                            ws.range('J14').value = separacion_decimal
                            
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
                                
                                print(f"Colocando en el excel: {cantidad} -> G14, {diametro_con_comillas} -> H14, {separacion_decimal} -> J14")
                            else:
                                print(f"No se pudo extraer información del texto '{texto}'")
                    except Exception as e:
                        print(f"Error al procesar texto vertical '{texto}': {e}")
            # Forzar cálculo y obtener resultados
            try:
                # Actualizar Excel de manera más agresiva
                
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
                if len(textos_transversal) > 0:
                    # Revisar si hay textos con "@20"
                    tiene_espaciamiento_20 = False
                    for texto in textos_transversal:
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
                
    
                
                # Verificar si hay textos verticales pero as_tra1 es 0 o nul
                
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
                for texto in textos_longitudinal:
                    match = re.search(r'@(\d+)', texto)
                    if match:
                        espaciamiento = match.group(1)
                        as_long = float(f"0.{espaciamiento}")
                        break
                
                for texto in textos_transversal:
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
                    # Verificar si tenemos aceros adicionales
    # Verificar si tenemos aceros adicionales o valores calculados manualmente
                    tiene_acero_adicional = len(textos_long_adi) > 0 or len(textos_tra_adi) > 0
                    tiene_valores_default = (len(textos_longitudinal) == 0 and len(textos_transversal) == 0 
                                        and len(textos_long_adi) == 0 and len(textos_tra_adi) == 0)
                    
                    if tiene_acero_adicional:
                        # Si hay aceros adicionales, usar esos valores que ya calculamos
                        print("PRELOSA MACIZA con ACEROS ADICIONALES - usando valores calculados previamente")
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                        as_tra1_texto = "1Ø6 mm@.28"  # Valor fijo para prelosas macizas
                        
                        # Para AS_TRA2 - usar el valor calculado de K18
                        if as_tra2 is not None:
                            as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(as_tra2)}"
                        else:
                            as_tra2_texto = None
                    elif tiene_valores_default:
                        # Si se usaron valores por defecto, no resetear, usar los valores ya calculados
                        print("PRELOSA MACIZA SIN ACEROS - usando valores calculados con valores por defecto")
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                        as_tra1_texto = "1Ø6 mm@.28"  # Valor fijo para prelosas macizas
                        
                        # Para AS_TRA2 - usar el valor calculado de K18
                        if as_tra2 is not None:
                            as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(as_tra2)}"
                        else:
                            as_tra2_texto = None
                    else:
                        # Procesamiento normal para aceros regulares
                        # Para acero horizontal (AS_LONG)
                        if len(textos_longitudinal) > 0:
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
                    if len(textos_longitudinal) > 0:
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    else:
                        # Si no hay textos horizontales, usar valor original
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_original)}"
                    
                    # Para acero vertical
                    if len(textos_transversal) > 0:
                        # Caso especial para espaciamiento 20 en vertical
                        for texto in textos_transversal:
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
                

                # Filtrar polilíneas longitudinales y adicionales
                polilineas_longitudinal = [p for p in polilineas_dentro if "LONGITUDINAL" in p.dxf.layer.upper() and "ADI" not in p.dxf.layer.upper()]
                polilineas_long_adi = [p for p in polilineas_dentro if "LONG ADI" in p.dxf.layer.upper()]

                # Calcular la orientación considerando ambos tipos de acero
                angulo_rotacion = calcular_orientacion_prelosa(vertices, polilineas_longitudinal, polilineas_long_adi)
                print(f"Orientando bloque a {angulo_rotacion:.2f}° (alineado con acero longitudinal)")
                
                # Crear una copia de la definición del bloque con la orientación calculada
                definicion_bloque_orientada = definicion_bloque.copy()
                definicion_bloque_orientada['rotation'] = angulo_rotacion
                
                # Insertar bloque con los valores formateados y la orientación correcta
                bloque = insertar_bloque_acero(msp, definicion_bloque_orientada, centro_prelosa, as_long_texto, as_tra1_texto, as_tra2_texto)
                
                if bloque:
                    total_bloques += 1
                    print(f"{tipo_prelosa} CONCLUIDA CON EXITO")
                    #limpiar celda g5
                    ws.range('G5').value = None
                    ws.range('G6').value = None
                    #limpiar celda g15
                    ws.range('G15').value = None
                    ws.range('G16').value = None
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
        
        capas_acero = ["ACERO LONGITUDINAL", "ACERO TRANSVERSAL", "ACERO LONG ADI",
        "BD-ACERO LONGITUDINAL", "BD-ACERO TRANSVERSAL", "ACERO TRA ADI",
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

    
# Punto de entrada principal del script
if __name__ == "__main__":

    file_path = "PLANO1.dxf"
    excel_path = "CONVERTIDOR.xlsx"
    
    # Generar nombre de archivo de salida con número aleatorio
    nombre_archivo_base = os.path.splitext(os.path.basename(file_path))[0]
    numero_random = random.randint(10, 99)  # Número aleatorio de 2 dígitos
    output_dxf_path = f"{nombre_archivo_base}_{numero_random}.dxf"
    
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