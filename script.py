
#ezdxf: sirve para leer y escribir archivos DXF
import ezdxf
#shapely: sirve para trabajar con geometría y realizar operaciones espaciales
from shapely.geometry import Point, Polygon, LineString  
#re: sirve para trabajar con expresiones regulares
import re
#os: sirve para interactuar con el sistema operativo
import os
#sys: proporciona acceso a variables y funciones que interactúan con el intérprete de Python
import sys
#traceback: permite extraer, formatear y imprimir información sobre excepciones
import xlwings as xw

import traceback
import time
import random



# Forzar la consola a aceptar UTF-8 en Windows
if os.name == 'nt':
    try:
        os.system('chcp 65001')
        if sys.stdout is not None and hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Función para reemplazar caracteres especiales
def reemplazar_caracteres_especiales(texto):
    texto = texto.replace("%%C", "∅")
    texto = texto.replace("\\A1;", "")  # Eliminar \A1; que aparece en algunos textos
    texto = re.sub(r'\\[A-Za-z0-9]+;', '', texto)
    return texto

def obtener_textos_dentro_de_polilinea(polilinea, textos, capa_polilinea=None):
    vertices = [(p[0], p[1]) for p in polilinea]
    poligono = Polygon(vertices)
    textos_en_polilinea = []
    
    # NUEVO: Lista para almacenar textos potencialmente fragmentados
    textos_fragmentados = []
    
    # Determinar orientación de la polilínea
    xs = [v[0] for v in vertices]
    ys = [v[1] for v in vertices]
    ancho = max(xs) - min(xs)
    alto = max(ys) - min(ys)
    es_vertical = alto > ancho
    
    # Determinar tipo de acero basado en la capa
    tipo_acero = "INDETERMINADO"
    if capa_polilinea:
        capa = capa_polilinea.upper()
        if "LONG ADI" in capa:
            tipo_acero = "LONG_ADI"
        elif "TRA ADI" in capa:
            tipo_acero = "TRA_ADI"
        elif "LONGITUDINAL" in capa:
            tipo_acero = "LONGITUDINAL"
        elif "TRANSVERSAL" in capa:
            tipo_acero = "TRANSVERSAL"

    # Notaciones alternativas de diámetros que son válidas
    conversion_diametros = {
        "M6": "6mm",
        "M8": "8mm",
        "#3": "3/8\"",
        "#4": "1/2\"",
        "#5": "5/8\""
    }

    def validar_formato_texto(texto):
        # Función existente sin cambios
        if texto is None:
            return ""
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
    
    # NUEVA FUNCIÓN: Verificar si el texto tiene un formato válido para especificación de acero
    def es_formato_acero_valido(texto):
        # Si está vacío, no es válido
        if not texto or texto.strip() == "":
            return False
        
        # Verificar notaciones alternativas directas (como M6, M8, #3, etc.)
        # Si el texto es exactamente una de estas notaciones, es válido
        for notacion in conversion_diametros.keys():
            if notacion in texto:
                return True
        
        # Patrones comunes para especificaciones de acero
        patrones = [
            # 1Ø3/8", 1Ø1/2", etc.
            r'\d+Ø[\d/]+"',
            # 1Ø8mm, 1Ø6mm, etc.
            r'\d+Ø\d+\s*mm',
            # 1∅3/8", 1∅1/2", etc.
            r'\d+∅[\d/]+"',
            # 1∅8mm, 1∅6mm, etc.
            r'\d+∅\d+\s*mm',
            # Formatos con espaciamiento: @.20, @20, etc.
            r'@[.,]?\d+',
            # Formatos como #3, #4, etc.
            r'#\d+',
            # Formatos como M6, M8, etc.
            r'M\d+',
            # Notación de diámetro simple: 3/8", 1/2", etc.
            r'[\d/]+"',
            # Notación de diámetro simple en mm: 6mm, 8mm, etc.
            r'\d+\s*mm',
            # Asegurar que contenga al menos un símbolo de diámetro (Ø o ∅)
            r'[Ø∅]'
        ]
        
        # Si el texto cumple con alguno de los patrones, es válido
        for patron in patrones:
            if re.search(patron, texto):
                return True
        
        # NUEVA COMPROBACIÓN: Si el texto contiene caracteres especiales como paréntesis
        # y no cumple con ningún patrón anterior, probablemente no sea una especificación de acero
        if any(c in texto for c in "(){}[]<>*"):
            return False
        
        # Verificar si contiene un número seguido inmediatamente de un Ø o ∅
        if re.search(r'\d+[Ø∅]', texto):
            return True
        
        # Por defecto, si no se detectaron patrones claros, ser conservador y rechazar el texto
        return False
    
    # NUEVO: Función para limpiar y normalizar el símbolo del diámetro en formato especial
    def limpiar_simbolo_diametro(texto):
        # Si el texto contiene un símbolo de diámetro con formato especial
        if re.search(r'\\f.*Symbol.*[Ø∅]', texto):
            # Intentar extraer solo el símbolo
            return "Ø"
        return texto
    
    # NUEVO: Función para identificar si un texto es un número, símbolo de diámetro o medida
    def clasificar_fragmento(texto):
        texto = texto.strip()
        # Es un formato especial de símbolo de diámetro
        if re.search(r'\\f.*Symbol.*[Ø∅]', texto):
            return "SIMBOLO_ESPECIAL"
        # Es un número solo
        elif re.match(r'^\d+$', texto):
            return "NUMERO"
        # Es un símbolo de diámetro solo
        elif texto in ["Ø", "∅"]:
            return "SIMBOLO"
        # Es una medida (con pulgadas o mm)
        elif re.match(r'^[\d/]+"', texto) or re.match(r'^\d+\s*mm', texto) or "(Inf.)" in texto:
            return "MEDIDA"
        # Es una notación completa con (Inf.)
        elif "Inf" in texto and ("3/8" in texto or "1/2" in texto):
            return "MEDIDA_COMPLETA"
        # No es ninguno de los anteriores
        return "OTRO"

    # Recorremos todos los elementos para recopilar los textos y sus posiciones
    textos_detallados = []
    
    # NUEVO: Procesar primero el caso especial del símbolo con formato especial
    tiene_simbolo_especial = False
    simbolo_especial_pos = None
    
    # Primera pasada: recopilar todos los textos y sus posiciones
    for elemento in textos:
        # Procesar textos normales (TEXT, MTEXT)
        if elemento.dxftype() in ['TEXT', 'MTEXT']:
            punto_texto = Point(elemento.dxf.insert)
            if poligono.contains(punto_texto):
                if elemento.dxftype() == 'MTEXT':
                    texto_contenido = elemento.text
                else:
                    texto_contenido = elemento.dxf.text
                
                # Aplicar validaciones de formato
                texto_formateado = validar_formato_texto(texto_contenido)
                
                # Identificar símbolos especiales
                tipo = clasificar_fragmento(texto_formateado)
                if tipo == "SIMBOLO_ESPECIAL":
                    tiene_simbolo_especial = True
                    simbolo_especial_pos = (elemento.dxf.insert[0], elemento.dxf.insert[1])
                
                # Almacenar el texto y su posición
                if texto_formateado.strip():
                    textos_detallados.append({
                        'texto': texto_formateado,
                        'x': elemento.dxf.insert[0],
                        'y': elemento.dxf.insert[1],
                        'tipo': tipo,
                        'procesado': False
                    })
    
    # NUEVO: Procesar específicamente el caso del símbolo especial
    textos_procesados = []
    
    # Si tenemos un símbolo especial, intentamos combinarlo con medidas cercanas
    if tiene_simbolo_especial:
        textos_medida = [t for t in textos_detallados if t['tipo'] == "MEDIDA" or t['tipo'] == "MEDIDA_COMPLETA"]
        simbolos_especiales = [t for t in textos_detallados if t['tipo'] == "SIMBOLO_ESPECIAL"]
        
        # NUEVA LÓGICA: Intentar crear una especificación para cada medida usando el símbolo especial
        for medida in textos_medida:
            # Por defecto, asumimos un "1" como cantidad
            cantidad = "1"
            
            # Buscar si hay un número cercano que podría ser la cantidad
            numeros_cercanos = [t for t in textos_detallados if t['tipo'] == "NUMERO"]
            if numeros_cercanos:
                # Usar el número más cercano horizontalmente a la medida
                numeros_cercanos.sort(key=lambda n: abs(n['x'] - medida['x']) + abs(n['y'] - medida['y']) * 2)
                cantidad = numeros_cercanos[0]['texto']
                numeros_cercanos[0]['procesado'] = True
            
            # Crear la especificación completa: cantidad + símbolo + medida
            especificacion = f"{cantidad}Ø{medida['texto']}"
            print(f"Formando especificación combinada: {especificacion}")
            textos_procesados.append(especificacion)
            medida['procesado'] = True
        
        # Marcar todos los símbolos especiales como procesados
        for simbolo in simbolos_especiales:
            simbolo['procesado'] = True
    
    # Procesar los textos que ya son válidos individualmente
    for texto in textos_detallados:
        if not texto['procesado'] and es_formato_acero_valido(texto['texto']):
            textos_procesados.append(texto['texto'])
            texto['procesado'] = True
    
    # NUEVO: Verificar si quedaron textos de medida sin procesar (sin combinar con símbolo)
    for texto in textos_detallados:
        if not texto['procesado'] and (texto['tipo'] == "MEDIDA" or texto['tipo'] == "MEDIDA_COMPLETA"):
            # Si es una medida válida por sí sola, la añadimos
            if es_formato_acero_valido(texto['texto']):
                textos_procesados.append(texto['texto'])
                texto['procesado'] = True
    
    # NUEVO: Si no se procesó ningún texto pero hay medidas y símbolos, forzar la combinación
    if not textos_procesados:
        medidas = [t for t in textos_detallados if t['tipo'] == "MEDIDA" or t['tipo'] == "MEDIDA_COMPLETA"]
        for medida in medidas:
            # Añadir la medida con un "1Ø" prefijado
            textos_procesados.append(f"1Ø{medida['texto']}")
            print(f"Forzando combinación para medida: 1Ø{medida['texto']}")
            medida['procesado'] = True
    
    # Ahora procesar normalmente para capturar los casos que no son textos fragmentados
    # (Bloques, atributos, etc.)
    
    # Recorremos todos los elementos
    for elemento in textos:
        # Procesar textos normales (TEXT, MTEXT) - ya los procesamos arriba
        if elemento.dxftype() in ['TEXT', 'MTEXT']:
            continue  # Ya procesado arriba
        
        # MEJORA PARA BLOQUES: Detectar bloques incluso si solo intersectan parcialmente
        elif elemento.dxftype() == 'INSERT':
            punto_bloque = Point(elemento.dxf.insert)
            
            # Criterios más flexibles para bloques:
            # 1. Si el punto de inserción está dentro
            # 2. Si está muy cerca de la polilínea (distancia < 3.0 unidades)
            # 3. Si el bloque es relevante para el tipo de acero (basado en el nombre del bloque)
            
            # Aumentamos la tolerancia para capturar bloques incluso si están parcialmente fuera
            tolerancia_bloques = 1.0  # Unidades del dibujo
            
            # Verificar si el bloque es relevante por su nombre (si disponible)
            es_bloque_acero = False
            if hasattr(elemento.dxf, 'name'):
                nombre_bloque = elemento.dxf.name.upper()
                es_bloque_acero = any(term in nombre_bloque for term in ['ACERO', 'REFUERZO', 'ARMADURA', 'BARRA'])
            
            # Aumentar tolerancia si parece ser un bloque de acero
            if es_bloque_acero:
                tolerancia_bloques = 1.5  # Mayor tolerancia para bloques específicos de acero
            
            # Verificar si el bloque intersecta con la polilínea
            if (poligono.contains(punto_bloque) or 
                poligono.distance(punto_bloque) < tolerancia_bloques):
                
                print(f"Analizando bloque en posición ({elemento.dxf.insert[0]}, {elemento.dxf.insert[1]})")
                
                # NUEVO: Intentar extraer texto directamente del contenido del bloque
                try:
                    # Si el bloque tiene una definición y contiene textos
                    if hasattr(elemento, 'block') and elemento.block is not None:
                        for entidad in elemento.block:
                            if entidad.dxftype() in ['TEXT', 'MTEXT']:
                                try:
                                    if entidad.dxftype() == 'MTEXT':
                                        texto_bloque = entidad.text
                                    else:
                                        texto_bloque = entidad.dxf.text
                                        
                                    texto_formateado = validar_formato_texto(texto_bloque)
                                    # NUEVO: Validar formato de acero
                                    if texto_formateado.strip() and es_formato_acero_valido(texto_formateado):
                                        print(f"Texto extraído del interior del bloque: {texto_formateado}")
                                        textos_procesados.append(texto_formateado)
                                    else:
                                        print(f"Texto de bloque descartado (formato no válido): {texto_formateado}")
                                except Exception as e:
                                    print(f"Error al extraer texto del bloque: {str(e)}")
                except Exception as e:
                    print(f"Error al examinar contenido del bloque: {str(e)}")
                
                # Procesamiento de atributos (sin cambios)
                atributos_encontrados = False
                
                # Método 1: Obtener atributos usando .attribs
                try:
                    if hasattr(elemento, 'attribs'):
                        for attrib in elemento.attribs:
                            if hasattr(attrib.dxf, 'tag') and hasattr(attrib.dxf, 'text'):
                                # Filtrar atributos relevantes
                                if attrib.dxf.tag in ['ACERO', 'AS_LONG', 'AS_TRA1', 'AS_TRA2']:
                                    texto_contenido = attrib.dxf.text
                                    texto_formateado = validar_formato_texto(texto_contenido)
                                    # NUEVO: Validar formato de acero
                                    if texto_formateado.strip() and es_formato_acero_valido(texto_formateado):
                                        print(f"Encontrado atributo en bloque: {attrib.dxf.tag} = {texto_formateado}")
                                        textos_procesados.append(texto_formateado)
                                        atributos_encontrados = True
                                    else:
                                        print(f"Atributo descartado (formato no válido): {texto_formateado}")
                except Exception as e:
                    print(f"Error al procesar atributos (método 1): {str(e)}")
                
                # Método 3: Usar get_attribs si está disponible
                try:
                    if hasattr(elemento, 'get_attribs'):
                        for attrib in elemento.get_attribs():
                            if hasattr(attrib.dxf, 'tag') and hasattr(attrib.dxf, 'text'):
                                if attrib.dxf.tag in ['ACERO', 'AS_LONG', 'AS_TRA1', 'AS_TRA2']:
                                    texto_contenido = attrib.dxf.text
                                    texto_formateado = validar_formato_texto(texto_contenido)
                                    # NUEVO: Validar formato de acero
                                    if texto_formateado.strip() and es_formato_acero_valido(texto_formateado):
                                        print(f"Encontrado atributo (método 3): {attrib.dxf.tag} = {texto_formateado}")
                                        textos_procesados.append(texto_formateado)
                                        atributos_encontrados = True
                                    else:
                                        print(f"Atributo descartado (método 3, formato no válido): {texto_formateado}")
                except Exception as e:
                    print(f"Error al procesar atributos (método 3): {str(e)}")
                
                if atributos_encontrados:
                    print(f"=> Se encontraron atributos en bloque en posición ({elemento.dxf.insert[0]}, {elemento.dxf.insert[1]})")
                else:
                    print(f"No se encontraron atributos relevantes en bloque en posición ({elemento.dxf.insert[0]}, {elemento.dxf.insert[1]})")

    # Eliminar posibles duplicados manteniendo el orden
    textos_en_polilinea = []
    for texto in textos_procesados:
        if texto not in textos_en_polilinea:
            textos_en_polilinea.append(texto)
    
    print(f"Total de {len(textos_en_polilinea)} textos/atributos encontrados dentro de la polilínea")
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
            
    
    print(f"=> Capas de polilíneas encontradas: {capas_encontradas}")

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
                    polilineas_dentro.append(polilinea)
        except Exception as e:
            print(f"Error al procesar polilínea: {e}")
    
    # Información de depuración adicional
    print(f"=> Total de polilíneas de acero encontradas dentro de la prelosa: {len(polilineas_dentro)}")
    
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
    
    print("=> No se encontró el bloque de acero. Se creará uno genérico.")
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

# Función para insertar un bloque de acero en el modelo
def insertar_bloque_acero(msp, definicion_bloque, centro, as_long, as_tra1, as_tra2=None):
    """
    Inserta un bloque de acero en el centro de la prelosa con los valores calculados,
    respetando la escala del bloque original que el usuario colocó en el plano.
    Maneja rotaciones en cualquier ángulo, incluyendo casos inclinados.
    """
    try:
        # Usar directamente los valores recibidos
        str_as_long = as_long
        str_as_tra1 = as_tra1
        str_as_tra2 = as_tra2 if as_tra2 is not None else ""
        
        # Verificar y desbloquear la capa si es necesario
        capa_destino = definicion_bloque.get('capa', 'BD-ACERO POSITIVO')
        
        # Verificar si la capa existe y desbloquearla
        doc = msp.doc
        if capa_destino in doc.layers:
            layer = doc.layers.get(capa_destino)
            if hasattr(layer.dxf, 'lock') and layer.dxf.lock:
                layer.dxf.lock = False

        # Verificar capa alternativa
        capa_alternativa = '- BD - ACERO POSITIVO'
        if capa_alternativa in doc.layers:
            layer = doc.layers.get(capa_alternativa)
            if hasattr(layer.dxf, 'lock') and layer.dxf.lock:
                layer.dxf.lock = False
            
            # Si la capa original no existe pero la alternativa sí, usar la alternativa
            if capa_destino not in doc.layers:
                capa_destino = capa_alternativa
        
        # Obtener la rotación original
        rotation = definicion_bloque.get('rotation', 0.0)
        print(f"    Rotación original del bloque: {rotation:.2f}°")
        
        # Normalizar a 0-360
        rotation = rotation % 360
        
        # LÓGICA CORREGIDA: Manejar las rotaciones para asegurar que el texto sea legible
        # Para casos horizontales (0° o 180°)
        if rotation == 0.0 or abs(rotation - 180.0) < 0.1:
            # Si el bloque está horizontal y a 180°, girar 180° para que el texto sea legible
            if abs(rotation - 180.0) < 0.1:
                rotation = 0.0
                print(f"    Corrigiendo rotación horizontal a: {rotation:.2f}°")
        
        # Para casos verticales (90° o 270°)
        elif abs(rotation - 90.0) < 0.1 or abs(rotation - 270.0) < 0.1:
            # Si el bloque está vertical y a 270°, ajustar a 90° para que el texto sea legible
            if abs(rotation - 270.0) < 0.1:
                rotation = 90.0
                print(f"    Corrigiendo rotación vertical a: {rotation:.2f}°")
        
        # Para casos inclinados (cualquier otro ángulo)
        else:
            # Determinar cuadrante y ajustar para legibilidad
            if 0 < rotation < 90:
                # Primer cuadrante - mantener
                pass
            elif 90 < rotation < 180:
                # Segundo cuadrante - rotar a primer cuadrante
                rotation = (rotation - 180) % 360
                print(f"    Rotación ajustada para ángulo inclinado 2do cuadrante: {rotation:.2f}°")
            elif 180 < rotation < 270:
                # Tercer cuadrante - rotar a cuarto cuadrante
                rotation = (rotation - 180) % 360
                print(f"    Rotación ajustada para ángulo inclinado 3er cuadrante: {rotation:.2f}°")
            elif 270 < rotation < 360:
                # Cuarto cuadrante - mantener
                pass
        
        # NUEVO: Usar las escalas del bloque original, o ajustar solo si necesario
        # Obtener escalas del bloque original
        xscale_original = definicion_bloque.get('xscale', 1.0)
        yscale_original = definicion_bloque.get('yscale', 1.0)
        
        # Obtener dimensiones de la polilínea si están disponibles
        ancho_polilinea = definicion_bloque.get('polilinea_ancho', 0)
        alto_polilinea = definicion_bloque.get('polilinea_alto', 0)
        
        # MODIFICADO: Factor de aumento para hacer el bloque un poco más grande
        factor_aumento = 1.25  # Aumenta el tamaño en un 25%
        
        # Escala base: usar la del bloque original con el factor de aumento
        xscale = xscale_original * factor_aumento
        yscale = yscale_original * factor_aumento
        
        # Si la polilínea es muy pequeña y las escalas originales son demasiado grandes, 
        # ajustar con un factor de reducción
        if ancho_polilinea > 0 and alto_polilinea > 0:
            # Determinar si necesitamos reducir la escala (bloque demasiado grande para la polilínea)
            area_polilinea = ancho_polilinea * alto_polilinea
            limite_area = 100  # Valor de ejemplo, ajustar según necesidad
            
            if area_polilinea < limite_area and (xscale > 0.5 or yscale > 0.5):
                factor_reduccion = min(1.0, area_polilinea / limite_area)
                xscale = xscale * factor_reduccion
                yscale = yscale * factor_reduccion
                print(f"    Reduciendo escala para polilínea pequeña. Factor: {factor_reduccion:.3f}")
        
        print(f"    Usando escala ajustada: X={xscale:.3f}, Y={yscale:.3f} (original * {factor_aumento})")
        
        # Insertar el bloque con la rotación ajustada y escala modificada
        bloque = msp.add_blockref(
            name=definicion_bloque['nombre'],
            insert=centro,
            dxfattribs={
                'layer': capa_destino,
                'xscale': xscale,
                'yscale': yscale,
                'rotation': rotation
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

# Función para desbloquear la capa de acero positivo
def desbloquear_capa_acero_positivo(doc):
    try:
        # Nombre exacto de la capa
        nombre_capa = " - BD - ACERO POSITIVO"
        
        # Obtener la capa
        capa = doc.layers.get(nombre_capa)
        
        # Desbloquear la capa
        if capa:
            capa.dxf.flags = 0  # Código para desbloquear
            print(f"[ÉXITO] Capa '{nombre_capa}' desbloqueada correctamente")
        else:
            print(f"[ADVERTENCIA] Capa '{nombre_capa}' no encontrada")
    
    except Exception as e:
        print(f"[ERROR] No se pudo desbloquear la capa: {e}")

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
def procesar_prelosas_con_bloques(file_path, excel_path, output_dxf_path, valores_predeterminados=None):
    """
    Procesa las prelosas identificando tipos y contenidos,
    calcula valores usando Excel y coloca bloques con los resultados.
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
        ║                                     por DODOD SOLUTIONS                                        ║
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
        'PRELOSA MACIZA 15': {
            'espaciamiento': '0.15'
        },
        'PRELOSA ALIGERADA 20': {
            'espaciamiento': '0.605'
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
            else:
                # Añadir nuevos tipos personalizados
                default_valores[tipo] = valores
                print(f"Añadido tipo personalizado: {tipo} con valores: {valores}")

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
        print("Limpiando...")
        try:
            
            # Limpiar segundas filas horizontales
            ws.range('G5').value = 0
            
            # Limpiar segundas filas verticales
            ws.range('G15').value = 0

            # Forzar cálculo para actualizar con estos valores por defecto
            wb.app.calculate()
            
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
            definicion_bloque = {
                'nombre': 'BD-ACERO PRELOSA',
                'capa': 'BD-ACERO POSITIVO',
                'xscale': 1.0,
                'yscale': 1.0,
                'rotation': 0.0
            }

        polilineas_por_tipo = {}
        
        # Primero obtener todas las polilíneas del modelo space
        todas_polilineas = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE']
        
        # Obtener todas las capas de tipos de prelosa definidas
        tipos_prelosa = list(default_valores.keys())
        
        print(f"Buscando polilíneas para los siguientes tipos: {tipos_prelosa}")
        
        # Asignar cada polilínea a su respectivo tipo según la capa
        for entity in todas_polilineas:
            capa = entity.dxf.layer
            # Si la capa existe como clave en default_valores, es un tipo válido
            if capa in tipos_prelosa:
                if capa not in polilineas_por_tipo:
                    polilineas_por_tipo[capa] = []
                polilineas_por_tipo[capa].append(entity)
            
        # Definir la función para clasificar tipos de prelosa
        def clasificar_tipo_prelosa(tipo):
            """Determina la categoría base de la prelosa según su nombre"""
            if "MACIZA" in tipo.upper():
                return "MACIZA"
            elif "ALIGERADA" in tipo.upper() and "2 SENT" in tipo.upper():
                return "ALIGERADA_2SENT"
            elif "ALIGERADA" in tipo.upper():
                return "ALIGERADA"
            else:
                return "MACIZA"
                

        
        # Obtener polilíneas y textos
        polilineas_macizas = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA MACIZA"]
        polilineas_macizas_15 = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA MACIZA 15"]
        polilineas_aligeradas = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 20"]
        polilineas_aligeradas_2sent = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 20 - 2 SENT"]
        polilineas_aligeradas_25 = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 25"]
        polilineas_aligeradas_25_2sent = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 25 - 2 SENT"]
        polilinea_macizas_tipo3 = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA MACIZA TIPO 3"]
        polilinea_macizas_tipo4 = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA MACIZA TIPO 4"]
        polilineas_aligeradas_30 = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 30"]
        polilineas_aligeradas_30_2sent = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == "PRELOSA ALIGERADA 30 - 2 SENT"]
        polilineas_acero = [entity for entity in msp if entity.dxftype() == 'LWPOLYLINE' and 
                 entity.dxf.layer in ["ACERO LONGITUDINAL", "ACERO TRANSVERSAL", "ACERO LONG ADI", "ACERO TRA ADI",
                                      "BD-ACERO LONGITUDINAL", "BD-ACERO TRANSVERSAL",
                                      "ACERO", "REFUERZO", "ARMADURA"]]
        textos = [entity for entity in msp if entity.dxftype() in ['TEXT', 'MTEXT', 'INSERT', 'MULTILEADER']]
        
        # Contadores para estadísticas
        total_prelosas = 0
        total_bloques = 0
        
        # Print all layer names in the DXF file
        def clasificar_tipo_prelosa(tipo):
            
            tipo_upper = tipo.upper()
            
            # Primero verificar el caso especial de ALI-MAC
            if "ALI-MAC" in tipo_upper:
                return "ALIGERADA_2SENT"  # Las prelosas ALI-MAC se comportan como aligeradas 2 sent
            
            # Luego verificar los casos normales
            if "MACIZA" in tipo_upper:
                return "MACIZA"
            elif "ALIGERADA" in tipo_upper and "2 SENT" in tipo_upper:
                return "ALIGERADA_2SENT"
            elif "ALIGERADA" in tipo_upper:
                return "ALIGERADA"
            else:
                # Si no se reconoce, usar maciza por defecto
                print(f"ADVERTENCIA: Tipo de prelosa no reconocido: {tipo}. Usando MACIZA por defecto.")
                return "MACIZA"

        def calcular_orientacion_prelosa(vertices, polilineas_longitudinal=None, polilineas_long_adi=None):
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
                        # NUEVO: Detectar si la polilínea está inclinada
                        vertices_long_array = np.array(vertices_long)
                        
                        # Calcular el rango en X e Y para determinar la orientación espacial
                        min_x = np.min(vertices_long_array[:, 0])
                        max_x = np.max(vertices_long_array[:, 0])
                        min_y = np.min(vertices_long_array[:, 1])
                        max_y = np.max(vertices_long_array[:, 1])
                        
                        rango_x = max_x - min_x
                        rango_y = max_y - min_y
                        
                        # Si la diferencia entre los rangos es pequeña, puede estar inclinada
                        es_inclinada = abs(rango_x - rango_y) < 0.5 * max(rango_x, rango_y)
                        
                        # Si la polilínea parece estar inclinada, calcular ángulo exacto
                        if es_inclinada:
                            # Calcular el ángulo usando el primer y último punto para mejor dirección
                            p1 = vertices_long[0]
                            p2 = vertices_long[-1]
                            
                            # Calcular vector dirección y su ángulo
                            dx = p2[0] - p1[0]
                            dy = p2[1] - p1[1]
                            angulo = math.degrees(math.atan2(dy, dx))
                            
                            print(f"Polilínea LONGITUDINAL inclinada. Ángulo exacto: {angulo:.2f}°")
                            return angulo
                        else:
                            # Usar la lógica original para polilíneas no inclinadas
                            if rango_y > rango_x:  # Orientación predominantemente vertical
                                angulo_final = 90.0
                                print(f"Polilínea LONGITUDINAL orientada verticalmente. Ángulo: {angulo_final}°")
                            else:  # Orientación predominantemente horizontal
                                angulo_final = 0.0
                                print(f"Polilínea LONGITUDINAL orientada horizontalmente. Ángulo: {angulo_final}°")
                            
                            return angulo_final
                
                # 2. Si no hay ACERO LONGITUDINAL, intentar con ACERO LONG ADI
                if polilineas_long_adi and len(polilineas_long_adi) > 0:
                    print("No se encontró ACERO LONGITUDINAL. Usando orientación de ACERO LONG ADI para el bloque")
                    
                    # Obtener la primera polilínea de acero adicional
                    polilinea_long_adi = polilineas_long_adi[0]
                    vertices_long_adi = polilinea_long_adi.get_points('xy')
                    
                    # Necesitamos al menos 2 puntos para determinar una dirección
                    if len(vertices_long_adi) >= 2:
                        # NUEVO: Detectar si la polilínea está inclinada
                        vertices_long_adi_array = np.array(vertices_long_adi)
                        
                        # Calcular el rango en X e Y para determinar la orientación espacial
                        min_x = np.min(vertices_long_adi_array[:, 0])
                        max_x = np.max(vertices_long_adi_array[:, 0])
                        min_y = np.min(vertices_long_adi_array[:, 1])
                        max_y = np.max(vertices_long_adi_array[:, 1])
                        
                        rango_x = max_x - min_x
                        rango_y = max_y - min_y
                        
                        # Si la diferencia entre los rangos es pequeña, puede estar inclinada
                        es_inclinada = abs(rango_x - rango_y) < 0.5 * max(rango_x, rango_y)
                        
                        # Si la polilínea parece estar inclinada, calcular ángulo exacto
                        if es_inclinada:
                            # Calcular el ángulo usando el primer y último punto para mejor dirección
                            p1 = vertices_long_adi[0]
                            p2 = vertices_long_adi[-1]
                            
                            # Calcular vector dirección y su ángulo
                            dx = p2[0] - p1[0]
                            dy = p2[1] - p1[1]
                            angulo = math.degrees(math.atan2(dy, dx))
                            
                            print(f"Polilínea LONG ADI inclinada. Ángulo exacto: {angulo:.2f}°")
                            return angulo
                        else:
                            # Usar la lógica original para polilíneas no inclinadas
                            if rango_y > rango_x:  # Orientación predominantemente vertical
                                angulo_final = 90.0
                                print(f"Polilínea LONG ADI orientada verticalmente. Ángulo: {angulo_final}°")
                            else:  # Orientación predominantemente horizontal
                                angulo_final = 0.0
                                print(f"Polilínea LONG ADI orientada horizontalmente. Ángulo: {angulo_final}°")
                            
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
            
            # Variables para almacenar textos por tipo de acero
            textos_longitudinal = []
            textos_transversal = []
            textos_adicionales = []
            textos_long_adi = []
            textos_tra_adi = []
            
            # Procesar polilíneas de acero
            for polilinea_anidada in polilineas_dentro:
                vertices_anidada = polilinea_anidada.get_points('xy')
                
                textos_dentro = obtener_textos_dentro_de_polilinea(
                    vertices_anidada,
                    textos,
                    capa_polilinea=polilinea_anidada.dxf.layer
                )
                                
                print(f"Polilínea anidada en {tipo_prelosa.lower()} {idx+1} tiene {len(textos_dentro)} textos dentro.")
                
                # Clasificar textos según el tipo de acero
                tipo_acero = polilinea_anidada.dxf.layer.upper()
                if "LONGITUDINAL" in tipo_acero:
                    for texto in textos_dentro:
                        print("=" * 50)
                        print(f"Texto encontrado en ACERO LONGITUDINAL: {texto}")
                        textos_longitudinal.append(texto)
                elif "TRANSVERSAL" in tipo_acero:
                    for texto in textos_dentro:
                        print("=" * 50)
                        print(f"Texto encontrado en ACERO TRANSVERSAL: {texto}")
                        textos_transversal.append(texto)
                elif "ACERO LONG ADI" in tipo_acero:
                    for texto in textos_dentro:
                        print("=" * 50)
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
            
            # No limpiar celdas, solo sobrescribir
            # Almacenar los valores originales de K8 y K17 antes de cualquier modificación 
            k8_actual = ws.range('K8').value
            k17_actual = ws.range('K17').value
                        
            # Casos especiales para PRELOSA ALIGERADA 20
            categoria_base = clasificar_tipo_prelosa(tipo_prelosa)
            print(f"Procesando prelosa tipo '{tipo_prelosa}' (categoría: {categoria_base})")

            # Obtener valores específicos para este tipo
            espaciamiento_predeterminado = float(default_valores.get(tipo_prelosa, {}).get('espaciamiento', 0.20))
            acero_predeterminado = default_valores.get(tipo_prelosa, {}).get('acero', "3/8\"")
            print(f"Usando valores: espaciamiento={espaciamiento_predeterminado}, acero={acero_predeterminado}")

            # Aplicar lógica según la categoría
            if categoria_base == "MACIZA":
                print("\n=== PROCESANDO PRELOSA MACIZA ===")
                
                # Diccionario de conversión para notaciones alternativas de diámetros
                conversion_diametros = {
                    "M6": "6mm",
                    "M8": "8mm",
                    "#3": "3/8\"",
                    "#4": "1/2\"",
                    "#5": "5/8\""
                }
                
                # Obtener valores predeterminados de tkinter
                espaciamiento_predeterminado = float(default_valores.get(tipo_prelosa, {}).get('espaciamiento', 0.20))
                acero_predeterminado = default_valores.get(tipo_prelosa, {}).get('acero', "3/8\"")
                
                # Variables para rastrear si se han encontrado datos
                datos_longitudinales_encontrados = False
                datos_transversales_encontrados = False
                
                # Función para convertir diámetros alternativos al formato correcto
                def convertir_diametro(diametro_texto):
                    # Verificar si está en nuestro diccionario de conversión
                    if diametro_texto in conversion_diametros:
                        return conversion_diametros[diametro_texto]
                    return diametro_texto
                
                # Función para extraer información de las notaciones de acero
                def procesar_texto_acero(texto):
                    # Variable para almacenar resultados
                    resultado = {"cantidad": "1", "diametro_con_comillas": None, "separacion_decimal": espaciamiento_predeterminado}
                    
                    try:
                        # Limpiar formato DXF si existe (como {\W0.8;texto})
                        texto_limpio = texto
                        formato_match = re.search(r'\{.*?;(.*?)\}', texto)
                        if formato_match:
                            texto_limpio = formato_match.group(1)
                            print(f"  → Limpiando formato DXF: '{texto}' -> '{texto_limpio}'")
                        
                        # Verificar si el texto (limpio) contiene notación directa (M6, #3, etc.)
                        if '#' in texto_limpio or 'M' in texto_limpio:
                            # Buscar patrones como "#3@20", "M8@30" sin el símbolo ∅
                            match = re.search(r'(M\d+|#\d+)[@,.]?\d+', texto_limpio)
                            if match:
                                resultado["diametro_con_comillas"] = convertir_diametro(match.group(1))
                                return resultado
                                
                            # Si no hay patrón con @, buscar solo el diámetro
                            match = re.search(r'(M\d+|#\d+)', texto_limpio)
                            if match:
                                resultado["diametro_con_comillas"] = convertir_diametro(match.group(1))
                                return resultado
                        
                        # Extraer cantidad (número antes de ∅)
                        cantidad_match = re.search(r'^(\d+)∅', texto_limpio)
                        resultado["cantidad"] = cantidad_match.group(1) if cantidad_match else "1"
                        
                        # Buscar notaciones alternativas con el símbolo ∅ (M6, M8, #3, etc.)
                        alt_match = re.search(r'∅(M\d+|#\d+)', texto_limpio)
                        if alt_match:
                            diametro_alt = alt_match.group(1)
                            resultado["diametro_con_comillas"] = convertir_diametro(diametro_alt)
                        else:
                            # Caso específico para milímetros
                            # Caso específico para milímetros (con tolerancia a espacios)
                            if "mm" in texto_limpio.lower():
                                mm_match = re.search(r'∅?\s*(\d+)\s*(?:mm\.?|\.?mm)', texto_limpio, re.IGNORECASE)
                                if mm_match:
                                    diametro = mm_match.group(1)
                                    resultado["diametro_con_comillas"] = f"{diametro}mm"
                            else:
                                # Caso para fraccionales
                                diametro_match = re.search(r'∅?([\d/]+)', texto_limpio)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    # Asegurarnos de añadir comillas si es necesario
                                    if "\"" not in diametro and "/" in diametro:
                                        resultado["diametro_con_comillas"] = f"{diametro}\""
                                    else:
                                        resultado["diametro_con_comillas"] = diametro
                            
                            # Si no se pudo extraer con ningún formato específico, usar regex genérico
                            if resultado["diametro_con_comillas"] is None:
                                diametro_match = re.search(r'∅?([\d/]+)', texto_limpio)
                                if diametro_match:
                                    diametro = diametro_match.group(1)
                                    # Si no tiene unidades y tenemos "mm" en alguna parte, asumimos mm
                                    if "mm" in texto_limpio.lower():
                                        resultado["diametro_con_comillas"] = f"{diametro}mm"
                                    else:
                                        resultado["diametro_con_comillas"] = diametro
                        
                        # Extraer espaciamiento del texto (común para ambos formatos)
                        espaciamiento_match = re.search(r'@[.,]?(\d+)', texto_limpio)
                        if espaciamiento_match:
                            separacion = int(espaciamiento_match.group(1))
                            resultado["separacion_decimal"] = separacion / 100
                        else:
                            # Usar valor predeterminado específico del tipo cuando no hay espaciamiento
                            resultado["separacion_decimal"] = espaciamiento_predeterminado
                            print(f"  → No se encontró espaciamiento en '{texto}', usando valor predeterminado: {resultado['separacion_decimal']}")
                        
                        return resultado
                            
                    except Exception as e:
                        print(f"  → Error procesando texto: {str(e)}")
                        return resultado
                
                # Procesar textos longitudinales principales
                if len(textos_longitudinal) > 0:
                    print(f"• Encontrados {len(textos_longitudinal)} textos longitudinales")
                    datos_longitudinales_encontrados = True
                    
                    # Procesar primer texto horizontal (G4, H4, J4)
                    if len(textos_longitudinal) >= 1:
                        texto = textos_longitudinal[0]
                        try:
                            # Usar la nueva función de procesamiento
                            resultado = procesar_texto_acero(texto)
                            
                            # Escribir en Excel (siempre)
                            ws.range('G4').value = int(resultado["cantidad"])
                            ws.range('H4').value = resultado["diametro_con_comillas"]
                            ws.range('J4').value = resultado["separacion_decimal"]
                            
                            print(f"  → Texto 1: '{texto}' → G4={resultado['cantidad']}, H4={resultado['diametro_con_comillas']}, J4={resultado['separacion_decimal']}")
                        except Exception as e:
                            print(f"  → Error procesando texto 1: {str(e)}")
                    
                    # Procesar segundo texto horizontal (G5, H5, J5) si existe
                    if len(textos_longitudinal) >= 2:
                        texto = textos_longitudinal[1]
                        try:
                            # Usar la nueva función de procesamiento
                            resultado = procesar_texto_acero(texto)
                            
                            # Escribir en Excel (siempre)
                            ws.range('G5').value = int(resultado["cantidad"])
                            ws.range('H5').value = resultado["diametro_con_comillas"]
                            ws.range('J5').value = resultado["separacion_decimal"]
                            
                            print(f"  → Texto 2: '{texto}' → G5={resultado['cantidad']}, H5={resultado['diametro_con_comillas']}, J5={resultado['separacion_decimal']}")
                        except Exception as e:
                            print(f"  → Error procesando texto 2: {str(e)}")

                # Procesar textos transversales
                if len(textos_transversal) > 0:
                    print(f"• Encontrados {len(textos_transversal)} textos transversales")
                    datos_transversales_encontrados = True
                    
                    # Procesar primer texto vertical (G14, H14, J14)
                    if len(textos_transversal) >= 1:
                        texto = textos_transversal[0]
                        try:
                            # Usar la nueva función de procesamiento
                            resultado = procesar_texto_acero(texto)
                            
                            # Escribir en Excel (siempre)
                            ws.range('G14').value = int(resultado["cantidad"])
                            ws.range('H14').value = resultado["diametro_con_comillas"]
                            ws.range('J14').value = resultado["separacion_decimal"]
                            
                            print(f"  → Transversal 1: '{texto}' → G14={resultado['cantidad']}, H14={resultado['diametro_con_comillas']}, J14={resultado['separacion_decimal']}")
                        except Exception as e:
                            print(f"  → Error procesando texto transversal 1: {str(e)}")
                    
                    # Procesar segundo texto vertical (G15, H15, J15) si existe
                    if len(textos_transversal) >= 2:
                        texto = textos_transversal[1]
                        try:
                            # Usar la nueva función de procesamiento
                            resultado = procesar_texto_acero(texto)
                            
                            # Escribir en Excel (siempre)
                            ws.range('G15').value = int(resultado["cantidad"])
                            ws.range('H15').value = resultado["diametro_con_comillas"]
                            ws.range('J15').value = resultado["separacion_decimal"]
                            
                            print(f"  → Transversal 2: '{texto}' → G15={resultado['cantidad']}, H15={resultado['diametro_con_comillas']}, J15={resultado['separacion_decimal']}")
                        except Exception as e:
                            print(f"  → Error procesando texto transversal 2: {str(e)}")
                    
                    # NUEVO: Procesar tercer texto vertical (G16, H16, J16) si existe
                    if len(textos_transversal) >= 3:
                        texto = textos_transversal[2]
                        try:
                            # Usar la nueva función de procesamiento
                            resultado = procesar_texto_acero(texto)
                            
                            # Escribir en Excel (siempre)
                            ws.range('G16').value = int(resultado["cantidad"])
                            ws.range('H16').value = resultado["diametro_con_comillas"]
                            ws.range('J16').value = resultado["separacion_decimal"]
                            
                            print(f"  → Transversal 3: '{texto}' → G16={resultado['cantidad']}, H16={resultado['diametro_con_comillas']}, J16={resultado['separacion_decimal']}")
                        except Exception as e:
                            print(f"  → Error procesando texto transversal 3: {str(e)}")

                # Procesar textos longitudinales adicionales
                if len(textos_long_adi) > 0:
                    print(f"• Encontrados {len(textos_long_adi)} textos longitudinales adicionales")
                    datos_longitudinales_encontrados = True
                    
                    # Obtener el espaciamiento por defecto de los valores de tkinter
                    espaciamiento_macizas_adi = espaciamiento_predeterminado  # Usar el valor del tipo actual
                    
                    # Colocar valores por defecto en primera fila
                    print(f"  → Colocando valores default: G4=1, H4={acero_predeterminado}, J4={espaciamiento_macizas_adi}")
                    ws.range('G4').value = 1
                    ws.range('H4').value = acero_predeterminado
                    ws.range('J4').value = espaciamiento_macizas_adi
                    
                    # Colocar los mismos valores por defecto en fila vertical si no hay transversales
                    if not datos_transversales_encontrados:
                        print(f"  → Colocando valores default: G14=1, H14={acero_predeterminado}, J14={espaciamiento_macizas_adi}")
                        ws.range('G14').value = 1
                        ws.range('H14').value = acero_predeterminado
                        ws.range('J14').value = espaciamiento_macizas_adi
                    
                    # Procesar los textos de acero long adi
                    datos_textos = []
                    
                    for i, texto in enumerate(textos_long_adi):
                        try:
                            # Usar la nueva función de procesamiento
                            resultado = procesar_texto_acero(texto)
                            
                            # Continuar solo si se extrajo un diámetro
                            if resultado["diametro_con_comillas"]:
                                # Guardar los datos procesados
                                datos_textos.append([int(resultado["cantidad"]), resultado["diametro_con_comillas"], resultado["separacion_decimal"]])
                                print(f"  → Long Adi #{i+1}: '{texto}' → cantidad={resultado['cantidad']}, diámetro={resultado['diametro_con_comillas']}, separación={resultado['separacion_decimal']}")
                            else:
                                print(f"  → No se pudo extraer diámetro de '{texto}'")
                        except Exception as e:
                            print(f"  → Error procesando texto adicional: {str(e)}")

                    # Colocar los valores extraídos en las filas adicionales (G5, H5, J5, etc.)
                    for i, datos in enumerate(datos_textos):
                        fila = 5 + i  # Comienza en fila 5
                        cantidad, diametro, separacion = datos
                        
                        ws.range(f'G{fila}').value = cantidad
                        ws.range(f'H{fila}').value = diametro
                        ws.range(f'J{fila}').value = separacion
                    
                    # Procesar aceros transversales adicionales
                    if len(textos_tra_adi) > 0:
                        print(f"• Encontrados {len(textos_tra_adi)} textos transversales adicionales")
                        datos_transversales_encontrados = True
                        
                        # Procesar los textos de acero transversal adi
                        datos_textos_tra = []
                        
                        for i, texto in enumerate(textos_tra_adi):
                            try:
                                # Usar la nueva función de procesamiento
                                resultado = procesar_texto_acero(texto)
                                
                                # Continuar solo si se extrajo un diámetro
                                if resultado["diametro_con_comillas"]:
                                    # Guardar los datos procesados
                                    datos_textos_tra.append([int(resultado["cantidad"]), resultado["diametro_con_comillas"], resultado["separacion_decimal"]])
                                    print(f"  → Trans Adi #{i+1}: '{texto}' → cantidad={resultado['cantidad']}, diámetro={resultado['diametro_con_comillas']}, separación={resultado['separacion_decimal']}")
                                else:
                                    print(f"  → No se pudo extraer diámetro de '{texto}'")
                            except Exception as e:
                                print(f"  → Error procesando texto transversal adicional: {str(e)}")
                                                    
                        # Colocar los valores extraídos en las filas adicionales (G15, H15, J15, etc.)
                        for i, datos in enumerate(datos_textos_tra):
                            fila = 15 + i  # Comienza en fila 15
                            cantidad, diametro, separacion = datos
                            
                            ws.range(f'G{fila}').value = cantidad
                            ws.range(f'H{fila}').value = diametro
                            ws.range(f'J{fila}').value = separacion

                # NUEVO: Caso cuando hay solo longitudinal pero no transversal
                elif datos_longitudinales_encontrados and not datos_transversales_encontrados:
                    print("• Solo se encontraron textos longitudinales - Usando valores predeterminados para transversal")
                    
                    # Colocar valores predeterminados para transversal
                    ws.range('G14').value = 1
                    ws.range('H14').value = acero_predeterminado
                    ws.range('J14').value = espaciamiento_predeterminado
                    print(f"  → Colocando valores default transversal: G14=1, H14={acero_predeterminado}, J14={espaciamiento_predeterminado}")
                    
                    # Limpiar celdas adicionales para evitar interferencias
                    ws.range('G15').value = 0
                    ws.range('G16').value = 0  # Limpiar también celda para tercer transversal

                # NUEVO: Caso cuando hay solo transversal pero no longitudinal
                elif datos_transversales_encontrados and not datos_longitudinales_encontrados:
                    print("• Solo se encontraron textos transversales - Usando valores predeterminados para longitudinal")
                    
                    # Colocar valores predeterminados para longitudinal
                    ws.range('G4').value = 1
                    ws.range('H4').value = acero_predeterminado
                    ws.range('J4').value = espaciamiento_predeterminado
                    print(f"  → Colocando valores default longitudinal: G4=1, H4={acero_predeterminado}, J4={espaciamiento_predeterminado}")
                    
                    # Limpiar celdas adicionales para evitar interferencias
                    ws.range('G5').value = 0

                # Caso cuando no hay textos
                elif len(textos_longitudinal) == 0 and len(textos_transversal) == 0 and len(textos_long_adi) == 0 and len(textos_tra_adi) == 0:
                    print("• No se encontraron textos de acero - Usando valores predeterminados")
                    
                    # CORREGIDO: Usar el valor del tipo actual, no hardcoded para 'PRELOSA MACIZA'
                    espaciamiento_macizas_adi = float(default_valores.get(tipo_prelosa, {}).get('espaciamiento', 0.20))
                    acero_predeterminado = default_valores.get(tipo_prelosa, {}).get('acero', "3/8\"")
                    
                    print(f"  → Colocando valores default: G4=1, H4={acero_predeterminado}, J4={espaciamiento_macizas_adi}")
                    
                    # Colocar valores por defecto en Excel
                    ws.range('G4').value = 1
                    ws.range('H4').value = acero_predeterminado
                    ws.range('J4').value = espaciamiento_macizas_adi

                    ws.range('G14').value = 1
                    ws.range('H14').value = acero_predeterminado
                    ws.range('J14').value = espaciamiento_macizas_adi
                    
                    # Limpiar otras celdas para evitar interferencias
                    ws.range('G5').value = 0
                    ws.range('G15').value = 0
                    ws.range('G16').value = 0  # Limpiar también celda para tercer transversal

                # Forzar recálculo y obtener valores calculados
                print("• Forzando recálculo de Excel...")
                ws.book.app.calculate()

                # Guardar valores calculados
                k8_valor = ws.range('K8').value
                k9_valor = ws.range('K9').value
                k17_valor = ws.range('K17').value
                k18_valor = ws.range('K18').value
                k19_valor = ws.range('K19').value
                k20_valor = ws.range('K20').value  # Nuevo: para tercer texto transversal

                print(f"• Resultados calculados (longitudinal): K8={k8_valor}, K9={k9_valor}")
                print(f"• Resultados calculados (transversal): K17={k17_valor}, K18={k18_valor}, K19={k19_valor}, K20={k20_valor if k20_valor else 'N/A'}")

                # ACERO LONGITUDINAL - Validar y seleccionar el valor adecuado con jerarquía
                if k8_valor is None or k8_valor <= 0.1 or k8_valor < 0:
                    # Si K8 no es válido, verificar K9
                    print("  → K8 fuera de rango o inválido, verificando K9")
                    if k9_valor is not None and k9_valor > 0:
                        print(f"  → Usando K9 para acero longitudinal (1/2\"): {k9_valor}")
                        as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                        if k9_valor <= 0.1:
                            print("  → ADVERTENCIA: K9 menor o igual a 0.1 pero se usará de todos modos")
                    elif k9_valor is not None and k9_valor < 0:
                        print(f"  → ALERTA: El valor K9 es negativo: {k9_valor}")
                        as_long_texto = f"ERROR: VALOR NEGATIVO EN K9 ({k9_valor})"
                    else:
                        print("  → K9 es None o inválido, usando valor de respaldo")
                        as_long_texto = "ERROR: ACERO INSUFICIENTE PARA ESTA PRELOSA"
                elif k8_valor < 0:
                    print(f"  → ALERTA: El valor K8 es negativo: {k8_valor}")
                    # Verificar K9 como alternativa
                    if k9_valor is not None and k9_valor > 0:
                        print(f"  → Usando K9 como alternativa: {k9_valor}")
                        as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                    else:
                        as_long_texto = f"ERROR: VALOR NEGATIVO en K8 ({k8_valor})"
                else:
                    # K8 está en rango válido
                    as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"

                # ACERO TRANSVERSAL 1 - siempre a 0.28 en prelosas macizas
                as_tra1_texto = "1Ø6 mm@.28"

                # ACERO TRANSVERSAL 2 - Validar con jerarquía: K18 > K19
                if k18_valor is None or k18_valor <= 0.1 or k18_valor < 0:
                    # Si K18 no es válido, verificar K19
                    print("  → K18 fuera de rango o inválido, verificando K19")
                    if k19_valor is not None and k19_valor > 0:
                        print(f"  → Usando K19 para acero transversal 2 (3/8\"): {k19_valor}")
                        as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                        if k19_valor <= 0.1:
                            print("  → ADVERTENCIA: K19 menor o igual a 0.1 pero se usará de todos modos")
                    elif k19_valor is not None and k19_valor < 0:
                        print(f"  → ALERTA: El valor K19 es negativo: {k19_valor}")
                        as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K19 ({k19_valor})"
                    else:
                        print("  → K19 es None o inválido, usando valor de respaldo")
                        as_tra2_texto = "ERROR: VERIFICAR CÁLCULOS DE ACERO"
                elif k18_valor < 0:
                    print(f"  → ALERTA: El valor K18 es negativo: {k18_valor}")
                    # Verificar K19 como alternativa
                    if k19_valor is not None and k19_valor > 0:
                        print(f"  → Usando K19 como alternativa: {k19_valor}")
                        as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                    else:
                        as_tra2_texto = f"ERROR: VALOR NEGATIVO en K18 ({k18_valor})"
                else:
                    # K18 está en rango válido
                    as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"

                # NUEVO: ACERO TRANSVERSAL 3 - Validar con K20
                as_tra3_texto = None
                if len(textos_transversal) >= 3 or (k20_valor is not None and k20_valor > 0):
                    print("  → Procesando tercer acero transversal")
                    if k20_valor is not None:
                        if k20_valor > 0:
                            # Usar el diámetro que se extrajo del texto o 5/8" si no se puede determinar
                            if len(textos_transversal) >= 3:
                                # Intentar usar el diámetro del tercer texto transversal
                                texto = textos_transversal[2]
                                # Limpiar formato DXF si existe
                                texto_limpio = texto
                                formato_match = re.search(r'\{.*?;(.*?)\}', texto)
                                if formato_match:
                                    texto_limpio = formato_match.group(1)
                                
                                # MEJORADO: Verificar si el texto contiene notación directa
                                diametro_texto = None
                                
                                if '#' in texto_limpio or 'M' in texto_limpio:
                                    # Buscar patrones como "#3@20", "M8@30" sin el símbolo ∅
                                    match = re.search(r'(M\d+|#\d+)[@,.]?\d+', texto_limpio)
                                    if match:
                                        diametro_texto = convertir_diametro(match.group(1))
                                    else:
                                        # Si no hay patrón con @, buscar solo el diámetro
                                        match = re.search(r'(M\d+|#\d+)', texto_limpio)
                                        if match:
                                            diametro_texto = convertir_diametro(match.group(1))
                                
                                if diametro_texto is None:
                                    # Buscar con el símbolo ∅
                                    diametro_match = re.search(r'∅?([\d/]+)', texto_limpio)
                                    if diametro_match:
                                        diametro = diametro_match.group(1)
                                        if "mm" in texto_limpio.lower():
                                            diametro_texto = f"{diametro}mm"
                                        elif "/" in diametro and "\"" not in diametro:
                                            diametro_texto = f"{diametro}\""
                                        else:
                                            diametro_texto = diametro
                                
                                if diametro_texto is None:
                                    diametro_texto = "5/8\""  # Valor por defecto si no se puede extraer
                            else:
                                diametro_texto = "5/8\""  # Valor por defecto si no hay tercer texto
                            
                            as_tra3_texto = f"1Ø{diametro_texto}@.{formatear_valor_espaciamiento(k20_valor)}"
                            print(f"  → Usando K20 para acero transversal 3 ({diametro_texto}): {k20_valor}")
                        elif k20_valor < 0:
                            print(f"  → ALERTA: El valor K20 es negativo: {k20_valor}")
                            as_tra3_texto = f"ERROR: VALOR NEGATIVO EN K20 ({k20_valor})"
                        else:
                            # K20 es 0 o muy pequeño
                            print("  → K20 es cero o demasiado pequeño, no se agregará tercer acero")
                    else:
                        print("  → K20 es None, no se agregará tercer acero")

                print("• Valores finales para bloque:")
                print(f"  → AS_LONG: {as_long_texto}")
                print(f"  → AS_TRA1: {as_tra1_texto}")
                if as_tra2_texto:
                    print(f"  → AS_TRA2: {as_tra2_texto}")
                if as_tra3_texto:
                    print(f"  → AS_TRA3: {as_tra3_texto}")

                print("=== FIN PROCESAMIENTO PRELOSA MACIZA ===\n")

                # Limpiar celdas para evitar interferencias cuando se procese la siguiente prelosa
            
            elif categoria_base == "ALIGERADA":
                print("\n=== PROCESANDO PRELOSA ALIGERADA ===")
                
                # Diccionario de conversión para notaciones alternativas de diámetros
                conversion_diametros = {
                    "M6": "6mm",
                    "M8": "8mm",
                    "#3": "3/8\"",
                    "#4": "1/2\"",
                    "#5": "5/8\""
                }
                
                # Función para convertir diámetros alternativos al formato correcto
                def convertir_diametro(diametro_texto):
                    # Verificar si está en nuestro diccionario de conversión
                    if diametro_texto in conversion_diametros:
                        return conversion_diametros[diametro_texto]
                    return diametro_texto
                
                # Función optimizada para extraer diámetro con formatos diversos
                # Función optimizada para extraer diámetro con formatos diversos
                def extraer_diametro(texto):
                    # Limpiar formato DXF si existe (como {\W0.8;texto})
                    texto_limpio = texto
                    formato_match = re.search(r'\{.*?;(.*?)\}', texto)
                    if formato_match:
                        texto_limpio = formato_match.group(1)
                        print(f"  → Limpiando formato DXF: '{texto}' -> '{texto_limpio}'")
                    
                    # Caso específico para '1Ø 8 mm@.175(Inf.)' y similares con espacios
                    match = re.search(r'^\d+Ø\s*(\d+)\s*mm', texto_limpio)
                    if match:
                        return f"{match.group(1)}mm"
                    
                    # Caso específico para patrones como "1∅3/8"@.225(Inf.)"
                    match = re.search(r'^\d+∅([\d/]+\")', texto_limpio)
                    if match:
                        return match.group(1)
                    
                    # Caso específico para textos como 1Ø3/8" con o sin texto adicional
                    match = re.search(r'\d+Ø([\d/]+\")', texto_limpio)
                    if match:
                        return match.group(1)
                    
                    # Verificar si el texto (limpio) contiene notación directa (M6, #3, etc.)
                    if '#' in texto_limpio or 'M' in texto_limpio:
                        # Buscar patrones como "#3@20", "M8@30" sin el símbolo ∅
                        match = re.search(r'(M\d+|#\d+)[@,.]?\d+', texto_limpio)
                        if match:
                            return convertir_diametro(match.group(1))
                            
                        # Si no hay patrón con @, buscar solo el diámetro
                        match = re.search(r'(M\d+|#\d+)', texto_limpio)
                        if match:
                            return convertir_diametro(match.group(1))
                    
                    # Buscar notaciones alternativas con el símbolo ∅ (M6, M8, #3, etc.)
                    match = re.search(r'∅(M\d+|#\d+)', texto_limpio)
                    if match:
                        return convertir_diametro(match.group(1))
                    
                    # Caso para milímetros con posible espacio entre número y 'mm'
                    if "mm" in texto_limpio.lower():
                        # Buscar patrón con posible espacio entre número y 'mm'
                        match = re.search(r'∅?\s*(\d+)\s*(?:mm\.?|\.?mm)', texto_limpio, re.IGNORECASE)
                        if match:
                            return f"{match.group(1)}mm"
                    
                    # Caso para fraccionales
                    match = re.search(r'∅?([\d/]+)', texto_limpio)
                    if match:
                        diametro = match.group(1)
                        # Si es un número simple y el texto menciona mm, aplicar formato mm
                        if "mm" in texto_limpio.lower() and "/" not in diametro:
                            return f"{diametro}mm"
                        # Si es fraccional, añadir comillas si es necesario
                        elif "/" in diametro and "\"" not in diametro:
                            return f"{diametro}\""
                        else:
                            return diametro
                    
                    return None  # No se pudo extraer diámetro
                
                # Función para procesar texto de acero y extraer información
                def procesar_texto_acero(texto, valores_default):
                    resultado = {
                        "cantidad": "1", 
                        "diametro": valores_default.get("acero", "3/8\""), 
                        "separacion": valores_default.get("espaciamiento", 0.605)
                    }
                    
                    try:
                        # Limpiar formato DXF si existe
                        texto_limpio = texto
                        formato_match = re.search(r'\{.*?;(.*?)\}', texto)
                        if formato_match:
                            texto_limpio = formato_match.group(1)
                            print(f"  → Limpiando formato para procesamiento: '{texto}' -> '{texto_limpio}'")
                        
                        # Extraer cantidad (número antes de ∅, si no hay se asume 1)
                        cantidad_match = re.search(r'^(\d+)∅', texto_limpio)
                        resultado["cantidad"] = cantidad_match.group(1) if cantidad_match else "1"
                        
                        # Extraer diámetro del texto
                        diametro_con_comillas = extraer_diametro(texto_limpio)
                        if diametro_con_comillas:
                            resultado["diametro"] = diametro_con_comillas
                            print(f"  → Diámetro extraído: {diametro_con_comillas}")
                        else:
                            print(f"  → No se pudo extraer diámetro, usando valor por defecto: {resultado['diametro']}")
                        
                        # Extraer espaciamiento del texto (buscar patrones como @20, @.20, etc.)
                        espaciamiento_match = re.search(r'@[,.]?(\d+)', texto_limpio)
                        if espaciamiento_match:
                            separacion = espaciamiento_match.group(1)
                            resultado["separacion"] = float(f"0.{separacion}")
                        
                        return resultado
                    except Exception as e:
                        print(f"  → Error procesando texto '{texto}': {str(e)}")
                        return resultado
                
                # Obtener valores predeterminados específicos para este tipo de prelosa
                valores_default = {
                    "espaciamiento": float(default_valores.get(tipo_prelosa, {}).get('espaciamiento', 0.605)),
                    "acero": default_valores.get(tipo_prelosa, {}).get('acero', "3/8\"")
                }
                
                print(f"Usando valores predeterminados para {tipo_prelosa}:")
                print(f"  - Espaciamiento: {valores_default['espaciamiento']}")
                print(f"  - Acero: {valores_default['acero']}")
                
                # Variables para rastrear si se han encontrado datos
                datos_longitudinales_encontrados = False
                datos_transversales_encontrados = False
                
                # Imprimir todos los textos encontrados para depuración
                print("TEXTOS ENCONTRADOS PARA DEPURACIÓN:")
                print(f"Textos transversales ({len(textos_transversal)}): {textos_transversal}")
                print(f"Textos longitudinales ({len(textos_longitudinal)}): {textos_longitudinal}")
                print(f"Textos longitudinales adicionales ({len(textos_long_adi)}): {textos_long_adi}")
                print(f"Textos transversales adicionales ({len(textos_tra_adi)}): {textos_tra_adi}")
                
                # Combinar textos verticales y horizontales para procesar
                textos_a_procesar = textos_transversal + textos_longitudinal
                print(f"Total textos a procesar (vertical + horizontal): {len(textos_a_procesar)}")
                
                # Procesar los textos (independientemente si son verticales u horizontales)
                if len(textos_a_procesar) > 0:
                    print(f"Procesando {len(textos_a_procesar)} textos en {tipo_prelosa}")
                    datos_longitudinales_encontrados = True
                    
                    # Limpiar celda g5 preventivamente
                    ws.range('G5').value = 0
                    
                    # Procesar hasta 2 textos principales (si hay)
                    for i, texto in enumerate(textos_a_procesar[:2]):
                        fila = 4 + i  # Comenzar en fila 4 (G4, H4, J4) y seguir con fila 5
                        print(f"Procesando texto #{i+1}: '{texto}' para fila {fila}")
                        
                        try:
                            # Procesar el texto y obtener datos
                            resultado = procesar_texto_acero(texto, valores_default)
                            
                            # Escribir en Excel
                            ws.range(f'G{fila}').value = int(resultado["cantidad"])
                            ws.range(f'H{fila}').value = resultado["diametro"]
                            ws.range(f'J{fila}').value = resultado["separacion"]
                            
                            print(f"  → Colocado en Excel: G{fila}={resultado['cantidad']}, " +
                                f"H{fila}={resultado['diametro']}, J{fila}={resultado['separacion']}")
                        except Exception as e:
                            if i == 0:  # Si es el primer texto, colocar valores por defecto en caso de error
                                print(f"  → Error procesando texto: {e}. Usando valores por defecto")
                                ws.range(f'G{fila}').value = 1
                                ws.range(f'H{fila}').value = valores_default["acero"]
                                ws.range(f'J{fila}').value = valores_default["espaciamiento"]
                else:
                    print(f"ADVERTENCIA: No se encontraron textos principales para {tipo_prelosa}")
                    
                    # Si no hay textos principales pero hay adicionales, colocar valores por defecto
                    if len(textos_long_adi) > 0 or len(textos_tra_adi) > 0:
                        print("Colocando valores por defecto en celdas principales para el caso adicional")
                        ws.range('G4').value = 1
                        ws.range('H4').value = valores_default["acero"]
                        ws.range('J4').value = valores_default["espaciamiento"]
                        datos_longitudinales_encontrados = True
                    else:
                        # Si no hay ningún texto, usar valores predeterminados para ambos
                        print("No se encontraron textos de ningún tipo. Colocando valores predeterminados")
                        
                        # Acero longitudinal (G4, H4, J4)
                        ws.range('G4').value = 1
                        ws.range('H4').value = valores_default["acero"]
                        ws.range('J4').value = valores_default["espaciamiento"]
                        
                        # Acero transversal (G14, H14, J14)
                        ws.range('G14').value = 1
                        ws.range('H14').value = "6mm"
                        ws.range('J14').value = 0.50
                        
                        datos_longitudinales_encontrados = True
                        datos_transversales_encontrados = True
                
                # Procesar textos longitudinales adicionales
                if len(textos_long_adi) > 0:
                    print("=" * 60)
                    print(f"PROCESANDO {len(textos_long_adi)} TEXTOS LONG ADI EN {tipo_prelosa}")
                    print("=" * 60)
                    
                    # Procesar los textos de acero long adi
                    datos_textos = []
                    
                    for i, texto in enumerate(textos_long_adi):
                        print(f"TEXTO #{i+1}: '{texto}'")
                        try:
                            # Procesar texto y guardar resultados
                            resultado = procesar_texto_acero(texto, valores_default)
                            datos_textos.append([
                                int(resultado["cantidad"]), 
                                resultado["diametro"], 
                                resultado["separacion"]
                            ])
                            print(f"  ✓ Datos procesados: cantidad={resultado['cantidad']}, " +
                                f"diámetro={resultado['diametro']}, separación={resultado['separacion']}")
                        except Exception as e:
                            print(f"  ✗ Error procesando texto: {e}")
                    
                    # Colocar los valores extraídos en filas adicionales
                    print("\nColocando valores en filas adicionales:")
                    for i, datos in enumerate(datos_textos):
                        fila = 5 + i  # Comienza en fila 5
                        cantidad, diametro, separacion = datos
                        
                        ws.range(f'G{fila}').value = cantidad
                        ws.range(f'H{fila}').value = diametro
                        ws.range(f'J{fila}').value = separacion
                        print(f"  ✓ Fila {fila}: G{fila}={cantidad}, H{fila}={diametro}, J{fila}={separacion}")
                
                # Procesar textos transversales
                if len(textos_transversal) > 0:
                    print(f"\n• Encontrados {len(textos_transversal)} textos transversales")
                    datos_transversales_encontrados = True
                    
                    # Procesar primer texto vertical (G14, H14, J14)
                    if len(textos_transversal) >= 1:
                        texto = textos_transversal[0]
                        try:
                            # Valores default específicos para transversales
                            valores_default_trans = valores_default.copy()
                            valores_default_trans["espaciamiento"] = 0.50
                            valores_default_trans["acero"] = "6mm"
                            
                            # Procesar texto
                            resultado = procesar_texto_acero(texto, valores_default_trans)
                            
                            # Escribir en Excel
                            ws.range('G14').value = int(resultado["cantidad"])
                            ws.range('H14').value = resultado["diametro"]
                            ws.range('J14').value = resultado["separacion"]
                            
                            print(f"  → Transversal: G14={resultado['cantidad']}, " +
                                f"H14={resultado['diametro']}, J14={resultado['separacion']}")
                        except Exception as e:
                            # En caso de error, usar valores predeterminados para transversal
                            print(f"  → Error en transversal: {e}. Usando valores por defecto")
                            ws.range('G14').value = 1
                            ws.range('H14').value = "6mm"
                            ws.range('J14').value = 0.50
                elif datos_longitudinales_encontrados:
                    # Si no hay textos transversales, usar valores predeterminados
                    print("• No hay textos transversales - Usando valores predeterminados")
                    ws.range('G14').value = 1
                    ws.range('H14').value = "6mm"
                    ws.range('J14').value = 0.50
                    datos_transversales_encontrados = True
                
                # Procesar textos transversales adicionales (tra_adi)
                if len(textos_tra_adi) > 0:
                    print("\n" + "=" * 60)
                    print(f"PROCESANDO {len(textos_tra_adi)} TEXTOS TRANSVERSALES ADI")
                    print("=" * 60)
                    
                    # Valores default específicos para transversales adicionales
                    valores_default_trans = {"espaciamiento": 0.50, "acero": "6mm"}
                    
                    # Procesar los textos
                    datos_textos_tra = []
                    for i, texto in enumerate(textos_tra_adi):
                        print(f"TEXTO TRANSVERSAL #{i+1}: '{texto}'")
                        try:
                            # Procesar texto y guardar resultados
                            resultado = procesar_texto_acero(texto, valores_default_trans)
                            datos_textos_tra.append([
                                int(resultado["cantidad"]), 
                                resultado["diametro"], 
                                resultado["separacion"]
                            ])
                            print(f"  ✓ Datos procesados: cantidad={resultado['cantidad']}, " +
                                f"diámetro={resultado['diametro']}, separación={resultado['separacion']}")
                        except Exception as e:
                            print(f"  ✗ Error procesando texto: {e}")
                    
                    # Colocar valores en filas adicionales
                    print("\nColocando valores transversales en filas adicionales:")
                    for i, datos in enumerate(datos_textos_tra):
                        fila = 15 + i  # Comienza en fila 15
                        cantidad, diametro, separacion = datos
                        
                        ws.range(f'G{fila}').value = cantidad
                        ws.range(f'H{fila}').value = diametro
                        ws.range(f'J{fila}').value = separacion
                        print(f"  ✓ Fila {fila}: G{fila}={cantidad}, H{fila}={diametro}, J{fila}={separacion}")
                
                # Verificar valores antes de recalcular
                print("VALORES ANTES DE RECALCULAR:")
                print(f"  Celda K8 = {ws.range('K8').value}")
                print(f"  Celda K9 = {ws.range('K9').value}")
                print(f"  Celda K10 = {ws.range('K10').value}")
                print(f"  Celda K17 = {ws.range('K17').value}")
                print(f"  Celda K18 = {ws.range('K18').value}")
                print(f"  Celda K19 = {ws.range('K19').value}")
                
                # Forzar recálculo y obtener valores calculados
                print("• Forzando recálculo de Excel...")
                ws.book.app.calculate()
                
                # Guardar valores calculados
                k8_valor = ws.range('K8').value
                k9_valor = ws.range('K9').value
                k10_valor = ws.range('K10').value
                k17_valor = ws.range('K17').value
                k18_valor = ws.range('K18').value
                k19_valor = ws.range('K19').value
                
                print(f"• Resultados (longitudinal): K8={k8_valor}, K9={k9_valor}, K10={k10_valor}")
                print(f"• Resultados (transversal): K17={k17_valor}, K18={k18_valor}, K19={k19_valor}")
                
                # ACERO LONGITUDINAL - Usando jerarquía K8 > K9 > K10 con validación de valores negativos
                if k8_valor is None or k8_valor <= 0.1:
                    print("  → K8 fuera de rango o inválido, verificando K9")
                    if k9_valor is not None and k9_valor > 0:
                        print(f"  → Usando K9 para acero longitudinal (1/2\"): {k9_valor}")
                        as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                    elif k9_valor is not None and k9_valor < 0:
                        print(f"  → K9 es negativo, verificando K10")
                        if k10_valor is not None and k10_valor > 0:
                            print(f"  → Usando K10 para acero longitudinal (8mm): {k10_valor}")
                            as_long_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k10_valor)}"
                        else:
                            print("  → Valores inválidos, usando valor de respaldo")
                            as_long_texto = f"1Ø{valores_default['acero']}@.{formatear_valor_espaciamiento(valores_default['espaciamiento'])}"
                    else:
                        print("  → K9 inválido, verificando K10")
                        if k10_valor is not None and k10_valor > 0:
                            print(f"  → Usando K10: {k10_valor}")
                            as_long_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k10_valor)}"
                        else:
                            print("  → Usando valores de respaldo")
                            as_long_texto = f"1Ø{valores_default['acero']}@.{formatear_valor_espaciamiento(valores_default['espaciamiento'])}"
                elif k8_valor < 0:
                    print(f"  → K8 es negativo, verificando alternativas")
                    if k9_valor is not None and k9_valor > 0:
                        print(f"  → Usando K9: {k9_valor}")
                        as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                    elif k10_valor is not None and k10_valor > 0:
                        print(f"  → Usando K10: {k10_valor}")
                        as_long_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k10_valor)}"
                    else:
                        print("  → Usando valores de respaldo")
                        as_long_texto = f"1Ø{valores_default['acero']}@.{formatear_valor_espaciamiento(valores_default['espaciamiento'])}"
                else:
                    # K8 está en rango válido
                    as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                
                # Determinar tipo de tratamiento basado en el tipo específico de prelosa
                if "2 SENT" in tipo_prelosa.upper():
                    # Para prelosas aligeradas con 2 sentidos (bidireccionales)
                    as_tra1_texto = "1Ø6 mm@.28"
                    
                    # Para AS_TRA2 - Validación con manejo de valores negativos
                    if k18_valor is None or k18_valor <= 0.1 or k18_valor < 0:
                        if k19_valor is not None and k19_valor > 0:
                            print(f"  → Usando K19: {k19_valor}")
                            as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                        else:
                            print("  → Usando valor de respaldo para 2 SENT")
                            as_tra2_texto = "1Ø8 mm@.50"
                    else:
                        # K18 está en rango válido
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                else:
                    # Para prelosas aligeradas normales (un sentido)
                    as_tra1_texto = "1Ø6 mm@.50"
                    
                    # Validar as_tra2
                    if k18_valor is not None and k18_valor > 0 and k18_valor > 0.1:
                        print(f"  → Usando K18: {k18_valor}")
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                    elif k19_valor is not None and k19_valor > 0:
                        print(f"  → Usando K19: {k19_valor}")
                        as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                    else:
                        print("  → Usando valor predeterminado para aligeradas estándar")
                        as_tra2_texto = "1Ø8 mm@.50"
                
                # Verificación de valores críticos
                if ("ERROR" in as_long_texto) or ("ERROR" in as_tra2_texto):
                    print("\n¡ADVERTENCIA! Se detectaron errores en los valores calculados.")
                    print("Se recomienda revisar los cálculos o valores de entrada.")
                
                # Guardar valores finales
                print("\nVALORES FINALES PARA EL BLOQUE:")
                print(f"  → AS_LONG: {as_long_texto}")
                print(f"  → AS_TRA1: {as_tra1_texto}")
                print(f"  → AS_TRA2: {as_tra2_texto}")
                
                # Almacenar los valores finales en variables globales
                as_long_final = as_long_texto
                as_tra1_final = as_tra1_texto
                as_tra2_final = as_tra2_texto
                
                # Limpiar celdas para evitar interferencias en futuros cálculos

                
                print("=== FIN PROCESAMIENTO PRELOSA ALIGERADA ===\n")

            elif categoria_base == "ALIGERADA_2SENT":
                print("\n=== PROCESANDO PRELOSA ALIGERADA - 2 SENT ===")
                
                # Diccionario de conversión para notaciones alternativas de diámetros
                conversion_diametros = {
                    "M6": "6mm",
                    "M8": "8mm",
                    "#3": "3/8\"",
                    "#4": "1/2\"",
                    "#5": "5/8\""
                }
                
                # Función para convertir diámetros alternativos al formato correcto
                def convertir_diametro(diametro_texto):
                    # Verificar si está en nuestro diccionario de conversión
                    if diametro_texto in conversion_diametros:
                        return conversion_diametros[diametro_texto]
                    return diametro_texto
                
                # Función optimizada para extraer diámetro con formatos diversos
                def extraer_diametro(texto):
                    # Limpiar formato DXF si existe (como {\W0.8;texto})
                    texto_limpio = texto
                    formato_match = re.search(r'\{.*?;(.*?)\}', texto)
                    if formato_match:
                        texto_limpio = formato_match.group(1)
                        print(f"  → Limpiando formato DXF: '{texto}' -> '{texto_limpio}'")
                    
                    # CORREGIDO: Caso específico para '1Ø 8 mm@.175(Inf.)' y similares con espacios
                    match = re.search(r'^\d+Ø\s*(\d+)\s*mm', texto_limpio)
                    if match:
                        return f"{match.group(1)}mm"
                    
                    # Caso específico para patrones como "1∅1/2"" o "1∅5/8""
                    match = re.search(r'^\d+∅([\d/]+\")', texto_limpio)
                    if match:
                        return match.group(1)
                    
                    # Caso específico para textos como 1Ø3/8" con o sin texto adicional
                    match = re.search(r'\d+Ø([\d/]+\")', texto_limpio)
                    if match:
                        return match.group(1)
                        
                    # Verificar si el texto (limpio) contiene notación directa (M6, #3, etc.)
                    if '#' in texto_limpio or 'M' in texto_limpio:
                        # Buscar patrones como "#3@20", "M8@30" sin el símbolo ∅
                        match = re.search(r'(M\d+|#\d+)[@,.]?\d+', texto_limpio)
                        if match:
                            return convertir_diametro(match.group(1))
                            
                        # Si no hay patrón con @, buscar solo el diámetro
                        match = re.search(r'(M\d+|#\d+)', texto_limpio)
                        if match:
                            return convertir_diametro(match.group(1))
                    
                    # Buscar notaciones alternativas con el símbolo ∅ (M6, M8, #3, etc.)
                    match = re.search(r'∅(M\d+|#\d+)', texto_limpio)
                    if match:
                        return convertir_diametro(match.group(1))
                    
                    # CORREGIDO: Caso para milímetros con posible espacio entre número y 'mm'
                    if "mm" in texto_limpio.lower():
                        match = re.search(r'∅?\s*(\d+)\s*(?:mm\.?|\.?mm)', texto_limpio, re.IGNORECASE)
                        if match:
                            return f"{match.group(1)}mm"
                    
                    # Caso para fraccionales
                    match = re.search(r'∅?\s*([\d/]+)', texto_limpio)
                    if match:
                        diametro = match.group(1)
                        # Si es un número simple y el texto menciona mm, aplicar formato mm
                        if "mm" in texto_limpio.lower() and "/" not in diametro:
                            return f"{diametro}mm"
                        # Si es fraccional, añadir comillas si es necesario
                        elif "/" in diametro and "\"" not in diametro:
                            return f"{diametro}\""
                        else:
                            return diametro
                    
                    # NUEVO: Verificar cualquier número seguido de mm (con espacio opcional)
                    if "mm" in texto_limpio.lower():
                        match = re.search(r'(\d+)\s*mm', texto_limpio, re.IGNORECASE)
                        if match:
                            return f"{match.group(1)}mm"
                    
                    return None  # No se pudo extraer diámetro

                # Usar los valores predeterminados (vienen del tkinter)
                espaciamiento_predeterminado = float(default_valores.get(tipo_prelosa, {}).get('espaciamiento', 0.605))
                acero_predeterminado = default_valores.get(tipo_prelosa, {}).get('acero', "3/8\"")
                print(f"Usando valores predeterminados para {tipo_prelosa}:")
                print(f"  - Espaciamiento: {espaciamiento_predeterminado}")
                print(f"  - Acero: {acero_predeterminado}")

                # Limpiar celdas para evitar interferencias
                ws.range('G5').value = 0
                ws.range('G6').value = 0
                ws.range('G15').value = 0
                ws.range('G16').value = 0

                # Función auxiliar para procesar texto y escribir en Excel
                def procesar_texto_acero(texto, celda_cantidad, celda_diametro, celda_espaciamiento):
                    try:
                        # Limpiar formato DXF si existe
                        texto_limpio = texto
                        formato_match = re.search(r'\{.*?;(.*?)\}', texto)
                        if formato_match:
                            texto_limpio = formato_match.group(1)
                            print(f"  → Limpiando formato para procesamiento: '{texto}' -> '{texto_limpio}'")
                        
                        # Extraer cantidad (número antes de ∅)
                        cantidad_match = re.search(r'^(\d+)∅', texto_limpio)
                        cantidad = cantidad_match.group(1) if cantidad_match else "1"
                        
                        # Extraer diámetro del texto usando la función auxiliar
                        diametro_con_comillas = extraer_diametro(texto_limpio)
                        
                        # NUEVO: Verificación adicional para casos específicos como "1Ø 8 mm@.50(Inf.)"
                        if not diametro_con_comillas and ("8 mm" in texto_limpio or "8mm" in texto_limpio):
                            print(f"  → Detectado '8 mm' en el texto pero no se extrajo correctamente, forzando a '8mm'")
                            diametro_con_comillas = "8mm"
                        
                        if diametro_con_comillas:
                            print(f"  → Diámetro extraído: {diametro_con_comillas}")
                            cantidad = int(cantidad)  # Convertir a entero
                            
                            # Extraer espaciamiento del texto si existe
                            espaciamiento_match = re.search(r'@[,.]?(\d+)', texto_limpio)
                            if espaciamiento_match:
                                separacion = int(espaciamiento_match.group(1))
                                separacion_decimal = separacion / 100
                                print(f"  → Espaciamiento extraído del texto: @{separacion} → {separacion_decimal}")
                            else:
                                # Usar el valor predeterminado para el espaciamiento
                                separacion_decimal = espaciamiento_predeterminado
                                print(f"  → No se encontró espaciamiento en '{texto}', usando valor predeterminado: {espaciamiento_predeterminado}")
                            
                            # Escribir en Excel
                            ws.range(celda_cantidad).value = cantidad
                            ws.range(celda_diametro).value = diametro_con_comillas
                            ws.range(celda_espaciamiento).value = separacion_decimal
                            
                            print(f"  → Colocado en Excel: {celda_cantidad}={cantidad}, {celda_diametro}={diametro_con_comillas}, {celda_espaciamiento}={separacion_decimal}")
                            return True
                        else:
                            print(f"  → No se pudo extraer información del diámetro en el texto '{texto}'")
                            # Usar valores predeterminados
                            ws.range(celda_cantidad).value = 1
                            ws.range(celda_diametro).value = acero_predeterminado
                            ws.range(celda_espaciamiento).value = espaciamiento_predeterminado
                            print(f"  → Usando valores predeterminados: {celda_cantidad}=1, {celda_diametro}={acero_predeterminado}, {celda_espaciamiento}={espaciamiento_predeterminado}")
                            return False
                    except Exception as e:
                        print(f"  → Error al procesar texto '{texto}': {e}")
                        # Usar valores predeterminados en caso de error
                        ws.range(celda_cantidad).value = 1
                        ws.range(celda_diametro).value = acero_predeterminado
                        ws.range(celda_espaciamiento).value = espaciamiento_predeterminado
                        print(f"  → Error: usando valores predeterminados: {celda_cantidad}=1, {celda_diametro}={acero_predeterminado}, {celda_espaciamiento}={espaciamiento_predeterminado}")
                        return False
                
                # Caso 1: Si tenemos textos horizontales
                if len(textos_longitudinal) > 0:
                    print(f"Procesando {len(textos_longitudinal)} textos horizontales")
                    
                    # Procesar primer texto horizontal (G4, H4, J4)
                    if len(textos_longitudinal) >= 1:
                        procesar_texto_acero(textos_longitudinal[0], 'G4', 'H4', 'J4')
                    
                    # Procesar segundo texto horizontal (G5, H5, J5) si existe
                    if len(textos_longitudinal) >= 2:
                        procesar_texto_acero(textos_longitudinal[1], 'G5', 'H5', 'J5')
                
                # Caso 2: Si tenemos textos verticales
                if len(textos_transversal) > 0:
                    print(f"Procesando {len(textos_transversal)} textos verticales")
                    
                    # Procesar primer texto vertical (G14, H14, J14)
                    if len(textos_transversal) >= 1:
                        procesar_texto_acero(textos_transversal[0], 'G14', 'H14', 'J14')
                    
                    # Procesar segundo texto vertical (G15, H15, J15) si existe
                    if len(textos_transversal) >= 2:
                        procesar_texto_acero(textos_transversal[1], 'G15', 'H15', 'J15')
                    
                    # Procesar tercer texto vertical (G16, H16, J16) si existe
                    if len(textos_transversal) >= 3:
                        procesar_texto_acero(textos_transversal[2], 'G16', 'H16', 'J16')
                
                # NUEVO: Si no hay textos transversales, colocar valores predeterminados para transversales
                if len(textos_transversal) == 0:
                    print("No se encontraron textos transversales. Colocando valores por defecto para transversales:")
                    
                    # Colocar valores predeterminados para transversales (usando los que vienen del tkinter)
                    ws.range('G14').value = 1
                    ws.range('H14').value = acero_predeterminado
                    ws.range('J14').value = espaciamiento_predeterminado
                
                # Si no hay textos principales pero hay adicionales, poner valores predeterminados
                if len(textos_longitudinal) == 0 and len(textos_transversal) == 0 and (len(textos_long_adi) > 0 or len(textos_tra_adi) > 0):
                    print("No se encontraron textos principales pero hay adicionales. Colocando valores por defecto:")
                    
                    # Colocar valores por defecto para longitudinales
                    ws.range('G4').value = 1
                    ws.range('H4').value = acero_predeterminado
                    ws.range('J4').value = espaciamiento_predeterminado
                    
                    # Colocar valores por defecto para transversales
                    ws.range('G14').value = 1
                    ws.range('H14').value = acero_predeterminado
                    ws.range('J14').value = espaciamiento_predeterminado
                
                # NUEVO: Si no hay textos de ningún tipo, colocar valores predeterminados
                if len(textos_longitudinal) == 0 and len(textos_transversal) == 0 and len(textos_long_adi) == 0 and len(textos_tra_adi) == 0:
                    print("No se encontraron textos de ningún tipo. Colocando valores predeterminados:")
                    
                    # Colocar valores por defecto para longitudinales
                    ws.range('G4').value = 1
                    ws.range('H4').value = acero_predeterminado
                    ws.range('J4').value = espaciamiento_predeterminado
                    
                    # Colocar valores por defecto para transversales
                    ws.range('G14').value = 1
                    ws.range('H14').value = acero_predeterminado
                    ws.range('J14').value = espaciamiento_predeterminado
                
                # Forzar recálculo y guardar los valores calculados
                print("Forzando recálculo de Excel...")
                ws.book.app.calculate()
                
                # Intentar un segundo cálculo para asegurar que Excel procesó los valores
                wb.app.calculate()
                
                # Obtener valores de celdas para validación
                k8_valor = ws.range('K8').value
                k9_valor = ws.range('K9').value
                k18_valor = ws.range('K18').value
                k19_valor = ws.range('K19').value
                
                # Obtener los valores calculados (as_long y as_tra2)
                as_long = k8_valor  # as_long es el valor calculado en K8
                as_tra2 = k18_valor  # as_tra2 es el valor calculado en K18
                
                print(f"• Valores calculados: K8={k8_valor}, K9={k9_valor}")
                print(f"• Valores calculados: K18={k18_valor}, K19={k19_valor}")
                
                # ACERO LONGITUDINAL - Validar y seleccionar el valor adecuado con jerarquía
                if as_long is None or as_long <= 0.1:
                    print("  → as_long menor o igual a 0.1, verificando K8")
                    # Verificar K8 primero
                    if k8_valor is not None and k8_valor >= 0.1 and k8_valor > 0:
                        print(f"  → Usando K8 para acero longitudinal (3/8\"): {k8_valor}")
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                    elif k8_valor is not None and k8_valor < 0:
                        print(f"  → ALERTA: El valor K8 es negativo: {k8_valor}")
                        # Continuar con K9
                        if k9_valor is not None and k9_valor > 0:
                            print(f"  → Usando K9 como alternativa: {k9_valor}")
                            as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                        else:
                            as_long_texto = f"ERROR: VALOR NEGATIVO EN K8 ({k8_valor})"
                    else:
                        # Si K8 no es válido, verificar K9
                        print("  → K8 fuera de rango o inválido, verificando K9")
                        if k9_valor is not None and k9_valor > 0:
                            print(f"  → Usando K9 para acero longitudinal (1/2\"): {k9_valor}")
                            as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                            if k9_valor <= 0.1:
                                print("  → ADVERTENCIA: K9 menor o igual a 0.1 pero se usará de todos modos")
                        elif k9_valor is not None and k9_valor < 0:
                            print(f"  → ALERTA: El valor K9 es negativo: {k9_valor}")
                            as_long_texto = f"ERROR: VALOR NEGATIVO EN K9 ({k9_valor})"
                        else:
                            print("  → K9 es None o inválido, usando valor de respaldo")
                            as_long_texto = "ERROR: ACERO INSUFICIENTE PARA ESTA PRELOSA"
                elif as_long < 0:
                    print(f"  → ALERTA: El valor as_long es negativo: {as_long}")
                    # Verificar alternativas en orden: K8, K9
                    if k8_valor is not None and k8_valor > 0:
                        print(f"  → Usando K8 como alternativa: {k8_valor}")
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                    elif k9_valor is not None and k9_valor > 0:
                        print(f"  → Usando K9 como alternativa: {k9_valor}")
                        as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                    else:
                        as_long_texto = f"ERROR: VALOR NEGATIVO en as_long ({as_long})"
                else:
                    # as_long está en rango válido
                    as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                
                # Para acero vertical (AS_TRA1) - siempre a 0.28 en prelosas aligeradas 2 sent
                as_tra1_texto = "1Ø6 mm@.28"
                
                # ACERO TRANSVERSAL 2 - Validar con jerarquía
                if as_tra2 is not None:
                    if as_tra2 < 0:
                        print(f"  → ALERTA: El valor as_tra2 es negativo: {as_tra2}")
                        # Verificar K18 como alternativa
                        if k18_valor is not None and k18_valor > 0:
                            print(f"  → Usando K18 como alternativa: {k18_valor}")
                            as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                        elif k19_valor is not None and k19_valor > 0:
                            print(f"  → Usando K19 como alternativa: {k19_valor}")
                            as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                        else:
                            as_tra2_texto = f"ERROR: VALOR NEGATIVO en as_tra2 ({as_tra2})"
                    elif as_tra2 <= 0.1:
                        print("  → as_tra2 menor o igual a 0.1, verificando K18")
                        # Verificar K18 primero
                        if k18_valor is not None and k18_valor >= 0.1 and k18_valor > 0:
                            print(f"  → Usando K18 para acero transversal 2 (8mm): {k18_valor}")
                            as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                        elif k18_valor is not None and k18_valor < 0:
                            print(f"  → ALERTA: El valor K18 es negativo: {k18_valor}")
                            # Verificar K19 como alternativa
                            if k19_valor is not None and k19_valor > 0:
                                print(f"  → Usando K19 como alternativa: {k19_valor}")
                                as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                            else:
                                as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K18 ({k18_valor})"
                        else:
                            # Si K18 no es válido, verificar K19
                            print("  → K18 fuera de rango o inválido, verificando K19")
                            if k19_valor is not None and k19_valor > 0:
                                print(f"  → Usando K19 para acero transversal 2 (3/8\"): {k19_valor}")
                                as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                                if k19_valor <= 0.1:
                                    print("  → ADVERTENCIA: K19 menor o igual a 0.1 pero se usará de todos modos")
                            elif k19_valor is not None and k19_valor < 0:
                                print(f"  → ALERTA: El valor K19 es negativo: {k19_valor}")
                                as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K19 ({k19_valor})"
                            else:
                                print("  → K19 es None o inválido, usando valor de respaldo")
                                as_tra2_texto = "ERROR: VERIFICAR CÁLCULOS DE ACERO"
                    else:
                        # as_tra2 está en rango válido
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(as_tra2)}"
                else:
                    # Si as_tra2 es None, verificar K18 y K19 en orden
                    if k18_valor is not None and k18_valor >= 0.1 and k18_valor > 0:
                        print(f"  → as_tra2 es None, usando K18: {k18_valor}")
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                    elif k19_valor is not None and k19_valor > 0:
                        print(f"  → K18 inválido, usando K19: {k19_valor}")
                        as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                    elif k18_valor is not None and k18_valor < 0:
                        print(f"  → ALERTA: El valor K18 es negativo: {k18_valor}")
                        as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K18 ({k18_valor})"
                    elif k19_valor is not None and k19_valor < 0:
                        print(f"  → ALERTA: El valor K19 es negativo: {k19_valor}")
                        as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K19 ({k19_valor})"
                    else:
                        as_tra2_texto = "1Ø8 mm@.50"  # Valor por defecto si todo lo demás falla
                
                print("• Valores finales para bloque:")
                print(f"  → AS_LONG: {as_long_texto}")
                print(f"  → AS_TRA1: {as_tra1_texto}")
                print(f"  → AS_TRA2: {as_tra2_texto}")
                
                # Almacenar los valores finales en variables globales
                as_long_final = as_long_texto
                as_tra1_final = as_tra1_texto
                as_tra2_final = as_tra2_texto
                
                # Limpiar celdas para evitar interferencias en futuros cálculos
                
                print("=== FIN PROCESAMIENTO PRELOSA ALIGERADA - 2 SENT ===\n")
                
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
                k8_calculado = ws.range('K8').value
                k17_calculado = ws.range('K17').value
                
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
                print(f"  K8 calculado: {as_long}")
                # Verificar si hay textos vertics pero as_tra1 es 0 o nul-+
                print("\n== VALORES FINALES ==")
                print("=" * 40)
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
                
                # Para prelosas macizas, asignar valores específicos
                categoria_base = clasificar_tipo_prelosa(tipo_prelosa)
                
                if categoria_base == "MACIZA":
                    # Forzar recálculo de Excel para asegurar valores actualizados
                    print("\n=== PROCESANDO PRELOSA MACIZA ===")
                    
                    # Obtener valores de celdas K8, K9 para validación longitudinal
                    k8_valor = ws.range('K8').value
                    k9_valor = ws.range('K9').value
                    
                    # Obtener valores de celdas K18, K19 para validación transversal
                    k18_valor = ws.range('K18').value
                    k19_valor = ws.range('K19').value
                    
                    # Verificar si tenemos aceros adicionales o valores calculados manualmente
                    tiene_acero_adicional = len(textos_long_adi) > 0 or len(textos_tra_adi) > 0
                    tiene_valores_default = (len(textos_longitudinal) == 0 and len(textos_transversal) == 0 
                                        and len(textos_long_adi) == 0 and len(textos_tra_adi) == 0)
                    
                    # Determinar el tipo de prelosa para mensajes informativos
                    if tiene_acero_adicional:
                        print("PRELOSA MACIZA con ACEROS ADICIONALES - usando valores calculados previamente")
                    elif tiene_valores_default:
                        print("PRELOSA MACIZA SIN ACEROS - usando valores calculados con valores por defecto")
                    else:
                        print("PRELOSA MACIZA con ACEROS REGULARES")
                    
                    # ACERO LONGITUDINAL - Validar y seleccionar el valor adecuado con jerarquía
                    # Primero verificar si as_long está en rango válido (entre 0.4 y 0.1)
                    if as_long is None or as_long <= 0.1 or as_long < 0:
                        # Si as_long no es válido, verificar K8
                        if k8_valor is not None and k8_valor >= 0.1 and k8_valor > 0:
                            print(f"  → Usando K8 para acero longitudinal (3/8\"): {k8_valor}")
                            as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                        else:
                            # Si K8 no es válido, verificar K9
                            print("  → K8 fuera de rango o inválido, verificando K9")
                            if k9_valor is not None and k9_valor > 0:
                                print(f"  → Usando K9 para acero longitudinal (1/2\"): {k9_valor}")
                                as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                                if k9_valor <= 0.1:
                                    print("  → ADVERTENCIA: K9 menor o igual a 0.1 pero se usará de todos modos")
                            elif k9_valor is not None and k9_valor < 0:
                                print(f"  → ALERTA: El valor K9 es negativo: {k9_valor}")
                                as_long_texto = f"ERROR: VALOR NEGATIVO EN K9 ({k9_valor})"
                            else:
                                print("  → K9 es None o inválido, usando valor de respaldo")
                                as_long_texto = "ERROR: ACERO INSUFICIENTE PARA ESTA PRELOSA"
                    elif as_long < 0:
                        print(f"  → ALERTA: El valor as_long es negativo: {as_long}")
                        # Verificar K8 como alternativa
                        if k8_valor is not None and k8_valor > 0:
                            print(f"  → Usando K8 como alternativa: {k8_valor}")
                            as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                        elif k9_valor is not None and k9_valor > 0:
                            print(f"  → Usando K9 como alternativa: {k9_valor}")
                            as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                        else:
                            as_long_texto = f"ERROR: VALOR NEGATIVO en as_long ({as_long})"
                    else:
                        # as_long está en rango válido
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    
                    # ACERO TRANSVERSAL 1 - siempre a 0.28 en prelosas macizas
                    as_tra1_texto = "1Ø6 mm@.28"
                    
                    # ACERO TRANSVERSAL 2 - Validar con jerarquía: K18 > K19
                    if as_tra2 is not None:
                        if as_tra2 < 0:
                            print(f"  → ALERTA: El valor as_tra2 es negativo: {as_tra2}")
                            # Verificar K18 como alternativa
                            if k18_valor is not None and k18_valor > 0:
                                print(f"  → Usando K18 como alternativa: {k18_valor}")
                                as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                            elif k19_valor is not None and k19_valor > 0:
                                print(f"  → Usando K19 como alternativa: {k19_valor}")
                                as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                            else:
                                as_tra2_texto = f"ERROR: VALOR NEGATIVO en as_tra2 ({as_tra2})"
                        elif as_tra2 <= 0.1:
                            print("  → as_tra2 menor o igual a 0.1, verificando K18")
                            # Verificar K18
                            if k18_valor is not None and k18_valor >= 0.1 and k18_valor > 0:
                                print(f"  → Usando K18 para acero transversal 2 (8mm): {k18_valor}")
                                as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                            else:
                                # Si K18 no es válido, verificar K19
                                print("  → K18 fuera de rango o inválido, verificando K19")
                                if k19_valor is not None and k19_valor > 0:
                                    print(f"  → Usando K19 para acero transversal 2 (3/8\"): {k19_valor}")
                                    as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                                    if k19_valor <= 0.1:
                                        print("  → ADVERTENCIA: K19 menor o igual a 0.1 pero se usará de todos modos")
                                elif k19_valor is not None and k19_valor < 0:
                                    print(f"  → ALERTA: El valor K19 es negativo: {k19_valor}")
                                    as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K19 ({k19_valor})"
                                else:
                                    print("  → K19 es None o inválido, usando valor de respaldo")
                                    as_tra2_texto = "ERROR: VERIFICAR CÁLCULOS DE ACERO"
                        else:
                            # as_tra2 está en rango válido
                            as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(as_tra2)}"
                    else:
                        # Si as_tra2 es None, verificar K18 y K19 en orden
                        if k18_valor is not None and k18_valor >= 0.1 and k18_valor > 0:
                            print(f"  → as_tra2 es None, usando K18: {k18_valor}")
                            as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                        elif k19_valor is not None and k19_valor > 0:
                            print(f"  → K18 inválido, usando K19: {k19_valor}")
                            as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                        elif k18_valor is not None and k18_valor < 0:
                            print(f"  → ALERTA: El valor K18 es negativo: {k18_valor}")
                            as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K18 ({k18_valor})"
                        elif k19_valor is not None and k19_valor < 0:
                            print(f"  → ALERTA: El valor K19 es negativo: {k19_valor}")
                            as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K19 ({k19_valor})"
                        else:
                            as_tra2_texto = None
                    
                    print("• Valores finales para bloque:")
                    print(f"  → AS_LONG: {as_long_texto}")
                    print(f"  → AS_TRA1: {as_tra1_texto}")
                    if as_tra2_texto:
                        print(f"  → AS_TRA2: {as_tra2_texto}")
                    
                    print("=== FIN PROCESAMIENTO PRELOSA MACIZA ===\n")
                    
                    # Limpiar celdas para evitar interferencias
                    print("limpiar celdas")
                    for celda in ['G5', 'G6', 'G7', 'G15', 'G16', 'G17']:
                        ws.range(celda).value = 0
                
                elif categoria_base == "ALIGERADA":
                    print("\n=== PROCESANDO PRELOSA ALIGERADA ===")
                    # Obtener valores de celdas K8, K9, K10 para validación
                    k8_valor = ws.range('K8').value
                    k9_valor = ws.range('K9').value
                    k10_valor = ws.range('K10').value
                    
                    # ACERO LONGITUDINAL - Validar y seleccionar el valor adecuado con jerarquía
                    if as_long is None or as_long < 0.1:
                        print("  → as_long menor a 0.1, verificando K8")
                        # Verificar K8 primero
                        if k8_valor is not None and k8_valor >= 0.1 and k8_valor > 0:
                            print(f"  → Usando K8 para acero longitudinal (3/8\"): {k8_valor}")
                            as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                        elif k8_valor is not None and k8_valor < 0:
                            print(f"  → ALERTA: El valor K8 es negativo: {k8_valor}")
                            # Continuar con K9
                            if k9_valor is not None and k9_valor > 0:
                                print(f"  → Usando K9 como alternativa: {k9_valor}")
                                as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                            else:
                                as_long_texto = f"ERROR: VALOR NEGATIVO EN K8 ({k8_valor})"
                        else:
                            # Si K8 no es válido, verificar K9
                            print("  → K8 fuera de rango o inválido, verificando K9")
                            if k9_valor is not None and k9_valor >= 0.1 and k9_valor > 0:
                                print(f"  → Usando K9 para acero longitudinal (1/2\"): {k9_valor}")
                                as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                            elif k9_valor is not None and k9_valor < 0:
                                print(f"  → ALERTA: El valor K9 es negativo: {k9_valor}")
                                # Verificar K10 como alternativa
                                if k10_valor is not None and k10_valor > 0:
                                    print(f"  → Usando K10 como alternativa: {k10_valor}")
                                    as_long_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k10_valor)}"
                                else:
                                    as_long_texto = f"ERROR: VALOR NEGATIVO EN K9 ({k9_valor})"
                            else:
                                # Si K9 no es válido, verificar K10
                                print("  → K9 fuera de rango o inválido, verificando K10")
                                if k10_valor is not None and k10_valor >= 0.1 and k10_valor > 0:
                                    print(f"  → Usando K10 para acero longitudinal (8mm): {k10_valor}")
                                    as_long_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k10_valor)}"
                                elif k10_valor is not None and k10_valor < 0:
                                    print(f"  → ALERTA: El valor K10 es negativo: {k10_valor}")
                                    as_long_texto = f"ERROR: VALOR NEGATIVO EN K10 ({k10_valor})"
                                else:
                                    print("  → Todas las opciones son inválidas o menores a 0.1")
                                    as_long_texto = "ERROR: ACERO INSUFICIENTE PARA ESTA PRELOSA"
                    elif as_long < 0:
                        print(f"  → ALERTA: El valor as_long es negativo: {as_long}")
                        # Verificar alternativas en orden: K8, K9, K10
                        if k8_valor is not None and k8_valor > 0:
                            print(f"  → Usando K8 como alternativa: {k8_valor}")
                            as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                        elif k9_valor is not None and k9_valor > 0:
                            print(f"  → Usando K9 como alternativa: {k9_valor}")
                            as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                        elif k10_valor is not None and k10_valor > 0:
                            print(f"  → Usando K10 como alternativa: {k10_valor}")
                            as_long_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k10_valor)}"
                        else:
                            as_long_texto = f"ERROR: VALOR NEGATIVO en as_long ({as_long})"
                    else:
                        # as_long está en rango válido
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(as_long)}"
                    
                    # Para acero vertical (AS_TRA1) - siempre fijo en aligeradas
                    as_tra1_texto = "1Ø6 mm@.50"
                    
                    # Para AS_TRA2 - siempre fijo en aligeradas
                    as_tra2_texto = "1Ø8 mm@.50"
                    
                    print("• Valores finales para bloque:")
                    print(f"  → AS_LONG: {as_long_texto}")
                    print(f"  → AS_TRA1: {as_tra1_texto}")
                    print(f"  → AS_TRA2: {as_tra2_texto}")
                    
                    print("=== FIN PROCESAMIENTO PRELOSA ALIGERADA ===\n")
                    
                    # Limpiar celdas para evitar interferencias
                    print("limpiar celdas")
                    for celda in ['G5', 'G6', 'G7', 'G15', 'G16', 'G17']:
                        ws.range(celda).value = 0

                elif categoria_base == "ALIGERADA_2SENT":
                    print("\n=== PROCESANDO PRELOSA ALIGERADA - 2 SENT ===")
                    
                    # Obtener valores de celdas para validación
                    k8_valor = ws.range('K8').value
                    k9_valor = ws.range('K9').value
                    k18_valor = ws.range('K18').value
                    k19_valor = ws.range('K19').value
                    
                    print(f"• Valores calculados: K8={k8_valor}, K9={k9_valor}")
                    print(f"• Valores calculados: K18={k18_valor}, K19={k19_valor}")
                    
                    # ACERO LONGITUDINAL - Validar y seleccionar el valor adecuado con jerarquía
                    if k8_valor is None or k8_valor <= 0.1:  # Usar directamente k8_valor en lugar de as_long
                        print("  → as_long menor o igual a 0.1, verificando K8")
                        # Si K8 no es válido, verificar K9
                        print("  → K8 fuera de rango o inválido, verificando K9")
                        if k9_valor is not None and k9_valor > 0:
                            print(f"  → Usando K9 para acero longitudinal (1/2\"): {k9_valor}")
                            as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                            if k9_valor <= 0.1:
                                print("  → ADVERTENCIA: K9 menor o igual a 0.1 pero se usará de todos modos")
                        elif k9_valor is not None and k9_valor < 0:
                            print(f"  → ALERTA: El valor K9 es negativo: {k9_valor}")
                            as_long_texto = f"ERROR: VALOR NEGATIVO EN K9 ({k9_valor})"
                        else:
                            print("  → K9 es None o inválido, usando valor de respaldo")
                            as_long_texto = "ERROR: ACERO INSUFICIENTE PARA ESTA PRELOSA"
                    elif k8_valor < 0:  # Usar directamente k8_valor
                        print(f"  → ALERTA: El valor K8 es negativo: {k8_valor}")
                        # Verificar alternativas en orden: K8, K9
                        if k9_valor is not None and k9_valor > 0:
                            print(f"  → Usando K9 como alternativa: {k9_valor}")
                            as_long_texto = f"1Ø1/2\"@.{formatear_valor_espaciamiento(k9_valor)}"
                        else:
                            as_long_texto = f"ERROR: VALOR NEGATIVO en K8 ({k8_valor})"
                    else:
                        # K8 está en rango válido
                        as_long_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k8_valor)}"
                    
                    # Para acero vertical (AS_TRA1) - siempre a 0.28 en prelosas aligeradas 2 sent
                    as_tra1_texto = "1Ø6 mm@.28"
                    
                    # ACERO TRANSVERSAL 2 - Validar con jerarquía: K18 > K19
                    if k18_valor is None or k18_valor <= 0.1:  # Usar directamente k18_valor
                        print("  → as_tra2 menor o igual a 0.1, verificando K18")
                        print("  → K18 fuera de rango o inválido, verificando K19")
                        if k19_valor is not None and k19_valor > 0:
                            print(f"  → Usando K19 para acero transversal 2 (3/8\"): {k19_valor}")
                            as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                            if k19_valor <= 0.1:
                                print("  → ADVERTENCIA: K19 menor o igual a 0.1 pero se usará de todos modos")
                        elif k19_valor is not None and k19_valor < 0:
                            print(f"  → ALERTA: El valor K19 es negativo: {k19_valor}")
                            as_tra2_texto = f"ERROR: VALOR NEGATIVO EN K19 ({k19_valor})"
                        else:
                            print("  → K19 es None o inválido, usando valor de respaldo")
                            as_tra2_texto = "ERROR: VERIFICAR CÁLCULOS DE ACERO"
                    elif k18_valor < 0:  # Usar directamente k18_valor
                        print(f"  → ALERTA: El valor K18 es negativo: {k18_valor}")
                        # Verificar K19 como alternativa
                        if k19_valor is not None and k19_valor > 0:
                            print(f"  → Usando K19 como alternativa: {k19_valor}")
                            as_tra2_texto = f"1Ø3/8\"@.{formatear_valor_espaciamiento(k19_valor)}"
                        else:
                            as_tra2_texto = f"ERROR: VALOR NEGATIVO en K18 ({k18_valor})"
                    else:
                        # K18 está en rango válido
                        as_tra2_texto = f"1Ø8 mm@.{formatear_valor_espaciamiento(k18_valor)}"
                    
                    print("• Valores finales para bloque:")
                    print(f"  → AS_LONG: {as_long_texto}")
                    print(f"  → AS_TRA1: {as_tra1_texto}")
                    print(f"  → AS_TRA2: {as_tra2_texto}")
                    
                    # Almacenar los valores finales en variables globales para que el bloque los use
                    as_long_final = as_long_texto
                    as_tra1_final = as_tra1_texto
                    as_tra2_final = as_tra2_texto
                    
                    print("=== FIN PROCESAMIENTO PRELOSA ALIGERADA - 2 SENT ===\n")
                    
                    # Limpiar celdas para evitar interferencias
                    print("limpiar celdas")
                    for celda in ['G5', 'G6', 'G7', 'G15', 'G16', 'G17']:
                        ws.range(celda).value = 0
                
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
                                
                print(f" ==== Valores formateados para inserción en bloque: ===")
                print(f"    AS_LONG: {as_long_texto}")
                print(f"    AS_TRA1: {as_tra1_texto}")
                if as_tra2_texto:
                    print(f"    AS_TRA2: {as_tra2_texto}")
                

                # Filtrar polilíneas longitudinales y adicionales
                polilineas_longitudinal = [p for p in polilineas_dentro if "LONGITUDINAL" in p.dxf.layer.upper() and "ADI" not in p.dxf.layer.upper()]
                polilineas_long_adi = [p for p in polilineas_dentro if "LONG ADI" in p.dxf.layer.upper()]

                # Calcular la orientación considerando ambos tipos de acero
                angulo_rotacion = calcular_orientacion_prelosa(vertices, polilineas_longitudinal, polilineas_long_adi)
                
                # Crear una copia de la definición del bloque con la orientación calculada
                # En la función procesar_prelosa, justo antes de llamar a insertar_bloque_acero:

                # Calcular las dimensiones de la polilínea
                vertices = polilinea.get_points('xy')
                xs = [v[0] for v in vertices]
                ys = [v[1] for v in vertices]
                min_x, max_x = min(xs), max(xs)
                min_y, max_y = min(ys), max(ys)
                ancho_polilinea = max_x - min_x
                alto_polilinea = max_y - min_y

                # Agregar las dimensiones a la definición del bloque
                definicion_bloque_orientada = definicion_bloque.copy()
                definicion_bloque_orientada['rotation'] = angulo_rotacion
                definicion_bloque_orientada['polilinea_ancho'] = ancho_polilinea
                definicion_bloque_orientada['polilinea_alto'] = alto_polilinea

                # Insertar bloque con los valores formateados y la orientación correcta, ajustado al tamaño de la polilínea
                bloque = insertar_bloque_acero(msp, definicion_bloque_orientada, centro_prelosa, as_long_texto, as_tra1_texto, as_tra2_texto)
                
                if bloque:
                    total_bloques += 1
                    print(f"{tipo_prelosa} CONCLUIDA CON EXITO ===============================")
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
       
        for tipo, polilineas in polilineas_por_tipo.items():
            print(f"Encontradas {len(polilineas)} polilíneas de tipo {tipo}")

        # Procesar cada tipo usando polilineas_por_tipo
        for tipo_prelosa, polilineas in polilineas_por_tipo.items():
            print(f"Procesando {len(polilineas)} polilíneas de tipo {tipo_prelosa}")
            
            # Procesar cada polilínea según el tipo
            for idx, polilinea in enumerate(polilineas):
                procesar_prelosa(polilinea, tipo_prelosa, idx)
                
        
        # Cerrar Excel y guardar DXF
        try:
            wb.save()
            wb.close()
            app.quit()
        except:
            print("Error al cerrar Excel, continuando...")
        
        capas_acero = ["ACERO LONGITUDINAL", "ACERO TRANSVERSAL", "ACERO LONG ADI",
        "BD-ACERO LONGITUDINAL", "BD-ACERO TRANSVERSAL", "ACERO TRA ADI"]    
        
        eliminar_entidades_por_capa(doc, capas_acero)

        desbloquear_capa_acero_positivo(doc)

        
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