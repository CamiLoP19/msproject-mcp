import os
from datetime import datetime
import pythoncom
import win32com.client

OUTPUT_PATH = r"C:\Users\janer\Downloads\Taller_UPC_Resuelto.mpp"

TASKS_DATA = [
    # Hito inicial requerido por la guía (Punto 13)
    (1, "Inicio del Proyecto", "0d", True, None),
    
    # Nivel 1: Proyecto Resumen
    (1, "Plan de desarrollo de software", None, False, None),
    
    # Fase 1: Alcance
    (2, "Alcance", None, False, None),
    (3, "Alcance del proyecto", "4h", False, "Luis Perez"),
    (3, "Patrocinio seguro de proyectos", "1d", False, "Luis Perez"),
    (3, "Recursos preliminares", "1d", False, "Luis Perez"),
    (3, "Recursos básicos seguros", "1d", False, "Luis Perez"),
    (3, "Alcance completo", "0d", True, None),
    
    # Fase 2: Requisitos de análisis/software
    (2, "Requisitos de análisis/software", None, False, None),
    (3, "Análisis de necesidades", "5d", False, "Maria Claudia Rincon"),
    (3, "Borrador de especificaciones preliminares de software", "3d", False, "Maria Claudia Rincon"),
    (3, "Presupuesto preliminar", "2d", False, "Luis Perez"),
    (3, "Revisión de especificaciones/el presupuesto del software con el equipo", "4h", False, "Luis Perez,Maria Claudia Rincon"),
    (3, "Comentarios sobre las especificaciones del software", "1d", False, "Maria Claudia Rincon"),
    (3, "Cronograma de entrega", "1d", False, "Luis Perez"),
    (3, "Aprocaciones para proceder (concepto, cronograma, presupuesto)", "4h", False, "Luis Perez"),
    (3, "Asegure los recursos necesarios", "1d", False, "Luis Perez"),
    (3, "Análisis completo", "0d", True, None),
    
    # Fase 3: Diseño
    (2, "Diseño", None, False, None),
    (3, "Revisión de las especificaciones preliminares del software", "2d", False, "Donald Trump"),
    (3, "Especificaciones funcionales.", "5d", False, "Donald Trump"),
    (3, "Prototipo basado en especificaciones funcionales.", "4d", False, "Donald Trump"),
    (3, "Especificaciones funcionales", "2d", False, "Donald Trump"),
    (3, "Incorporar comentarios en las especificaciones funcionales", "1d", False, "Donald Trump"),
    (3, "Obtener aprobación para continuar", "4h", False, "Luis Perez"),
    (3, "Diseño completo", "0d", True, None),
    
    # Fase 4: Desarrollo
    (2, "Desarrollo", None, False, None),
    (3, "Revisión de especificaciones funcionales", "1d", False, "Sofía Vergara"),
    (3, "Identificar parámetros de diseño modulares/por niveles", "1d", False, "Sofía Vergara"),
    (3, "Asignar personal de desarrollo", "1d", False, "Luis Perez"),
    (3, "Desarrollar código", "15d", False, "Sofía Vergara,Computadores[2]"),
    (3, "Pruebas de desarrollador (depuración primaria)", "15d", False, "Sofía Vergara,Computadores[2]"),
    (3, "Desarrollo completo", "0d", True, None),
    
    # Fase 5: Pruebas
    (2, "Pruebas", None, False, None),
    (3, "Desarrollar planes de prueba de unidad usando especificaciones de producto", "4d", False, "Maria Claudia Rincon"),
    (3, "Desarrolle planes de prueba de integración usando especificaciones de productos", "4d", False, "Maria Claudia Rincon"),
    (3, "Examen de la unidad", None, False, None),
    (4, "Revisar código modular", "5d", False, "Sofía Vergara"),
    (4, "Pruebe los módulos de los componentes según las especificaciones del producto", "2d", False, "Sofía Vergara"),
    (4, "Identificar anomalías en las especificaciones del producto.", "3d", False, "Maria Claudia Rincon"),
    (4, "Modificar código", "3d", False, "Sofía Vergara"),
    (4, "Vuelva a probar el código modificado", "2d", False, "Sofía Vergara"),
    (4, "Prueba unitaria completa", "0d", True, None),
    (3, "Pruebas de integración", None, False, None),
    (4, "Integración del módulo de prueba", "5d", False, "Sofía Vergara,Servidor"),
    (4, "Identificación de anomalías en las especificaciones.", "2d", False, "Maria Claudia Rincon"),
    (4, "Modificación del código", "3d", False, "Sofía Vergara"),
    (4, "Vuelva a probar el código modificado", "2d", False, "Sofía Vergara"),
    (4, "Prueba de integración completa", "0d", True, None),
    
    # Fase 6: Capacitación
    (2, "Capacitación", None, False, None),
    (3, "Desarrollar especificaciones de capacitación para usuarios finales.", "3d", False, "Maria Claudia Rincon"),
    (3, "Desarrollar especificaciones de capacitación para el personal de soporte de la mesa de ayuda", "3d", False, "Maria Claudia Rincon"),
    (3, "Identificar la metodología de entrega de capacitación (capacitación basada en computadora, aula, etc.)", "2d", False, "Luis Perez"),
    (3, "Desarrollar materiales de capacitación.", "3s", False, "Maria Claudia Rincon"),
    (3, "Realizar un estudio de usabilidad de la formación", "4d", False, "Donald Trump"),
    (3, "Finalizar los materiales de capacitación", "3d", False, "Maria Claudia Rincon"),
    (3, "Desarrollar un mecanismo de entrega de capacitación", "2d", False, "Luis Perez"),
    (3, "Materiales de capacitación completos", "0d", True, None),
    
    # Fase 7: Documentación
    (2, "Documentación", None, False, None),
    (3, "Especificación de ayuda", "1d", False, "Donald Trump"),
    (3, "Sistema de ayuda", "3s", False, "Donald Trump"),
    (3, "Revisar la documentación de ayuda", "3d", False, "Maria Claudia Rincon"),
    (3, "Incorporar comentarios de la documentación de ayuda", "2d", False, "Donald Trump"),
    (3, "Desarrollar especificaciones de manuales de usuario.", "2d", False, "Donald Trump"),
    (3, "Desarrollar manuales de usuario.", "3s", False, "Donald Trump"),
    (3, "Revisar toda la documentación del usuario", "2d", False, "Maria Claudia Rincon"),
    (3, "Incorporar comentarios de la documentación del usuario", "2d", False, "Donald Trump"),
    (3, "Documentación completa", "0d", True, None),
    
    # Fase 8: Piloto
    (2, "Piloto", None, False, None),
    (3, "Identificar grupo de prueba", "1d", False, "Luis Perez"),
    (3, "Desarrollo de un mecanismo de entrega de software", "1d", False, "Sofía Vergara"),
    (3, "Instalar/implementar software", "1d", False, "Sofía Vergara,Ponchadora,Cable[10]"),
    (3, "Obtención de comentarios de los usuarios", "1s", False, "Maria Claudia Rincon"),
    (3, "Evaluación de la información de las pruebas", "1d", False, "Maria Claudia Rincon"),
    (3, "Piloto completo", "0d", True, None),
    
    # Fase 9: Despliegue
    (2, "Despliegue", None, False, None),
    (3, "Determinar la estrategia de implementación final", "1d", False, "Luis Perez"),
    (3, "Desarrollar metodología de implementación", "1d", False, "Luis Perez"),
    (3, "Recursos de implementación seguros", "1d", False, "Luis Perez"),
    (3, "Capacitar al personal de apoyo", "1d", False, "Maria Claudia Rincon"),
    (3, "Implementar software", "1d", False, "Sofía Vergara,Servidor,Cable[40]"),
    (3, "Implementación completa", "0d", True, None),
    
    # Fase 10: Revisión posterior a la implementación
    (2, "Revisión posterior a la implementación", None, False, None),
    (3, "Documentar lecciones aprendidas", "1d", False, "Luis Perez"),
    (3, "Distribuir a los miembros del equipo", "1d", False, "Luis Perez"),
    (3, "Crear equipo de mantenimiento de software.", "1d", False, "Luis Perez"),
    (3, "Revisión posterior a la implementación completa", "0d", True, None),
    
    # Hito Final (Punto 13)
    (1, "Plantilla de desarrollo de software completa", "0d", True, None),
    (1, "Fin del Proyecto", "0d", True, None),
]

# Tabla exacta del taller:
# Nombre, Tipo, Cargo/Grupo, Disponibilidad (MaxUnits), Vr. Hora / Costo Unit
WORK_RESOURCES = [
    ("Sofía Vergara", "Programador", 1.0, "35000/h"),
    ("Donald Trump", "Diseñador Gráfico", 1.0, "25000/h"),
    ("Maria Claudia Rincon", "Analista", 0.5, "30000/h"),
    ("Luis Perez", "Gerente", 1.0, "45000/h"),
]

MATERIAL_RESOURCES = [
    ("Servidor", "Unidad", "50000"),
    ("Computadores", "Unidad", "40000"),
    ("Ponchadora", "Unidad", "200000"),
    ("Cable", "Metro", "20000"),
]

def generate():
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("MSProject.Application")
    app.Visible = True
    app.DisplayAlerts = False
    
    p = app.Projects.Add(False)
    p.Title = "Plan de Desarrollo de Software - Taller UPC"
    p.ProjectStart = datetime(datetime.now().year, 10, 1, 8, 0)
    
    # 1. Configurar recursos de Tipo Trabajo
    res_dict = {}
    for name, cargo, max_units, rate in WORK_RESOURCES:
        r = p.Resources.Add(name)
        r.Type = 0 # pjResourceTypeWork
        r.Group = cargo
        r.MaxUnits = max_units
        try:
            r.StandardRate = rate
        except Exception:
            pass
        res_dict[name] = r

    # 2. Configurar recursos de Tipo Material
    for name, unit_label, cost_unit in MATERIAL_RESOURCES:
        r = p.Resources.Add(name)
        r.Type = 1 # pjResourceTypeMaterial
        r.MaterialLabel = unit_label
        try:
            r.StandardRate = cost_unit
        except Exception:
            pass
        res_dict[name] = r

    # 3. Crear tareas y asignar recursos
    for level, name, duration, milestone, resources in TASKS_DATA:
        t = p.Tasks.Add(name)
        t.OutlineLevel = level
        # Configurar autoprogramación y quitar estimado
        try:
            t.Manual = False
        except Exception:
            pass
        t.Estimated = False
        
        if milestone:
            t.Duration = 0
            t.Milestone = True
            t.Estimated = False
        elif duration:
            try:
                t.Duration = app.DurationValue(duration)
            except Exception:
                val = duration.replace("w", "s")
                t.Duration = app.DurationValue(val)
            t.Estimated = False
            
        if resources:
            for item in resources.split(","):
                item = item.strip()
                # Verificar si viene con unidades ej. Cable[40]
                units = 1.0
                clean_name = item
                if "[" in item and "]" in item:
                    clean_name = item.split("[")[0].strip()
                    try:
                        units = float(item.split("[")[1].replace("]", "").strip())
                    except Exception:
                        units = 1.0
                
                r_obj = res_dict.get(clean_name)
                if r_obj:
                    try:
                        t.Assignments.Add(t.ID, r_obj.ID, units)
                    except Exception:
                        pass

    # 4. Vincular tareas
    task_by_name = {t.Name: t for t in p.Tasks}
    
    # Hito de inicio a primera tarea
    if "Inicio del Proyecto" in task_by_name and "Alcance del proyecto" in task_by_name:
        task_by_name["Alcance del proyecto"].Predecessors = str(task_by_name["Inicio del Proyecto"].ID)
        
    # Secuencia dentro de Alcance:
    alcance_seq = ["Alcance del proyecto", "Patrocinio seguro de proyectos", "Recursos preliminares", "Recursos básicos seguros", "Alcance completo"]
    for i in range(len(alcance_seq) - 1):
        t1 = task_by_name.get(alcance_seq[i])
        t2 = task_by_name.get(alcance_seq[i+1])
        if t1 and t2: t2.Predecessors = str(t1.ID)

    # Punto 11 de la guía: Actividad 1 de Alcance con Actividad 1 de Requisitos
    t_alcance_1 = task_by_name.get("Alcance del proyecto")
    t_req_1 = task_by_name.get("Análisis de necesidades")
    if t_alcance_1 and t_req_1:
        t_req_1.Predecessors = str(t_alcance_1.ID)

    # Secuencia dentro de Requisitos:
    req_seq = [
        "Análisis de necesidades", "Borrador de especificaciones preliminares de software",
        "Presupuesto preliminar", "Revisión de especificaciones/el presupuesto del software con el equipo",
        "Comentarios sobre las especificaciones del software", "Cronograma de entrega",
        "Aprocaciones para proceder (concepto, cronograma, presupuesto)", "Asegure los recursos necesarios",
        "Análisis completo"
    ]
    for i in range(len(req_seq) - 1):
        t1 = task_by_name.get(req_seq[i])
        t2 = task_by_name.get(req_seq[i+1])
        if t1 and t2: t2.Predecessors = str(t1.ID)

    t_analisis_fin = task_by_name.get("Análisis completo")
    t_dis_1 = task_by_name.get("Revisión de las especificaciones preliminares del software")
    if t_analisis_fin and t_dis_1:
        t_dis_1.Predecessors = str(t_analisis_fin.ID)

    dis_seq = [
        "Revisión de las especificaciones preliminares del software", "Especificaciones funcionales.",
        "Prototipo basado en especificaciones funcionales.", "Especificaciones funcionales",
        "Incorporar comentarios en las especificaciones funcionales", "Obtener aprobación para continuar",
        "Diseño completo"
    ]
    for i in range(len(dis_seq) - 1):
        t1 = task_by_name.get(dis_seq[i])
        t2 = task_by_name.get(dis_seq[i+1])
        if t1 and t2: t2.Predecessors = str(t1.ID)

    t_dis_fin = task_by_name.get("Diseño completo")
    t_dev_1 = task_by_name.get("Revisión de especificaciones funcionales")
    if t_dis_fin and t_dev_1:
        t_dev_1.Predecessors = str(t_dis_fin.ID)

    dev_seq = [
        "Revisión de especificaciones funcionales", "Identificar parámetros de diseño modulares/por niveles",
        "Asignar personal de desarrollo", "Desarrollar código", "Pruebas de desarrollador (depuración primaria)",
        "Desarrollo completo"
    ]
    for i in range(len(dev_seq) - 1):
        t1 = task_by_name.get(dev_seq[i])
        t2 = task_by_name.get(dev_seq[i+1])
        if t1 and t2: t2.Predecessors = str(t1.ID)

    t_dev_fin = task_by_name.get("Desarrollo completo")
    t_pruebas_1 = task_by_name.get("Desarrollar planes de prueba de unidad usando especificaciones de producto")
    if t_dev_fin and t_pruebas_1:
        t_pruebas_1.Predecessors = str(t_dev_fin.ID)

    # Hito final
    t_ultimo_hito = task_by_name.get("Plantilla de desarrollo de software completa")
    t_fin_proy = task_by_name.get("Fin del Proyecto")
    if t_ultimo_hito and t_fin_proy:
        t_fin_proy.Predecessors = str(t_ultimo_hito.ID)

    # 5. Guardar Línea Base
    try:
        app.BaselineSave(True)
    except Exception:
        pass

    # 6. Guardar archivo
    if os.path.exists(OUTPUT_PATH):
        try:
            os.remove(OUTPUT_PATH)
        except Exception:
            pass
            
    p.Activate()
    app.FileSaveAs(OUTPUT_PATH)
    print("PROYECTO ACTUALIZADO CON ÉXITO")

if __name__ == "__main__":
    generate()
