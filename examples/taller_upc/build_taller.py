import os
import sys
from datetime import datetime
import pythoncom
import win32com.client

OUTPUT_PATH = r"C:\Users\janer\Downloads\Taller_UPC_Resuelto.mpp"

# Definición de la EDT y Actividades del Taller
# outline_level, name, duration, milestone, resources, predecessors
TASKS_DATA = [
    # Hito inicial requerido por la guía
    (1, "Inicio del Proyecto", "0d", True, None, None),
    
    # Nivel 1: Proyecto Resumen
    (1, "Plan de desarrollo de software", None, False, None, None),
    
    # Fase 1: Alcance
    (2, "Alcance", None, False, None, None),
    (3, "Alcance del proyecto", "4h", False, "Management", "1"),
    (3, "Patrocinio seguro de proyectos", "1d", False, "Management", None), # enlazado secuencial o al inicio
    (3, "Recursos preliminares", "1d", False, "Project Manager", None),
    (3, "Recursos básicos seguros", "1d", False, "Project Manager", None),
    (3, "Alcance completo", "0d", True, None, None),
    
    # Fase 2: Requisitos de análisis/software
    (2, "Requisitos de análisis/software", None, False, None, None),
    (3, "Análisis de necesidades", "5d", False, "Analyst", None),
    (3, "Borrador de especificaciones preliminares de software", "3d", False, "Analyst", None),
    (3, "Presupuesto preliminar", "2d", False, "Project Manager", None),
    (3, "Revisión de especificaciones/el presupuesto del software con el equipo", "4h", False, "Project Manager,Analyst", None),
    (3, "Comentarios sobre las especificaciones del software", "1d", False, "Analyst", None),
    (3, "Cronograma de entrega", "1d", False, "Project Manager", None),
    (3, "Aprobaciones para proceder (concepto, cronograma, presupuesto)", "4h", False, "Management,Project Manager", None),
    (3, "Asegure los recursos necesarios", "1d", False, "Project Manager", None),
    (3, "Análisis completo", "0d", True, None, None),
    
    # Fase 3: Diseño
    (2, "Diseño", None, False, None, None),
    (3, "Revisión de las especificaciones preliminares del software", "2d", False, "Analyst", None),
    (3, "Especificaciones funcionales.", "5d", False, "Analyst", None),
    (3, "Prototipo basado en especificaciones funcionales.", "4d", False, "Analyst", None),
    (3, "Especificaciones funcionales", "2d", False, "Management", None),
    (3, "Incorporar comentarios en las especificaciones funcionales", "1d", False, "Management", None),
    (3, "Obtener aprobación para continuar", "4h", False, "Management,Project Manager", None),
    (3, "Diseño completo", "0d", True, None, None),
    
    # Fase 4: Desarrollo
    (2, "Desarrollo", None, False, None, None),
    (3, "Revisión de especificaciones funcionales", "1d", False, "Developer", None),
    (3, "Identificar parámetros de diseño modulares/por niveles", "1d", False, "Developer", None),
    (3, "Asignar personal de desarrollo", "1d", False, "Developer", None),
    (3, "Desarrollar código", "15d", False, "Developer", None),
    (3, "Pruebas de desarrollador (depuración primaria)", "15d", False, "Developer", None),
    (3, "Desarrollo completo", "0d", True, None, None),
    
    # Fase 5: Pruebas
    (2, "Pruebas", None, False, None, None),
    (3, "Desarrollar planes de prueba de unidad usando especificaciones de producto", "4d", False, "Testers", None),
    (3, "Desarrolle planes de prueba de integración usando especificaciones de productos", "4d", False, "Testers", None),
    (3, "Examen de la unidad", None, False, None, None), # Subfase
    (4, "Revisar código modular", "5d", False, "Testers", None),
    (4, "Pruebe los módulos de los componentes según las especificaciones del producto", "2d", False, "Testers", None),
    (4, "Identificar anomalías en las especificaciones del producto.", "3d", False, "Testers", None),
    (4, "Modificar código", "3d", False, "Testers", None),
    (4, "Vuelva a probar el código modificado", "2d", False, "Testers", None),
    (4, "Prueba unitaria completa", "0d", True, None, None),
    (3, "Pruebas de integración", None, False, None, None), # Subfase
    (4, "Integración del módulo de prueba", "5d", False, "Testers", None),
    (4, "Identificación de anomalías en las especificaciones.", "2d", False, "Testers", None),
    (4, "Modificación del código", "3d", False, "Testers", None),
    (4, "Vuelva a probar el código modificado", "2d", False, "Testers", None),
    (4, "Prueba de integración completa", "0d", True, None, None),
    
    # Fase 6: Capacitación
    (2, "Capacitación", None, False, None, None),
    (3, "Desarrollar especificaciones de capacitación para usuarios finales.", "3d", False, "Trainers", None),
    (3, "Desarrollar especificaciones de capacitación para el personal de soporte de la mesa de ayuda", "3d", False, "Trainers", None),
    (3, "Identificar la metodología de entrega de capacitación (capacitación basada en computadora, aula, etc.)", "2d", False, "Trainers", None),
    (3, "Desarrollar materiales de capacitación.", "3s", False, "Trainers", None), # 3 semanas
    (3, "Realizar un estudio de usabilidad de la formación", "4d", False, "Trainers", None),
    (3, "Finalizar los materiales de capacitación", "3d", False, "Trainers", None),
    (3, "Desarrollar un mecanismo de entrega de capacitación", "2d", False, "Trainers", None),
    (3, "Materiales de capacitación completos", "0d", True, None, None),
    
    # Fase 7: Documentación
    (2, "Documentación", None, False, None, None),
    (3, "Especificación de ayuda", "1d", False, "Technical Communicators", None),
    (3, "Sistema de ayuda", "3s", False, "Technical Communicators", None),
    (3, "Revisar la documentación de ayuda", "3d", False, "Technical Communicators", None),
    (3, "Incorporar comentarios de la documentación de ayuda", "2d", False, "Technical Communicators", None),
    (3, "Desarrollar especificaciones de manuales de usuario.", "2d", False, "Technical Communicators", None),
    (3, "Desarrollar manuales de usuario.", "3s", False, "Technical Communicators", None),
    (3, "Revisar toda la documentación del usuario", "2d", False, "Technical Communicators", None),
    (3, "Incorporar comentarios de la documentación del usuario", "2d", False, "Technical Communicators", None),
    (3, "Documentación completa", "0d", True, None, None),
    
    # Fase 8: Piloto
    (2, "Piloto", None, False, None, None),
    (3, "Identificar grupo de prueba", "1d", False, "Project Manager", None),
    (3, "Desarrollo de un mecanismo de entrega de software", "1d", False, None, None),
    (3, "Instalar/implementar software", "1d", False, "Deployment Team", None),
    (3, "Obtención de comentarios de los usuarios", "1s", False, "Deployment Team", None),
    (3, "Evaluación de la información de las pruebas", "1d", False, "Deployment Team", None),
    (3, "Piloto completo", "0d", True, None, None),
    
    # Fase 9: Despliegue
    (2, "Despliegue", None, False, None, None),
    (3, "Determinar la estrategia de implementación final", "1d", False, "Deployment Team", None),
    (3, "Desarrollar metodología de implementación", "1d", False, "Deployment Team", None),
    (3, "Recursos de implementación seguros", "1d", False, "Deployment Team", None),
    (3, "Capacitar al personal de apoyo", "1d", False, "Deployment Team", None),
    (3, "Implementar software", "1d", False, "Deployment Team", None),
    (3, "Implementación completa", "0d", True, None, None),
    
    # Fase 10: Revisión posterior a la implementación
    (2, "Revisión posterior a la implementación", None, False, None, None),
    (3, "Documentar lecciones aprendidas", "1d", False, "Project Manager", None),
    (3, "Distribuir a los miembros del equipo", "1d", False, "Project Manager", None),
    (3, "Crear equipo de mantenimiento de software.", "1d", False, "Project Manager", None),
    (3, "Revisión posterior a la implementación completa", "0d", True, None, None),
    
    # Hito Final
    (1, "Plantilla de desarrollo de software completa", "0d", True, None, None),
    (1, "Fin del Proyecto", "0d", True, None, None),
]

RESOURCES = [
    ("Management", "work"),
    ("Project Manager", "work"),
    ("Analyst", "work"),
    ("Developer", "work"),
    ("Testers", "work"),
    ("Trainers", "work"),
    ("Technical Communicators", "work"),
    ("Deployment Team", "work"),
]

def build_project():
    pythoncom.CoInitialize()
    print("Iniciando Microsoft Project...")
    app = win32com.client.Dispatch("MSProject.Application")
    app.Visible = True
    app.DisplayAlerts = False
    
    # 1. Crear nuevo proyecto
    p = app.Projects.Add(False)
    p.Title = "Plan de Desarrollo de Software - Taller UPC"
    
    # Fecha de inicio: 1 de octubre
    # Se usa año corriente o especificado en la fecha
    current_year = datetime.now().year
    p.ProjectStart = datetime(current_year, 10, 1, 8, 0)
    print(f"Proyecto creado con inicio: {p.ProjectStart}")
    
    # 2. Calendario UPC
    try:
        app.BaseCalendarCreate(Name="Calendario UPC")
    except Exception:
        pass # Si ya existe en la plantilla global
    try:
        p.Calendar = "Calendario UPC"
        print("Calendario UPC asignado al proyecto.")
    except Exception as e:
        print(f"Nota en calendario: {e}")
        
    # Configuración de horas
    try:
        p.DefaultStartTime = "08:00"
        p.DefaultFinishTime = "19:00"
        p.HoursPerDay = 10
        p.HoursPerWeek = 50
        p.StartWeekOn = 2 # Lunes
    except Exception as e:
        print(f"Nota en opciones de horario: {e}")

    # 3. Dar de alta recursos
    print("Creando recursos...")
    res_map = {}
    for r_name, r_type in RESOURCES:
        try:
            r = p.Resources.Add(r_name)
            res_map[r_name] = r
        except Exception:
            pass

    # 4. Insertar tareas
    print("Insertando y estructurando tareas...")
    tasks_created = []
    
    for level, name, duration, milestone, resources, pred in TASKS_DATA:
        t = p.Tasks.Add(name)
        t.OutlineLevel = level
        
        if milestone:
            t.Duration = 0
            t.Milestone = True
        elif duration:
            try:
                t.Duration = app.DurationValue(duration)
            except Exception:
                # Si es semanas en español o inglés
                if "s" in duration or "w" in duration:
                    val = duration.replace("w", "s")
                    t.Duration = app.DurationValue(val)
                elif "h" in duration:
                    t.Duration = app.DurationValue(duration)
                elif "d" in duration:
                    t.Duration = app.DurationValue(duration)
        
        # Asignar recursos
        if resources:
            for r_sub in resources.split(","):
                r_sub = r_sub.strip()
                try:
                    r_obj = res_map.get(r_sub)
                    if r_obj:
                        t.Assignments.Add(t.ID, r_obj.ID, 1.0)
                    else:
                        t.ResourceNames = resources
                except Exception:
                    t.ResourceNames = resources
        
        tasks_created.append(t)
    
    print(f"Total tareas creadas: {len(tasks_created)}")

    # 5. Vinculación lógica requerida en la guía:
    # "Vincular Actividades : 1 del Alcance con la actividad 1 de Requisitos de Análisis/Software"
    # Actividad 1 de Alcance es 'Alcance del proyecto'
    # Actividad 1 de Requisitos es 'Análisis de necesidades'
    print("Aplicando vinculaciones solicitadas en la guía...")
    
    # Mapear nombres a tareas para vinculación exacta
    task_by_name = {t.Name: t for t in p.Tasks}
    
    # 1. Hito de inicio vinculado a la primera tarea
    if "Inicio del Proyecto" in task_by_name and "Alcance del proyecto" in task_by_name:
        task_by_name["Alcance del proyecto"].Predecessors = str(task_by_name["Inicio del Proyecto"].ID)
        
    # Secuencia dentro de Alcance:
    alcance_seq = [
        "Alcance del proyecto",
        "Patrocinio seguro de proyectos",
        "Recursos preliminares",
        "Recursos básicos seguros",
        "Alcance completo"
    ]
    for i in range(len(alcance_seq) - 1):
        t1 = task_by_name.get(alcance_seq[i])
        t2 = task_by_name.get(alcance_seq[i+1])
        if t1 and t2:
            t2.Predecessors = str(t1.ID)

    # Vínculo explícito de la guía: 1 de Alcance con 1 de Requisitos
    t_alcance_1 = task_by_name.get("Alcance del proyecto")
    t_req_1 = task_by_name.get("Análisis de necesidades")
    if t_alcance_1 and t_req_1:
        # La guía pide: "Vincular Actividades : 1 del Alcance con la actividad 1 de Requisitos de Análisis/Software"
        # Además el hito de fin de alcance enlaza con la siguiente fase
        t_req_1.Predecessors = str(t_alcance_1.ID)

    # Secuencia dentro de Requisitos:
    req_seq = [
        "Análisis de necesidades",
        "Borrador de especificaciones preliminares de software",
        "Presupuesto preliminar",
        "Revisión de especificaciones/el presupuesto del software con el equipo",
        "Comentarios sobre las especificaciones del software",
        "Cronograma de entrega",
        "Aprocaciones para proceder (concepto, cronograma, presupuesto)",
        "Asegure los recursos necesarios",
        "Análisis completo"
    ]
    for i in range(len(req_seq) - 1):
        t1 = task_by_name.get(req_seq[i])
        t2 = task_by_name.get(req_seq[i+1])
        if t1 and t2:
            t2.Predecessors = str(t1.ID)

    # Enlazar Análisis Completo con Inicio de Diseño
    t_analisis_fin = task_by_name.get("Análisis completo")
    t_dis_1 = task_by_name.get("Revisión de las especificaciones preliminares del software")
    if t_analisis_fin and t_dis_1:
        t_dis_1.Predecessors = str(t_analisis_fin.ID)

    # Secuencia dentro de Diseño:
    dis_seq = [
        "Revisión de las especificaciones preliminares del software",
        "Especificaciones funcionales.",
        "Prototipo basado en especificaciones funcionales.",
        "Especificaciones funcionales",
        "Incorporar comentarios en las especificaciones funcionales",
        "Obtener aprobación para continuar",
        "Diseño completo"
    ]
    for i in range(len(dis_seq) - 1):
        t1 = task_by_name.get(dis_seq[i])
        t2 = task_by_name.get(dis_seq[i+1])
        if t1 and t2:
            t2.Predecessors = str(t1.ID)

    # Enlazar Diseño Completo con Desarrollo
    t_dis_fin = task_by_name.get("Diseño completo")
    t_dev_1 = task_by_name.get("Revisión de especificaciones funcionales")
    if t_dis_fin and t_dev_1:
        t_dev_1.Predecessors = str(t_dis_fin.ID)

    dev_seq = [
        "Revisión de especificaciones funcionales",
        "Identificar parámetros de diseño modulares/por niveles",
        "Asignar personal de desarrollo",
        "Desarrollar código",
        "Pruebas de desarrollador (depuración primaria)",
        "Desarrollo completo"
    ]
    for i in range(len(dev_seq) - 1):
        t1 = task_by_name.get(dev_seq[i])
        t2 = task_by_name.get(dev_seq[i+1])
        if t1 and t2:
            t2.Predecessors = str(t1.ID)

    # Enlazar Desarrollo Completo con Pruebas
    t_dev_fin = task_by_name.get("Desarrollo completo")
    t_pruebas_1 = task_by_name.get("Desarrollar planes de prueba de unidad usando especificaciones de producto")
    if t_dev_fin and t_pruebas_1:
        t_pruebas_1.Predecessors = str(t_dev_fin.ID)

    # Enlazar Hito final de proyecto con fin
    t_ultimo_hito = task_by_name.get("Plantilla de desarrollo de software completa")
    t_fin_proy = task_by_name.get("Fin del Proyecto")
    if t_ultimo_hito and t_fin_proy:
        t_fin_proy.Predecessors = str(t_ultimo_hito.ID)

    # 6. Guardar línea base
    try:
        app.BaselineSave(True)
        print("Línea base guardada.")
    except Exception as e:
        print(f"Nota en línea base: {e}")

    # 7. Guardar archivo
    if os.path.exists(OUTPUT_PATH):
        try:
            os.remove(OUTPUT_PATH)
        except Exception:
            pass
            
    p.Activate()
    app.FileSaveAs(OUTPUT_PATH)
    print(f"¡PROYECTO GUARDADO CON ÉXITO EN: {OUTPUT_PATH}!")

if __name__ == "__main__":
    build_project()
