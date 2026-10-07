import sys, os, time
import win32com.client
import pythoncom

pythoncom.CoInitialize()
print("1. Buscando o abriendo MSProject...", flush=True)
try:
    app = win32com.client.GetActiveObject("MSProject.Application")
    print("Conectado a instancia activa existente", flush=True)
except Exception as e:
    print(f"No hay activa ({e}), haciendo Dispatch...", flush=True)
    app = win32com.client.Dispatch("MSProject.Application")
    app.Visible = True
    print("Dispatch exitoso", flush=True)

app.DisplayAlerts = False
print(f"App: {app.Name}, Version: {app.Version}", flush=True)

print("2. Creando nuevo proyecto (Projects.Add)...", flush=True)
p = app.Projects.Add(False)
print(f"Proyecto creado: {p.Name}", flush=True)

print("3. Cambiando título...", flush=True)
p.Title = "Test Titulo"
print(f"Título asignado: {p.Title}", flush=True)

print("4. Agregando tarea...", flush=True)
t = p.Tasks.Add("Tarea 1")
print(f"Tarea agregada: ID={t.ID}, Name={t.Name}", flush=True)

print("5. Modificando duracion...", flush=True)
t.Duration = app.DurationValue("3d")
print(f"Duración: {t.Duration}", flush=True)

print("6. Agregando recurso...", flush=True)
r = p.Resources.Add("Juan Perez")
print(f"Recurso: {r.Name}", flush=True)

print("7. Asignando recurso...", flush=True)
t.Assignments.Add(t.ID, r.ID, 1.0)
print("Asignación lista", flush=True)

print("8. Guardando como archivo local...", flush=True)
test_path = os.path.abspath("test_step.mpp")
if os.path.exists(test_path):
    os.remove(test_path)
app.FileSaveAs(test_path)
print(f"Guardado en {test_path}!", flush=True)

print("9. Cerrando proyecto...", flush=True)
app.FileCloseEx(0) # pjDoNotSave
print("¡TODO FUNCIONÓ PERFECTO!", flush=True)
