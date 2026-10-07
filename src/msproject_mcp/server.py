"""MCP server para Microsoft Project de escritorio (archivos .mpp) vía COM.

Requiere Windows + Microsoft Project instalado. Todas las llamadas COM se
ejecutan en un único hilo dedicado (COM es apartment-threaded).
"""

from __future__ import annotations

import asyncio
import functools
import os
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, date
from typing import Any

import pythoncom
import win32com.client
from mcp.server import MCPServer

mcp = MCPServer(
    name="msproject",
    instructions=(
        "Controla Microsoft Project de escritorio. Primero abre o crea un proyecto "
        "(open_project / create_project); las demás herramientas actúan sobre el "
        "proyecto activo. Los IDs de tarea son los de la columna ID de Project "
        "(cambian si insertas/borras filas). Duraciones como texto: '5d', '3h', '2w'. "
        "Fechas en ISO: 'YYYY-MM-DD' o 'YYYY-MM-DDTHH:MM'. Recuerda save_project."
    ),
)

# ---------------------------------------------------------------- COM thread
_executor = ThreadPoolExecutor(max_workers=1, initializer=pythoncom.CoInitialize)
_app: Any = None
COM_TIMEOUT = float(os.getenv("MSPROJECT_TIMEOUT", "90"))

PJ_DO_NOT_SAVE, PJ_SAVE = 0, 1
RESOURCE_TYPES = {"work": 0, "material": 1, "cost": 2}
RESOURCE_TYPE_NAMES = {v: k for k, v in RESOURCE_TYPES.items()}


def _get_app() -> Any:
    """Devuelve (y si hace falta reconecta) la instancia de MS Project."""
    global _app
    if _app is not None:
        try:
            _ = _app.Name  # ¿sigue viva?
            return _app
        except Exception:
            _app = None
    try:
        _app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        _app = win32com.client.Dispatch("MSProject.Application")
        _app.Visible = os.getenv("MSPROJECT_VISIBLE", "1") != "0"
    _app.DisplayAlerts = False
    return _app


def com(fn):
    """Ejecuta la función en el hilo COM y la expone como async."""

    @functools.wraps(fn)
    async def wrapper(*args, **kwargs):
        loop = asyncio.get_running_loop()
        fut = loop.run_in_executor(_executor, functools.partial(fn, *args, **kwargs))
        try:
            return await asyncio.wait_for(fut, timeout=COM_TIMEOUT)
        except asyncio.TimeoutError:
            raise RuntimeError(
                f"MS Project no respondió en {COM_TIMEOUT:.0f}s. Revisa si tiene un diálogo "
                "abierto (licencia, inicio de sesión, recuperar archivos) y ciérralo."
            ) from None

    return wrapper


# ---------------------------------------------------------------- helpers
def _project() -> Any:
    app = _get_app()
    if app.Projects.Count == 0:
        raise RuntimeError("No hay proyecto abierto. Usa open_project o create_project.")
    return app.ActiveProject


def _fmt_date(v: Any) -> str | None:
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d %H:%M")
    if v in (None, "", "NA"):
        return None
    return str(v)


def _parse_date(s: str) -> datetime:
    return datetime.fromisoformat(s)


def _dur(minutes: Any) -> str | None:
    try:
        return str(_get_app().DurationFormat(minutes))
    except Exception:
        return None if minutes in (None, "") else str(minutes)


def _duration_value(val: str) -> Any:
    """Convierte texto de duración a minutos, tolerando sufijos en inglés o español."""
    app = _get_app()
    try:
        return app.DurationValue(val)
    except Exception:
        s = val.strip()
        s_trans = re.sub(r'(?i)(\d+(?:[.,]\d+)?)\s*w(?:ks?|eeks?)?', r'\1s', s)
        s_trans = re.sub(r'(?i)(\d+(?:[.,]\d+)?)\s*mo(?:nths?|s)?', r'\1m', s_trans)
        try:
            return app.DurationValue(s_trans)
        except Exception:
            pass
        raise


def _set_predecessors(task: Any, pred_str: str) -> None:
    """Asigna predecesoras traduciendo vínculos (FS->FC, SS->CC, SF->CF) si Project está en español."""
    try:
        task.Predecessors = pred_str
    except Exception:
        trans = re.sub(r'(?i)\bFS\b', 'FC', pred_str)
        trans = re.sub(r'(?i)\bSS\b', 'CC', trans)
        trans = re.sub(r'(?i)\bSF\b', 'CF', trans)
        trans = re.sub(r'(?i)(\d+(?:[.,]\d+)?)\s*w', r'\1s', trans)
        task.Predecessors = trans


def _task(tid: int) -> Any:
    t = _project().Tasks.Item(tid)
    if t is None:
        raise ValueError(f"La fila {tid} está vacía.")
    return t


def _task_dict(t: Any, full: bool = False) -> dict:
    d = {
        "id": t.ID,
        "unique_id": t.UniqueID,
        "name": t.Name,
        "outline_level": t.OutlineLevel,
        "summary": bool(t.Summary),
        "milestone": bool(t.Milestone),
        "duration": _dur(t.Duration),
        "start": _fmt_date(t.Start),
        "finish": _fmt_date(t.Finish),
        "percent_complete": t.PercentComplete,
        "predecessors": t.Predecessors or None,
        "resources": t.ResourceNames or None,
        "critical": bool(t.Critical),
    }
    if full:
        d.update(
            wbs=t.WBS,
            outline_number=t.OutlineNumber,
            successors=t.Successors or None,
            actual_start=_fmt_date(t.ActualStart),
            actual_finish=_fmt_date(t.ActualFinish),
            baseline_start=_fmt_date(t.BaselineStart),
            baseline_finish=_fmt_date(t.BaselineFinish),
            total_slack=_dur(t.TotalSlack),
            work=_dur(t.Work),
            cost=t.Cost,
            notes=t.Notes or None,
        )
    return d


def _iter_tasks():
    for t in _project().Tasks:
        if t is not None:
            yield t


def _resource(ref: str | int) -> Any:
    p = _project()
    if isinstance(ref, int) or str(ref).isdigit():
        r = p.Resources.Item(int(ref))
        if r is not None:
            return r
    for r in p.Resources:
        if r is not None and r.Name.lower() == str(ref).lower():
            return r
    raise ValueError(f"Recurso '{ref}' no encontrado.")


def _resource_dict(r: Any) -> dict:
    return {
        "id": r.ID,
        "name": r.Name,
        "type": RESOURCE_TYPE_NAMES.get(r.Type, r.Type),
        "max_units_pct": round(r.MaxUnits * 100, 2) if r.Type == 0 else None,
        "standard_rate": str(r.StandardRate),
        "work": _dur(r.Work),
        "cost": r.Cost,
        "overallocated": bool(r.Overallocated),
    }


def _project_info(p: Any) -> dict:
    return {
        "name": p.Name,
        "path": p.FullName,
        "title": p.Title,
        "start": _fmt_date(p.ProjectStart),
        "finish": _fmt_date(p.ProjectFinish),
        "status_date": _fmt_date(p.StatusDate),
        "tasks": p.Tasks.Count,
        "resources": p.Resources.Count,
        "saved": bool(p.Saved),
    }


# ---------------------------------------------------------------- archivos
@mcp.tool()
@com
def open_project(path: str, read_only: bool = False) -> dict:
    """Abre un archivo .mpp (o .xml/.mpx) y lo deja como proyecto activo."""
    app = _get_app()
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    if not app.FileOpenEx(path, read_only):
        raise RuntimeError(f"Project no pudo abrir {path}")
    return _project_info(app.ActiveProject)


@mcp.tool()
@com
def create_project(start_date: str | None = None, title: str | None = None, path: str | None = None) -> dict:
    """Crea un proyecto nuevo. Si das `path` lo guarda de inmediato como .mpp."""
    app = _get_app()
    p = app.Projects.Add(False)
    if start_date:
        p.ProjectStart = _parse_date(start_date)
    if title:
        p.Title = title
    if path:
        p.Activate()
        app.FileSaveAs(os.path.abspath(path))
    return _project_info(app.ActiveProject)


@mcp.tool()
@com
def list_open_projects() -> list[dict]:
    """Lista los proyectos abiertos en Project."""
    app = _get_app()
    active = app.ActiveProject.FullName if app.Projects.Count else None
    return [{**_project_info(p), "active": p.FullName == active} for p in app.Projects]


@mcp.tool()
@com
def switch_project(name: str) -> dict:
    """Activa otro proyecto abierto (por nombre de archivo o ruta)."""
    for p in _get_app().Projects:
        if name.lower() in (p.Name.lower(), p.FullName.lower()):
            p.Activate()
            return _project_info(p)
    raise ValueError(f"Proyecto '{name}' no está abierto.")


@mcp.tool()
@com
def save_project(path: str | None = None) -> dict:
    """Guarda el proyecto activo. Con `path` hace 'Guardar como' (.mpp o .xml)."""
    app = _get_app()
    _project()
    if path:
        path = os.path.abspath(path)
        if path.lower().endswith(".xml"):
            app.FileSaveAs(path, FormatID="MSProject.XML")
        else:
            app.FileSaveAs(path)
    else:
        app.FileSave()
    return _project_info(app.ActiveProject)


@mcp.tool()
@com
def export_pdf(path: str) -> str:
    """Exporta la vista actual del proyecto activo a PDF."""
    _project()
    path = os.path.abspath(path)
    _get_app().DocumentExport(path, 0)  # 0 = pjPDF
    return path


@mcp.tool()
@com
def close_project(save: bool = False) -> str:
    """Cierra el proyecto activo (opcionalmente guardando)."""
    p = _project()
    name = p.Name
    _get_app().FileCloseEx(PJ_SAVE if save else PJ_DO_NOT_SAVE)
    return f"Cerrado: {name}"


@mcp.tool()
@com
def get_project_info() -> dict:
    """Info general del proyecto activo (fechas, conteos, ruta)."""
    return _project_info(_project())


@mcp.tool()
@com
def update_project(start_date: str | None = None, status_date: str | None = None, title: str | None = None) -> dict:
    """Cambia fecha de inicio, fecha de estado o título del proyecto activo."""
    p = _project()
    if start_date:
        p.ProjectStart = _parse_date(start_date)
    if status_date:
        p.StatusDate = _parse_date(status_date)
    if title:
        p.Title = title
    return _project_info(p)


# ---------------------------------------------------------------- tareas
@mcp.tool()
@com
def list_tasks(search: str | None = None, include_summary: bool = True, only_incomplete: bool = False) -> list[dict]:
    """Lista las tareas del proyecto activo con fechas, avance, predecesoras y recursos."""
    out = []
    for t in _iter_tasks():
        if not include_summary and t.Summary:
            continue
        if only_incomplete and t.PercentComplete >= 100:
            continue
        if search and search.lower() not in (t.Name or "").lower():
            continue
        out.append(_task_dict(t))
    return out


@mcp.tool()
@com
def get_task(task_id: int) -> dict:
    """Detalle completo de una tarea (línea base, holgura, trabajo, costo, notas, asignaciones)."""
    t = _task(task_id)
    d = _task_dict(t, full=True)
    d["assignments"] = [
        {"resource": a.ResourceName, "units_pct": round(a.Units * 100, 2), "work": _dur(a.Work)}
        for a in t.Assignments
    ]
    return d


@mcp.tool()
@com
def add_task(
    name: str,
    duration: str | None = None,
    start: str | None = None,
    predecessors: str | None = None,
    before_id: int | None = None,
    outline_level: int | None = None,
    milestone: bool = False,
    notes: str | None = None,
) -> dict:
    """Agrega una tarea. `before_id` la inserta antes de esa fila (si no, al final).
    `predecessors` ej. '3' o '3FS+2d,5SS'. `outline_level` 1 = nivel superior."""
    p = _project()
    t = p.Tasks.Add(name, before_id) if before_id else p.Tasks.Add(name)
    if outline_level:
        t.OutlineLevel = outline_level
    if milestone:
        t.Duration = 0
        t.Milestone = True
    elif duration:
        t.Duration = _duration_value(duration)
    if start:
        t.Start = _parse_date(start)
    if predecessors:
        _set_predecessors(t, predecessors)
    if notes:
        t.Notes = notes
    return _task_dict(t)


@mcp.tool()
@com
def update_task(
    task_id: int,
    name: str | None = None,
    duration: str | None = None,
    start: str | None = None,
    finish: str | None = None,
    percent_complete: int | None = None,
    actual_start: str | None = None,
    actual_finish: str | None = None,
    predecessors: str | None = None,
    outline_level: int | None = None,
    notes: str | None = None,
) -> dict:
    """Edita una tarea: nombre, duración, fechas, % completado, reales, predecesoras
    (usa '' para quitarlas), nivel de esquema o notas."""
    t = _task(task_id)
    if name is not None:
        t.Name = name
    if duration:
        t.Duration = _duration_value(duration)
    if start:
        t.Start = _parse_date(start)
    if finish:
        t.Finish = _parse_date(finish)
    if actual_start:
        t.ActualStart = _parse_date(actual_start)
    if actual_finish:
        t.ActualFinish = _parse_date(actual_finish)
    if percent_complete is not None:
        t.PercentComplete = percent_complete
    if predecessors is not None:
        _set_predecessors(t, predecessors)
    if outline_level:
        t.OutlineLevel = outline_level
    if notes is not None:
        t.Notes = notes
    return _task_dict(t, full=True)


@mcp.tool()
@com
def indent_task(task_id: int, outdent: bool = False) -> dict:
    """Aplica sangría (convierte en subtarea) o la quita con outdent=True."""
    t = _task(task_id)
    t.OutlineOutdent() if outdent else t.OutlineIndent()
    return _task_dict(t)


@mcp.tool()
@com
def delete_task(task_id: int) -> str:
    """Elimina una tarea (y sus subtareas si es resumen). Los IDs siguientes se recorren."""
    t = _task(task_id)
    name = t.Name
    t.Delete()
    return f"Eliminada tarea {task_id}: {name}"


# ---------------------------------------------------------------- recursos
@mcp.tool()
@com
def list_resources() -> list[dict]:
    """Lista los recursos del proyecto activo."""
    return [_resource_dict(r) for r in _project().Resources if r is not None]


@mcp.tool()
@com
def add_resource(
    name: str,
    type: str = "work",
    standard_rate: str | None = None,
    max_units_pct: float | None = None,
    email: str | None = None,
) -> dict:
    """Crea un recurso. type: work | material | cost. standard_rate ej. '50/h'.
    max_units_pct ej. 100 = tiempo completo."""
    r = _project().Resources.Add(name)
    r.Type = RESOURCE_TYPES[type.lower()]
    if standard_rate and r.Type != 2:
        r.StandardRate = standard_rate
    if max_units_pct is not None and r.Type == 0:
        r.MaxUnits = max_units_pct / 100
    if email:
        r.EMailAddress = email
    return _resource_dict(r)


@mcp.tool()
@com
def update_resource(
    resource: str,
    name: str | None = None,
    standard_rate: str | None = None,
    max_units_pct: float | None = None,
) -> dict:
    """Edita un recurso (por nombre o ID)."""
    r = _resource(resource)
    if name:
        r.Name = name
    if standard_rate:
        r.StandardRate = standard_rate
    if max_units_pct is not None:
        r.MaxUnits = max_units_pct / 100
    return _resource_dict(r)


@mcp.tool()
@com
def delete_resource(resource: str) -> str:
    """Elimina un recurso (por nombre o ID) y sus asignaciones."""
    r = _resource(resource)
    name = r.Name
    r.Delete()
    return f"Eliminado recurso: {name}"


# ---------------------------------------------------------------- asignaciones
@mcp.tool()
@com
def assign_resource(task_id: int, resource: str, units_pct: float = 100) -> dict:
    """Asigna un recurso (nombre o ID) a una tarea. units_pct 100 = tiempo completo."""
    t = _task(task_id)
    r = _resource(resource)
    t.Assignments.Add(t.ID, r.ID, units_pct / 100)
    return _task_dict(t)


@mcp.tool()
@com
def unassign_resource(task_id: int, resource: str) -> dict:
    """Quita la asignación de un recurso a una tarea."""
    t = _task(task_id)
    r = _resource(resource)
    for a in t.Assignments:
        if a.ResourceID == r.ID:
            a.Delete()
            return _task_dict(t)
    raise ValueError(f"'{r.Name}' no está asignado a la tarea {task_id}.")


@mcp.tool()
@com
def list_assignments(resource: str | None = None) -> list[dict]:
    """Lista asignaciones (tarea ↔ recurso). Filtra por recurso si se indica."""
    rid = _resource(resource).ID if resource else None
    out = []
    for t in _iter_tasks():
        for a in t.Assignments:
            if rid is None or a.ResourceID == rid:
                out.append({
                    "task_id": t.ID,
                    "task": t.Name,
                    "resource": a.ResourceName,
                    "units_pct": round(a.Units * 100, 2),
                    "work": _dur(a.Work),
                    "start": _fmt_date(a.Start),
                    "finish": _fmt_date(a.Finish),
                    "percent_work_complete": a.PercentWorkComplete,
                })
    return out


# ---------------------------------------------------------------- seguimiento
@mcp.tool()
@com
def set_baseline() -> str:
    """Guarda la línea base (Baseline) de todo el proyecto activo."""
    p = _project()
    p.Activate()
    _get_app().BaselineSave(True)
    return f"Línea base guardada para {p.Name}"


@mcp.tool()
@com
def critical_path() -> list[dict]:
    """Tareas en la ruta crítica (sin resúmenes)."""
    return [_task_dict(t) for t in _iter_tasks() if t.Critical and not t.Summary]


@mcp.tool()
@com
def status_report(as_of: str | None = None) -> dict:
    """Reporte de estado: avance, tareas atrasadas (fin < fecha y sin terminar),
    tareas en curso, desviaciones contra línea base, recursos sobreasignados."""
    p = _project()
    ref = _parse_date(as_of) if as_of else datetime.now()
    total = done = 0
    late, in_progress, slipped = [], [], []
    for t in _iter_tasks():
        if t.Summary:
            continue
        total += 1
        pct = t.PercentComplete
        fin = t.Finish if isinstance(t.Finish, datetime) else None
        if pct >= 100:
            done += 1
            continue
        if fin and fin.replace(tzinfo=None) < ref:
            late.append({"id": t.ID, "name": t.Name, "finish": _fmt_date(t.Finish), "pct": pct,
                         "resources": t.ResourceNames or None})
        elif pct > 0:
            in_progress.append({"id": t.ID, "name": t.Name, "pct": pct, "finish": _fmt_date(t.Finish)})
        bf = t.BaselineFinish
        if isinstance(bf, datetime) and fin and fin > bf:
            slipped.append({"id": t.ID, "name": t.Name, "baseline_finish": _fmt_date(bf),
                            "finish": _fmt_date(fin), "variance": _dur(t.FinishVariance)})
    return {
        "project": _project_info(p),
        "as_of": ref.strftime("%Y-%m-%d"),
        "project_percent_complete": p.ProjectSummaryTask.PercentComplete,
        "tasks_total": total,
        "tasks_done": done,
        "late_tasks": late,
        "in_progress": in_progress,
        "slipped_vs_baseline": slipped,
        "overallocated_resources": [r.Name for r in p.Resources if r is not None and r.Overallocated],
    }


def main() -> None:
    mcp.run("stdio")


if __name__ == "__main__":
    main()
