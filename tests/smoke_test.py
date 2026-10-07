"""Prueba rápida: lista las herramientas vía cliente MCP y ejecuta un flujo real en Project."""
import asyncio, json, os, sys
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "smoke_test.mpp")


if os.path.exists(OUT):
    try:
        os.remove(OUT)
    except Exception:
        pass

async def main():
    params = StdioServerParameters(command=sys.executable, args=[os.path.join(HERE, "server.py")],
                                   env={**os.environ})
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        tools = (await s.list_tools()).tools
        print(f"{len(tools)} herramientas:", ", ".join(t.name for t in tools))

        async def call(tool_name, **args):
            res = await s.call_tool(tool_name, args)
            text = res.content[0].text if res.content else ""
            print(f"\n> {tool_name} {'ERROR' if res.is_error else ''}\n{text[:600]}")
            if res.is_error:
                raise SystemExit(1)
            return text

        await call("create_project", start_date="2026-10-12", title="Prueba MCP", path=OUT)
        await call("add_task", name="Fase 1")
        await call("add_task", name="Diseño", duration="3d")
        await call("add_task", name="Desarrollo", duration="1w", predecessors="2")
        await call("indent_task", task_id=2)
        await call("indent_task", task_id=3)
        await call("add_task", name="Entrega", milestone=True, predecessors="3")
        await call("add_resource", name="Ana", standard_rate="50/h")
        await call("assign_resource", task_id=2, resource="Ana")
        await call("set_baseline")
        await call("update_task", task_id=2, percent_complete=50, duration="5d")
        await call("list_tasks")
        await call("critical_path")
        await call("status_report", as_of="2026-10-20")
        await call("save_project")
        await call("close_project")


asyncio.run(main())
