"""Probe the Access Explorer backend over MCP using only a disposable copy."""
import argparse
import asyncio
import hashlib
import json
from pathlib import Path
import shutil
import tomllib
import uuid

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def digest(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest().upper()


async def probe(args):
    root = Path(__file__).resolve().parent.parent
    source = Path(args.template).resolve(strict=True)
    source_hash = digest(source)
    folder = root / "build" / "tests" / ("access-mcp-" + uuid.uuid4().hex)
    folder.mkdir(parents=True)
    target = folder / "scratch.accdb"
    shutil.copy2(source, target)
    config = tomllib.loads((Path.home() / ".codex" / "config.toml").read_text(encoding="utf-8"))
    registration = config["mcp_servers"]["access_explorer"]
    assert Path(registration["command"]).resolve() == Path(args.python).resolve()
    assert Path(registration["args"][0]).resolve() == Path(args.server).resolve()
    params = StdioServerParameters(command=args.python, args=[args.server])
    results = {}
    async with stdio_client(params) as (reader, writer):
        async with ClientSession(reader, writer) as session:
            await session.initialize()
            tools = (await session.list_tools()).tools
            results["ToolCount"] = len(tools)
            results["CodeExecutionEnabled"] = any(t.name == "access_run_vba" for t in tools)

            async def call(name, arguments):
                result = await session.call_tool(name, arguments)
                text = "\n".join(c.text for c in result.content if c.type == "text")
                if result.isError:
                    raise RuntimeError(text)
                try:
                    value = json.loads(text)
                except ValueError:
                    value = text
                if isinstance(value, dict) and "error" in value:
                    raise RuntimeError(str(value))
                return value

            try:
                modules = await call("access_list_objects", {"db_path": str(target), "object_type": "module"})
                assert "modSwift" in json.dumps(modules)
                results["ModuleDiscovery"] = True
                metadata = await call("access_execute_sql", {"db_path": str(target), "sql": "SELECT AppVersion, SchemaVersion FROM AppMetadata"})
                results["Metadata"] = metadata
                reports = await call("access_execute_sql", {"db_path": str(target), "sql": "SELECT Count(*) AS ReportCount FROM Report"})
                results["ReportInventory"] = reports
                code = await call("access_get_code", {"db_path": str(target), "object_type": "module", "object_name": "modSwift"})
                assert "APP_VERSION" in str(code)
                results["ModuleRead"] = True
            finally:
                results["Close"] = await call("access_close", {})
    assert digest(source) == source_hash, "Original template changed"
    results.update(Status="PASS", Source=str(source), SourceSHA256=source_hash, SourceUnchanged=True, Scratch=str(target), Server=str(Path(args.server).resolve()))
    evidence = folder / "verification.json"
    evidence.write_text(json.dumps(results, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(results, indent=2))
    print("Evidence: " + str(evidence))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--python", required=True)
    parser.add_argument("--server", required=True)
    parser.add_argument("--template", required=True)
    asyncio.run(probe(parser.parse_args()))
