"""Entry point: excel-mcp-2013 {stdio|http}.

IMPORTANTE: logging va a stderr. Con transporte stdio, stdout es el canal
JSON-RPC y cualquier print/log alli corrompe el protocolo.
"""

import logging
import sys

import typer

app = typer.Typer(help="MCP server for Excel automation via COM (target: Excel 2013).")


def _setup_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        stream=sys.stderr,
    )


@app.command()
def stdio(log_level: str = typer.Option("INFO", help="DEBUG, INFO, WARNING, ERROR")):
    """Run server via stdio transport (default para Claude Code, Cursor)."""
    _setup_logging(log_level)
    from .server import mcp

    mcp.run(transport="stdio")


@app.command()
def http(
    port: int = typer.Option(8080, help="Puerto HTTP"),
    host: str = typer.Option("127.0.0.1", help="Host de escucha"),
    log_level: str = typer.Option("INFO", help="DEBUG, INFO, WARNING, ERROR"),
):
    """Run server via HTTP (streamable) transport."""
    _setup_logging(log_level)
    from .server import mcp

    mcp.run(transport="http", host=host, port=port)


def main():
    app()


if __name__ == "__main__":
    main()
