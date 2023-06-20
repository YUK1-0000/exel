from asyncio import AbstractEventLoop, Future, run as arun, \
    get_running_loop, all_tasks, gather

from datetime import datetime
from pathlib import PurePath, Path

import openai

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell

from aiofiles import open as aopen


KEY = "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"
openai.api_key = KEY
started_at = datetime.now()


class Tools:
    def __init__(self, loop: AbstractEventLoop | None = None) -> None:
        self.loop = loop or get_running_loop()

    def to_path(self, path: PurePath | str) -> Path:
        return path if isinstance(path, Path) else Path(path)

    def iterdir(self, path: PurePath | str) -> Future[set[Path]]:
        return self.loop.run_in_executor(None, lambda: {
            child for child in self.to_path(path).iterdir()
            if child.suffix == ".xlsx"
        })

    def _open_sheet_block(self, path: Path) -> Worksheet:
        wb = load_workbook(path)
        sheet = wb["Sheet1"]
        wb.close()
        return sheet

    def open_sheet(self, path: PurePath | str) -> Future[Worksheet]:
        return self.loop.run_in_executor(
            None, self._open_sheet_block,
            self.to_path(path)
        )


async def do_question(cell: Cell) -> str:
    return "\n\n・%s %s\n%s" % (
        cell.coordinate, cell.value,
        (await openai.ChatCompletion.acreate(
            model="gpt-3.5-turbo", 
            messages=(
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": cell.value}
            )
        ))["choices"][0]["message"]["content"]
    )


async def process_sheet(tools: Tools, path: Path) -> None:
    for row in await tools.open_sheet(path):
        async with aopen(f"outputs/{started_at:%Y-%m-%d_%H:%M}.txt", "a") as f:
            await f.write("".join({await task for task in {
                tools.loop.create_task(
                    do_question(cell),
                    name="process_sheet: cell"
                )
                for cell in filter(lambda c: bool(c.value), row)
            }}))


async def main():
    tools = Tools()
    await gather(*(
        process_sheet(tools, child)
        for child in await tools.iterdir("inputs")
    ))
    for task in all_tasks():
        if task.get_name() != "Task-1":
            await task


arun(main())