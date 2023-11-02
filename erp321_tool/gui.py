from os import system
from sys import platform
from threading import Thread
from tkinter import DISABLED, NORMAL, Tk, Widget, messagebox, ttk
from traceback import format_exception
from typing import Awaitable, Callable

from erp321_tool.export_from_erp321 import export_from_erp321
from erp321_tool.import_to_erp321 import import_to_erp321
from erp321_tool.utils import syncify


def get_root():
    root = Tk()
    root.title("聚水潭导单回单工具")
    root.eval("tk::PlaceWindow . center")
    return root


def create_async_command_button(
    master: Widget, text: str, async_command: Callable[[], Awaitable[None]]
):
    def command():
        button.config(text=f"{text}中…", state=DISABLED)

        def done_callback(exception):
            button.config(text=text, state=NORMAL)
            if exception:
                messagebox.showinfo(message="".join(format_exception(exception)))

        Thread(target=syncify(async_command, done_callback)).start()

    button = ttk.Button(master, text=text, command=command)
    return button


def close_browser():
    match platform:
        case "win32":
            system("taskkill /im msedge.exe /F")
        case "darwin":
            system("pkill -x 'Microsoft Edge'")


def main():
    root = get_root()

    frame = ttk.Frame(root, padding=20)
    frame.grid(row=0, column=0)

    text = """
• 本工具仅支持使用Edge浏览器（如未安装，请安装后使用）
• 启用导单、回单功能前，必须关闭所有已打开的Edge浏览器
• 推单表格须与本工具保存在同一目录下
    """.strip()

    label = ttk.Label(
        frame,
        text=text,
        foreground="#444",
        background="#e1e1e1",
        padding=10,
    )
    label.grid(row=0, column=0)

    button_frame = ttk.Frame(frame)
    button_frame.grid(row=1, column=0, pady=(20, 0))

    import_button = create_async_command_button(button_frame, "导单", import_to_erp321)
    import_button.grid(row=1, column=0)

    export_button = create_async_command_button(button_frame, "回单", export_from_erp321)
    export_button.grid(row=1, column=1)

    close_browser_button = ttk.Button(
        button_frame, text="关闭Edge", command=close_browser
    )
    close_browser_button.grid(row=2, column=0)

    exit_browser = ttk.Button(button_frame, text="退出", command=root.destroy)
    exit_browser.grid(row=2, column=1)

    root.mainloop()


if __name__ == "__main__":
    main()
