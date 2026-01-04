#convert_vbs_to_python.py

import re

def parse_sap_vbs_line(line: str) -> str | None:
    s = line.strip()

    if not s:
        return None

    # ---------- 忽略控制结构 ----------
    if s.startswith(("If ", "End If", "Else")):
        return None

    # ---------- Boolean ----------
    s = s.replace("true", "True").replace("false", "False")


    # ---------- 属性赋值 ----------
    # .text = "ABC"
    if re.search(r'\.\w+\s*=\s*.+', s):
        return s

    # ---------- 方法 + 参数 ----------
    # .sendVKey 0
    m = re.match(r'(.+)\.(\w+)\s+(.+)', s)
    if m:
        obj, method, args = m.groups()
        return f"{obj}.{method}({args})"

    # ---------- 无参方法 ----------
    # .press
    if re.match(r'.+\.\w+$', s):
        return s + "()"

    # ---------- 默认原样 ----------
    return s

def convert_vbs_to_python(vbs_path: str, py_path: str, encoding="utf-16"):
    py_lines = [
        "import win32com.client",
        "import pythoncom",
        "",
        "SapGui = win32com.client.GetObject('SAPGUI').GetScriptingEngine.Children(0)",
        "session = SapGui.FindById('ses[0]')",
        ""
    ]

    skip_block = False  # ← 是否正在跳过 If ... End If 块

    with open(vbs_path, "r", encoding=encoding, errors="ignore") as f:
        for raw_line in f:
            line = raw_line.strip()

            # ---------- 进入需要跳过的 If 块 ----------
            if is_vbs_bootstrap_if(line):
                skip_block = True
                continue

            # ---------- 跳过 End If ----------
            if skip_block:
                if line.startswith("End If"):
                    skip_block = False
                continue

            # ---------- 正常 SAP 行 ----------
            py_line = parse_sap_vbs_line(raw_line)
            if py_line:
                py_lines.append(py_line)

    with open(py_path, "w", encoding="utf-8") as f:
        f.write("\n".join(py_lines))


def is_vbs_bootstrap_if(line: str) -> bool:
    s = line.strip()
    return (
        s.startswith("If Not IsObject(")
        or s.startswith("If IsObject(WScript)")
    )


if __name__ == "__main__":
    convert_vbs_to_python(
        r"C:\Users\admin\AppData\Roaming\SAP\SAP GUI\Scripts\TID.vbs",
        r"C:\Users\admin\AppData\Roaming\SAP\SAP GUI\Scripts\TID2.py",
        encoding="utf-16"   # 如果是 utf-8，直接改这里
    )
    print("转换完成")
