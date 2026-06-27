import os
import subprocess
import tkinter as tk
from tkinter import messagebox, scrolledtext

# =========================
# 配置区
# =========================
# 锁定的根目录
BASE_DIR = r"E:\codex_workspace"
# 默认的 Codex 配置文件名
CODEX_PROFILE = "deepseek"

# =========================
# 过滤并获取子项目
# =========================
def get_projects():
    if not os.path.exists(BASE_DIR):
        return []

    # 过滤掉不需要显示的系统/隐藏文件夹
    ignore = {".git", "node_modules", "__pycache__", ".vscode", ".idea"}

    items = []
    try:
        for f in os.listdir(BASE_DIR):
            path = os.path.join(BASE_DIR, f)
            if os.path.isdir(path) and f not in ignore and not f.startswith("."):
                items.append(f)
    except Exception as e:
        print(f"读取目录失败: {e}")

    return sorted(items)

# =========================
# 日志输出
# =========================
def log(msg):
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)
    root.update()

# =========================
# 核心启动逻辑（全在根目录下唤醒）
# =========================
def launch(mode="project"):
    selected = listbox.curselection()

    if mode == "root":
        # 模式1：直接针对整个根目录启动
        target_path = BASE_DIR
        log_target = "ROOT 工作区"
    else:
        # 模式2：针对选中的子项目启动
        if not selected:
            messagebox.showwarning("提示", "请先在左侧选择一个项目！")
            return
        project = listbox.get(selected[0])
        if project == "No Projects Found":
            return
        target_path = os.path.join(BASE_DIR, project)
        log_target = f"子项目 [{project}]"

    log(f"📁 目标路径: {target_path}")

    if not os.path.exists(target_path):
        messagebox.showerror("错误", "该路径不存在！")
        return

    # 1. 启动 VSCode
    try:
        # 使用 shell=True 兼容 Windows 环境变量
        subprocess.Popen(["code", target_path], shell=True)
        log(f"🚀 VSCode 针对 {log_target} 启动指令已发送")
    except Exception as e:
        log(f"⚠ VSCode 启动失败: {e}")

    # 2. 在根目录下启动 Codex CLI，并将目标路径作为参数传给它
    try:
        # 【关键改动】：
        # cwd 永远锁定在 BASE_DIR (根目录)
        # 在 codex 命令最后，追加传入 target_path 参数
        subprocess.Popen(
            ["cmd", "/k", "codex", "--profile", CODEX_PROFILE, target_path],
            cwd=BASE_DIR, 
            creationflags=subprocess.CREATE_NEW_CONSOLE  # 强制弹出新的 CMD 交互窗口
        )
        log(f"🤖 Codex ({CODEX_PROFILE}) 已在根目录下针对 {log_target} 成功唤醒")
    except Exception as e:
        log(f"❌ Codex 启动失败: {e}")

# =========================
# GUI 界面构建
# =========================
root = tk.Tk()
root.title("AI IDE Launcher v3.2 (根目录增强版)")
root.geometry("700x520")
root.config(padx=10, pady=10)

# 左侧项目列表区域
frame_left = tk.Frame(root)
frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

tk.Label(frame_left, text="Workspace Projects", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))

listbox = tk.Listbox(frame_left, font=("Consolas", 10), selectbackground="#1f77b4")
listbox.pack(fill=tk.BOTH, expand=True)

# 绑定双击快捷键：双击列表项直接启动子项目
listbox.bind("<Double-Button-1>", lambda event: launch("project"))

# 右侧日志控制台区域
frame_right = tk.Frame(root)
frame_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

tk.Label(frame_right, text="Console Log", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))

log_box = scrolledtext.ScrolledText(frame_right, font=("Consolas", 9), bg="#f5f5f5")
log_box.pack(fill=tk.BOTH, expand=True)

# 载入工作区项目
projects = get_projects()
if not projects:
    listbox.insert(tk.END, "No Projects Found")
    listbox.config(state=tk.DISABLED)
else:
    for p in projects:
        listbox.insert(tk.END, p)

# 底部按钮控制区
frame_bottom = tk.Frame(root)
frame_bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

btn_root = tk.Button(frame_bottom, text="🚀 启动 ROOT Workspace", 
                     bg="#e1f5fe", fg="#0288d1", font=("Arial", 10),
                     command=lambda: launch("root"))
btn_root.pack(side=tk.LEFT, padx=5)

btn_project = tk.Button(frame_bottom, text="▶ 启动选中项目 (或双击项目)", 
                        bg="#e8f5e9", fg="#388e3c", font=("Arial", 10, "bold"),
                        command=lambda: launch("project"))
btn_project.pack(side=tk.LEFT, padx=5)

btn_exit = tk.Button(frame_bottom, text="❌ 退出", font=("Arial", 10),
                     command=root.quit)
btn_exit.pack(side=tk.RIGHT, padx=5)

# 初始化日志提示
log("✅ AI IDE Launcher v3.2 初始化就绪")
log(f"📁 锁定运行根目录: {BASE_DIR}")
if not projects:
    log("⚠ 未在工作区检测到任何有效子项目文件夹，请检查路径。")

root.mainloop()