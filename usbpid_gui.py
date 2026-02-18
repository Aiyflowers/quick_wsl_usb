#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
USB IPD GUI — USB Device Manager for WSL
A PyQt5 application that provides a graphical interface for usbipd-win,
allowing users to scan, bind, attach, and detach USB devices for WSL.

Author: aiyflowers
"""

import sys
import subprocess
import re
import webbrowser
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLabel, QStatusBar,
    QMessageBox, QHeaderView, QFrame, QGraphicsDropShadowEffect,
    QAbstractItemView
)
from PyQt5.QtCore import Qt, QTimer, QProcess
from PyQt5.QtGui import QFont, QColor, QIcon, QLinearGradient, QPalette, QBrush


# ──────────────────────────────────────────────────────────
#  Color Palette
# ──────────────────────────────────────────────────────────
COLORS = {
    "bg_dark":       "#0d1117",
    "bg_card":       "#161b22",
    "bg_card_alt":   "#1c2333",
    "border":        "#30363d",
    "accent":        "#58a6ff",
    "accent_hover":  "#79c0ff",
    "danger":        "#f85149",
    "danger_hover":  "#ff7b72",
    "success":       "#3fb950",
    "warning":       "#d29922",
    "text_primary":  "#e6edf3",
    "text_secondary":"#8b949e",
    "brand_pink":    "#e94560",
    "brand_gradient_start": "#6c5ce7",
    "brand_gradient_end":   "#e94560",
}


# ──────────────────────────────────────────────────────────
#  Backend: usbipd command wrapper
# ──────────────────────────────────────────────────────────
class UsbIpdManager:
    """Wraps usbipd CLI commands."""

    @staticmethod
    def _run(args, need_admin=False):
        """Run a usbipd command and return (success, stdout, stderr)."""
        cmd = ["usbipd"] + args
        try:
            if need_admin:
                # Use PowerShell Start-Process with -Verb RunAs for elevation
                ps_cmd = (
                    f'Start-Process -FilePath "usbipd" '
                    f'-ArgumentList "{" ".join(args)}" '
                    f'-Verb RunAs -Wait -WindowStyle Hidden'
                )
                result = subprocess.run(
                    ["powershell", "-Command", ps_cmd],
                    capture_output=True, text=True, timeout=30,
                    encoding="utf-8", errors="replace"
                )
            else:
                result = subprocess.run(
                    cmd, capture_output=True, text=True, timeout=15,
                    encoding="utf-8", errors="replace"
                )
            return (result.returncode == 0, result.stdout, result.stderr)
        except FileNotFoundError:
            return (False, "", "usbipd 未安装或不在 PATH 中。\n请先安装 usbipd-win: https://github.com/dorssel/usbipd-win")
        except subprocess.TimeoutExpired:
            return (False, "", "命令执行超时")
        except Exception as e:
            return (False, "", str(e))

    @classmethod
    def list_devices(cls):
        """
        Parse `usbipd list` output into a list of device dicts.
        Returns: (success, devices_list | error_msg)
        """
        ok, stdout, stderr = cls._run(["list"])
        if not ok:
            return (False, stderr or "无法获取设备列表")

        devices = []
        lines = stdout.strip().splitlines()

        # Find the header line to determine column positions
        header_idx = -1
        for i, line in enumerate(lines):
            if "BUSID" in line.upper():
                header_idx = i
                break

        if header_idx == -1:
            return (True, devices)

        # Parse data lines after header (skip separator lines)
        for line in lines[header_idx + 1:]:
            line_stripped = line.strip()
            if not line_stripped or line_stripped.startswith("-") or line_stripped.startswith("="):
                continue
            # Typical format: BUSID  VID:PID  DEVICE               STATE
            # e.g.: 1-1    046d:c52b  Logitech USB Input Device    Shared
            match = re.match(
                r'^(\d+-\d+(?:\.\d+)*)\s+'   # BUSID like 1-1 or 1-1.2
                r'([0-9a-fA-F]{4}:[0-9a-fA-F]{4})\s+'  # VID:PID
                r'(.+?)\s{2,}'                 # Device name (greedy until 2+ spaces)
                r'(\S+.*)$',                   # State
                line_stripped
            )
            if match:
                devices.append({
                    "busid":  match.group(1),
                    "vidpid": match.group(2),
                    "name":   match.group(3).strip(),
                    "state":  match.group(4).strip(),
                })
            else:
                # Fallback: try splitting by 2+ whitespace
                parts = re.split(r'\s{2,}', line_stripped)
                if len(parts) >= 4:
                    devices.append({
                        "busid":  parts[0],
                        "vidpid": parts[1],
                        "name":   parts[2],
                        "state":  parts[3],
                    })

        return (True, devices)

    @classmethod
    def bind(cls, busid):
        ok, stdout, stderr = cls._run(["bind", "--busid", busid], need_admin=True)
        return (ok, stdout + stderr)

    @classmethod
    def attach(cls, busid):
        ok, stdout, stderr = cls._run(["attach", "--wsl", "--busid", busid])
        return (ok, stdout + stderr)

    @classmethod
    def detach(cls, busid):
        ok, stdout, stderr = cls._run(["detach", "--busid", busid])
        return (ok, stdout + stderr)


# ──────────────────────────────────────────────────────────
#  Main Window
# ──────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("USB IPD GUI — aiyflowers")
        self.setMinimumSize(900, 560)
        self.resize(1000, 620)
        self._build_ui()
        self._apply_styles()
        self.refresh_devices()

    # ── UI Construction ──────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(24, 20, 24, 16)
        root_layout.setSpacing(16)

        # Header ─────────────────────────────────────────
        header = QHBoxLayout()
        header.setSpacing(12)

        title = QLabel("⚡ USB IPD GUI")
        title.setObjectName("titleLabel")
        header.addWidget(title)

        header.addStretch()

        badge = QLabel("  ✦ aiyflowers  ")
        badge.setObjectName("brandBadge")
        header.addWidget(badge)

        root_layout.addLayout(header)

        # Subtitle ───────────────────────────────────────
        subtitle = QLabel("管理 USB 设备与 WSL 的绑定连接")
        subtitle.setObjectName("subtitleLabel")
        root_layout.addWidget(subtitle)

        # Separator ──────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setObjectName("separator")
        root_layout.addWidget(sep)

        # Table ──────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setObjectName("deviceTable")
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["BUS ID", "VID:PID", "设备名称", "状态"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)

        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        h.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        h.setSectionResizeMode(2, QHeaderView.Stretch)
        h.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        h.setHighlightSections(False)

        root_layout.addWidget(self.table, 1)

        # Buttons ────────────────────────────────────────
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(12)

        self.btn_refresh = self._make_button("🔄  刷新设备", "refreshBtn")
        self.btn_install = self._make_button("📦  一键安装环境", "installBtn")
        self.btn_bind    = self._make_button("🔗  绑定并连接 WSL", "bindBtn")
        self.btn_detach  = self._make_button("⛓  解绑设备", "detachBtn")

        btn_bar.addWidget(self.btn_refresh)
        btn_bar.addWidget(self.btn_install)
        btn_bar.addStretch()
        btn_bar.addWidget(self.btn_bind)
        btn_bar.addWidget(self.btn_detach)
        root_layout.addLayout(btn_bar)

        # Status bar ─────────────────────────────────────
        self.status = QStatusBar()
        self.status.setObjectName("statusBar")
        self.setStatusBar(self.status)
        self.status.showMessage("就绪")

        # Connections ────────────────────────────────────
        self.btn_refresh.clicked.connect(self.refresh_devices)
        self.btn_install.clicked.connect(self.on_install_env)
        self.btn_bind.clicked.connect(self.on_bind)
        self.btn_detach.clicked.connect(self.on_detach)
        self.table.selectionModel().selectionChanged.connect(self._update_button_states)
        self._update_button_states()

    def _make_button(self, text, name):
        btn = QPushButton(text)
        btn.setObjectName(name)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(40)
        btn.setMinimumWidth(140)
        return btn

    # ── Styles (QSS) ────────────────────────────────────
    def _apply_styles(self):
        self.setStyleSheet(f"""
            /* ── Global ──────────────────── */
            QMainWindow {{
                background-color: {COLORS['bg_dark']};
            }}
            QWidget {{
                color: {COLORS['text_primary']};
                font-family: "Segoe UI", "Microsoft YaHei UI", sans-serif;
                font-size: 13px;
            }}

            /* ── Header ─────────────────── */
            #titleLabel {{
                font-size: 26px;
                font-weight: 700;
                color: {COLORS['accent']};
                padding: 0;
            }}
            #brandBadge {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {COLORS['brand_gradient_start']},
                    stop:1 {COLORS['brand_gradient_end']});
                color: #ffffff;
                font-size: 13px;
                font-weight: 600;
                padding: 6px 16px;
                border-radius: 14px;
            }}
            #subtitleLabel {{
                color: {COLORS['text_secondary']};
                font-size: 13px;
                padding-left: 2px;
            }}

            /* ── Separator ──────────────── */
            #separator {{
                border: none;
                background-color: {COLORS['border']};
                max-height: 1px;
            }}

            /* ── Table ──────────────────── */
            #deviceTable {{
                background-color: {COLORS['bg_card']};
                alternate-background-color: {COLORS['bg_card_alt']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                gridline-color: transparent;
                selection-background-color: rgba(88, 166, 255, 0.15);
                selection-color: {COLORS['text_primary']};
                padding: 4px;
            }}
            #deviceTable::item {{
                padding: 8px 12px;
                border-bottom: 1px solid {COLORS['border']};
            }}
            #deviceTable::item:selected {{
                background-color: rgba(88, 166, 255, 0.18);
            }}
            QHeaderView::section {{
                background-color: {COLORS['bg_dark']};
                color: {COLORS['text_secondary']};
                font-weight: 600;
                font-size: 12px;
                text-transform: uppercase;
                padding: 10px 12px;
                border: none;
                border-bottom: 2px solid {COLORS['border']};
            }}

            /* ── Buttons ────────────────── */
            QPushButton {{
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: 600;
                font-size: 13px;
                background-color: {COLORS['bg_card']};
                color: {COLORS['text_primary']};
            }}
            QPushButton:hover {{
                border-color: {COLORS['accent']};
                background-color: {COLORS['bg_card_alt']};
            }}
            QPushButton:pressed {{
                background-color: {COLORS['bg_dark']};
            }}
            QPushButton:disabled {{
                color: {COLORS['text_secondary']};
                border-color: {COLORS['bg_card']};
                background-color: {COLORS['bg_dark']};
            }}

            /* Refresh button */
            #refreshBtn {{
                border-color: {COLORS['accent']};
                color: {COLORS['accent']};
            }}
            #refreshBtn:hover {{
                background-color: rgba(88, 166, 255, 0.12);
            }}

            /* Bind button */
            #bindBtn {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {COLORS['brand_gradient_start']},
                    stop:1 {COLORS['accent']});
                color: #ffffff;
                border: none;
            }}
            #bindBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #7c6cf7,
                    stop:1 {COLORS['accent_hover']});
            }}
            #bindBtn:disabled {{
                background: {COLORS['bg_card']};
                color: {COLORS['text_secondary']};
                border: 1px solid {COLORS['border']};
            }}

            /* Detach button */
            #detachBtn {{
                border-color: {COLORS['danger']};
                color: {COLORS['danger']};
            }}
            #detachBtn:hover {{
                background-color: rgba(248, 81, 73, 0.12);
                border-color: {COLORS['danger_hover']};
                color: {COLORS['danger_hover']};
            }}
            #detachBtn:disabled {{
                color: {COLORS['text_secondary']};
                border-color: {COLORS['bg_card']};
                background-color: {COLORS['bg_dark']};
            }}

            /* Install button */
            #installBtn {{
                border-color: {COLORS['success']};
                color: {COLORS['success']};
            }}
            #installBtn:hover {{
                background-color: rgba(63, 185, 80, 0.12);
            }}

            /* ── Status bar ─────────────── */
            QStatusBar {{
                background-color: {COLORS['bg_dark']};
                color: {COLORS['text_secondary']};
                font-size: 12px;
                border-top: 1px solid {COLORS['border']};
                padding: 4px 8px;
            }}
        """)

    # ── Device List ──────────────────────────────────────
    def refresh_devices(self):
        self.status.showMessage("⏳ 正在扫描 USB 设备…")
        QApplication.processEvents()

        ok, result = UsbIpdManager.list_devices()
        self.table.setRowCount(0)

        if not ok:
            self.status.showMessage(f"❌ {result}")
            return

        devices = result
        self.table.setRowCount(len(devices))

        for row, dev in enumerate(devices):
            self.table.setItem(row, 0, QTableWidgetItem(dev["busid"]))
            self.table.setItem(row, 1, QTableWidgetItem(dev["vidpid"]))
            self.table.setItem(row, 2, QTableWidgetItem(dev["name"]))

            state_item = QTableWidgetItem(dev["state"])
            state = dev["state"].lower()
            if "attached" in state:
                state_item.setForeground(QColor(COLORS["success"]))
            elif "shared" in state:
                state_item.setForeground(QColor(COLORS["accent"]))
            elif "not shared" in state or "not bound" in state:
                state_item.setForeground(QColor(COLORS["text_secondary"]))
            else:
                state_item.setForeground(QColor(COLORS["warning"]))
            self.table.setItem(row, 3, state_item)

        count = len(devices)
        self.status.showMessage(f"✅ 扫描完成 — 发现 {count} 个 USB 设备")
        self._update_button_states()

    # ── Actions ──────────────────────────────────────────
    def _get_selected_device(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        row = rows[0].row()
        return {
            "busid":  self.table.item(row, 0).text(),
            "vidpid": self.table.item(row, 1).text(),
            "name":   self.table.item(row, 2).text(),
            "state":  self.table.item(row, 3).text(),
        }

    def _update_button_states(self):
        dev = self._get_selected_device()
        has_sel = dev is not None
        self.btn_bind.setEnabled(has_sel)
        self.btn_detach.setEnabled(has_sel)

    def on_bind(self):
        dev = self._get_selected_device()
        if not dev:
            return

        reply = QMessageBox.question(
            self, "确认绑定",
            f"即将绑定设备 <b>{dev['name']}</b> ({dev['busid']})<br>"
            f"并将其连接到 WSL。<br><br>"
            f"<i>绑定操作需要管理员权限，系统可能弹出 UAC 提示。</i><br><br>"
            f"是否继续？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply != QMessageBox.Yes:
            return

        self.status.showMessage(f"⏳ 正在绑定 {dev['busid']}…（需要管理员权限）")
        QApplication.processEvents()

        # Step 1: Bind
        ok, msg = UsbIpdManager.bind(dev["busid"])
        if not ok and "already bound" not in msg.lower():
            self.status.showMessage(f"❌ 绑定失败: {msg.strip()}")
            QMessageBox.warning(self, "绑定失败", f"绑定设备失败:\n{msg.strip()}")
            return

        # Step 2: Attach to WSL
        self.status.showMessage(f"⏳ 正在连接 {dev['busid']} 到 WSL…")
        QApplication.processEvents()

        ok, msg = UsbIpdManager.attach(dev["busid"])
        if ok:
            self.status.showMessage(f"✅ 设备 {dev['busid']} 已成功绑定并连接到 WSL")
        else:
            self.status.showMessage(f"⚠️ 绑定成功但连接 WSL 失败: {msg.strip()}")
            QMessageBox.warning(self, "连接 WSL 失败",
                f"设备已绑定，但连接到 WSL 失败:\n{msg.strip()}\n\n"
                f"请确保 WSL 正在运行。")

        self.refresh_devices()

    def on_detach(self):
        dev = self._get_selected_device()
        if not dev:
            return

        reply = QMessageBox.question(
            self, "确认解绑",
            f"即将从 WSL 解绑设备 <b>{dev['name']}</b> ({dev['busid']})<br><br>"
            f"是否继续？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply != QMessageBox.Yes:
            return

        self.status.showMessage(f"⏳ 正在解绑 {dev['busid']}…")
        QApplication.processEvents()

        ok, msg = UsbIpdManager.detach(dev["busid"])
        if ok:
            self.status.showMessage(f"✅ 设备 {dev['busid']} 已从 WSL 解绑")
        else:
            self.status.showMessage(f"❌ 解绑失败: {msg.strip()}")
            QMessageBox.warning(self, "解绑失败", f"解绑设备失败:\n{msg.strip()}")

        self.refresh_devices()

    def _verify_usbipd_installed(self):
        """Check if usbipd is usable by running `usbipd list`."""
        try:
            result = subprocess.run(
                ["usbipd", "list"],
                capture_output=True, text=True, timeout=10,
                encoding="utf-8", errors="replace"
            )
            return result.returncode == 0
        except Exception:
            return False

    def _prompt_manual_install(self):
        """Show manual install dialog and open GitHub page on close."""
        QMessageBox.critical(
            self, "安装失败 — 请手动安装",
            "自动安装 usbipd-win 未成功。<br><br>"
            "关闭此弹窗后将自动打开 usbipd-win 的 GitHub 页面，<br>"
            "请按照页面说明手动下载安装。<br><br>"
            "<i>安装完成后请重新打开本程序。</i>"
        )
        webbrowser.open("https://github.com/dorssel/usbipd-win?tab=readme-ov-file")

    def on_install_env(self):
        """Install usbipd-win via winget, verify, and fallback to browser."""
        # Step 0: Check if already installed
        if self._verify_usbipd_installed():
            QMessageBox.information(
                self, "已安装",
                "usbipd-win 已经安装，无需重复安装。"
            )
            self.status.showMessage("✅ usbipd-win 已安装")
            return

        reply = QMessageBox.question(
            self, "安装 usbipd-win",
            "检测到 usbipd-win 尚未安装。<br><br>"
            "即将通过 <b>winget</b> 自动安装，流程如下：<br>"
            "① 执行 <code>winget install usbipd</code><br>"
            "② 验证 <code>usbipd list</code> 是否可用<br>"
            "③ 若失败则引导手动安装<br><br>"
            "是否继续？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply != QMessageBox.Yes:
            return

        # Step 1: Try winget install
        self.status.showMessage("⏳ 正在通过 winget 安装 usbipd-win …")
        QApplication.processEvents()

        winget_ok = False
        winget_output = ""
        try:
            result = subprocess.run(
                ["winget", "install", "usbipd"],
                capture_output=True, text=True, timeout=120,
                encoding="utf-8", errors="replace"
            )
            winget_output = (result.stdout + result.stderr).strip()
            winget_ok = (result.returncode == 0)
        except FileNotFoundError:
            winget_output = "未找到 winget 命令"
        except subprocess.TimeoutExpired:
            winget_output = "安装命令执行超时 (120秒)"
        except Exception as e:
            winget_output = str(e)

        # Step 2: Verify installation regardless of winget exit code
        self.status.showMessage("⏳ 正在验证 usbipd 安装…")
        QApplication.processEvents()

        if self._verify_usbipd_installed():
            self.status.showMessage("✅ usbipd-win 安装并验证成功")
            QMessageBox.information(
                self, "安装成功",
                "usbipd-win 安装完成，已通过 <code>usbipd list</code> 验证。"
            )
            self.refresh_devices()
            return

        # Step 3: Installation failed — prompt manual install and open browser
        self.status.showMessage("❌ 自动安装失败，请手动安装")
        self._prompt_manual_install()


# ──────────────────────────────────────────────────────────
#  Entry Point
# ──────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Force dark palette for Fusion base
    palette = QPalette()
    palette.setColor(QPalette.Window,          QColor(COLORS["bg_dark"]))
    palette.setColor(QPalette.WindowText,      QColor(COLORS["text_primary"]))
    palette.setColor(QPalette.Base,            QColor(COLORS["bg_card"]))
    palette.setColor(QPalette.AlternateBase,   QColor(COLORS["bg_card_alt"]))
    palette.setColor(QPalette.Text,            QColor(COLORS["text_primary"]))
    palette.setColor(QPalette.Button,          QColor(COLORS["bg_card"]))
    palette.setColor(QPalette.ButtonText,      QColor(COLORS["text_primary"]))
    palette.setColor(QPalette.Highlight,       QColor(COLORS["accent"]))
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
