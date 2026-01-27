from qfluentwidgets import (
    QConfig,
    ConfigItem,
    BoolValidator,
    RangeConfigItem,
    RangeValidator,
    OptionsConfigItem,
    OptionsValidator,
    Theme,
    qconfig,
    setThemeColor,
    themeColor
)
from qfluentwidgets.common.config import EnumSerializer
import os
import json
import sys
import winreg


class Config(QConfig):
    themeMode = OptionsConfigItem(
        "Appearance",
        "ThemeMode",
        Theme.LIGHT,
        OptionsValidator(Theme),
        EnumSerializer(Theme),
        restart=False,
    )
    themeId = ConfigItem("Appearance", "ThemeId", "default", restart=False)

    runAtStartup = ConfigItem("General", "RunAtStartup", False, BoolValidator())
    autoShowOverlay = ConfigItem("General", "AutoShowOverlay", True, BoolValidator())

    showClear = ConfigItem("Toolbar", "ShowClear", True, BoolValidator())
    showSpotlight = ConfigItem("Toolbar", "ShowSpotlight", True, BoolValidator())
    showTimer = ConfigItem("Toolbar", "ShowTimer", True, BoolValidator())
    showToolbarText = ConfigItem("Toolbar", "ShowToolbarText", False, BoolValidator())

    showStatusBar = ConfigItem("Overlay", "ShowStatusBar", True, BoolValidator())
    safeArea = RangeConfigItem("Overlay", "SafeArea", 0, RangeValidator(0, 100), restart=False)
    scale = RangeConfigItem("Overlay", "Scale", 1.0, RangeValidator(0.5, 2.0), restart=False)
    popWindowScale = RangeConfigItem("Overlay", "PopWindowScale", 1.0, RangeValidator(0.5, 3.0), restart=False)

    autoHandleInk = ConfigItem("PPT", "AutoHandleInk", True, BoolValidator())

    overlayScreen = OptionsConfigItem(
    "Overlay",
    "OverlayScreen",
    "Auto",
    OptionsValidator(["Auto", "Primary"] + [f"Screen {i}" for i in range(1, 21)]),
    restart=False,
)

    splashMode = OptionsConfigItem(
        "General",
        "SplashMode",
        "Always",
        OptionsValidator(["Always", "Never", "HideOnAutoStart", "TimeRange"]),
        restart=False,
    )
    splashStartTime = ConfigItem("General", "SplashStartTime", "08:00", restart=False)
    splashEndTime = ConfigItem("General", "SplashEndTime", "20:00", restart=False)

    quickLaunchApps = ConfigItem("Toolbar", "QuickLaunchApps", [], restart=False)
    toolbarOrder = ConfigItem("Toolbar", "ToolbarOrder", ["select", "pen", "eraser", "spotlight", "timer", "clear", "apps"], restart=False)


cfg = Config()
_root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 路径配置：开发环境保持原样，打包环境迁移到程序同级
if getattr(sys, "frozen", False):
    # 打包环境：settings.json 和 plugins 文件夹都在 exe 同级
    _base_exe_dir = os.path.dirname(sys.executable)
    SETTINGS_PATH = os.path.join(_base_exe_dir, "settings.json")
    PLUGINS_DIR = os.path.join(_base_exe_dir, "plugins")
else:
    # 开发环境：settings.json 在项目根目录，外置插件目录也在根目录下的 plugins_external (避免与源码 plugins 冲突)
    _root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    SETTINGS_PATH = os.path.join(_root_dir, "settings.json")
    PLUGINS_DIR = os.path.join(_root_dir, "plugins_external")

if not os.path.exists(PLUGINS_DIR):
    try:
        os.makedirs(PLUGINS_DIR)
    except:
        pass

qconfig.load(SETTINGS_PATH, cfg)


def _load_settings_json():
    if not os.path.exists(SETTINGS_PATH):
        return {}
    try:
        with open(SETTINGS_PATH, "rb") as f:
            raw = f.read()
    except Exception:
        return {}
    for enc in ("utf-8", "utf-8-sig", "gbk", "gb18030", "cp1252"):
        try:
            text = raw.decode(enc, errors="ignore")
        except Exception:
            continue
        try:
            data = json.loads(text)
        except Exception:
            continue
        return data if isinstance(data, dict) else {}
    return {}


def _set_run_at_startup(enabled: bool):
    """设置或取消开机自启 (Windows 注册表)"""
    if sys.platform != "win32":
        return
        
    app_name = "Kazuha"
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
        ) as key:
            if enabled:
                if getattr(sys, "frozen", False):
                    app_path = sys.executable
                else:
                    return
                
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{app_path}" --autostart')
            else:
                return
    except Exception as e:
        print(f"Error setting startup: {e}")


def _on_run_at_startup_changed(enabled):
    _set_run_at_startup(enabled)
    _save_cfg()


def _apply_theme_and_color(theme_value):
    if isinstance(theme_value, Theme):
        qconfig.theme = theme_value
    else:
        try:
            qconfig.theme = Theme(theme_value)
        except Exception:
            qconfig.theme = Theme.LIGHT

    if qconfig.theme == Theme.DARK:
        setThemeColor("#E1EBFF")
    else:
        setThemeColor("#3275F5")


_apply_theme_and_color(cfg.themeMode.value)


def _save_cfg():
    old_data = _load_settings_json()

    qconfig.save()

    new_data = _load_settings_json()

    merged = old_data if isinstance(old_data, dict) else {}
    if isinstance(new_data, dict):
        for cat, cat_val in new_data.items():
            if not isinstance(cat_val, dict):
                merged[cat] = cat_val
                continue
            if cat not in merged or not isinstance(merged[cat], dict):
                merged[cat] = {}
            for key, value in cat_val.items():
                merged[cat][key] = value

    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(merged, f, indent=4, ensure_ascii=False)
    except Exception:
        pass


def _on_theme_changed(theme):
    _apply_theme_and_color(theme)
    _save_cfg()


def _bind_auto_save():
    cfg.themeMode.valueChanged.connect(_on_theme_changed)
    cfg.themeId.valueChanged.connect(lambda *_: _save_cfg())
    cfg.runAtStartup.valueChanged.connect(_on_run_at_startup_changed)
    cfg.autoShowOverlay.valueChanged.connect(lambda *_: _save_cfg())
    cfg.showClear.valueChanged.connect(lambda *_: _save_cfg())
    cfg.showSpotlight.valueChanged.connect(lambda *_: _save_cfg())
    cfg.showTimer.valueChanged.connect(lambda *_: _save_cfg())
    cfg.showStatusBar.valueChanged.connect(lambda *_: _save_cfg())
    cfg.safeArea.valueChanged.connect(lambda *_: _save_cfg())
    cfg.scale.valueChanged.connect(lambda *_: _save_cfg())
    cfg.autoHandleInk.valueChanged.connect(lambda *_: _save_cfg())
    cfg.overlayScreen.valueChanged.connect(lambda *_: _save_cfg())
    cfg.splashMode.valueChanged.connect(lambda *_: _save_cfg())
    cfg.splashStartTime.valueChanged.connect(lambda *_: _save_cfg())
    cfg.splashEndTime.valueChanged.connect(lambda *_: _save_cfg())
    # cfg.toolbarLayout.valueChanged.connect(lambda *_: _save_cfg())


_bind_auto_save()
try:
    if cfg.runAtStartup.value:
        _set_run_at_startup(True)
except Exception:
    pass
_save_cfg()


def reload_cfg():
    qconfig.load(SETTINGS_PATH, cfg)
    _apply_theme_and_color(cfg.themeMode.value)
