"""doctools.lib.styles — Styles profile loader · 项目无关样式族 SSOT.

distilled from qual-supply/scripts/* (2026-05-25 W2)

用法:
    from doctools.lib.styles import load_profile
    profile = load_profile('zdwp')            # 显式 profile name
    profile = load_profile()                  # 走 default
    profile = load_profile(registry_path=Path(...))  # 自定义 yaml

profile 暴露字段(常用):
    .H1_STYLES / .H2_STYLES / .H3_STYLES / .H4_STYLES / .TITLE_STYLES
    .BODY_STYLES / .TABLE_CAPTION_STYLES / .FIG_CAPTION_STYLES / .TABLE_CELL_STYLES
    .TABLE_STYLE_PRIORITY / .FIGURE_STYLE_PRIORITY
    .CAPTION_NUMBER_FORMAT / .HEADING_NUMBER_FORMAT
    .ZDWP_NEXT_FIELD
    .TABLE_CELL_STYLE_ID / .H1_TARGET_STYLE_ID / ... (各 styleId 目标)

predicate:
    profile.is_h1(name) / is_h2 / is_h3 / is_h4 / is_title /
    profile.is_body(name) / is_table_caption / is_fig_caption / is_table_cell

formatter:
    profile.format_caption('table', h1=2, n=3)   -> "表 2-3　"
    profile.format_caption('figure_h2', h1=2, h2=1, n=4) -> "图2.1-4  "
    profile.format_heading(1, 3)                  -> "3 "
    profile.format_heading(3, 1, 2, 5)            -> "1.2.5 "
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Optional

try:
    import yaml
except ImportError as e:
    raise ImportError(
        "doctools.lib.styles requires PyYAML. Install with: pip install pyyaml"
    ) from e


DEFAULT_REGISTRY_PATH = Path(
    "~/Dev/tools/doctools/config/styles_registry.yaml"
).expanduser()


class StylesProfile:
    """Single profile from styles_registry.yaml.

    所有 fields 透传为属性(dict-like)。集合字段保留为 list (顺序敏感)。
    is_* predicate 用 set 加速 membership 测试。
    """

    def __init__(self, data: dict):
        # 透传所有 key 为属性
        for k, v in data.items():
            setattr(self, k, v)
        # 预计算 set (membership 测试快)
        self._h1_set = set(getattr(self, "H1_STYLES", []) or [])
        self._h2_set = set(getattr(self, "H2_STYLES", []) or [])
        self._h3_set = set(getattr(self, "H3_STYLES", []) or [])
        self._h4_set = set(getattr(self, "H4_STYLES", []) or [])
        self._title_set = set(getattr(self, "TITLE_STYLES", []) or [])
        self._body_set = set(getattr(self, "BODY_STYLES", []) or [])
        self._table_caption_set = set(getattr(self, "TABLE_CAPTION_STYLES", []) or [])
        self._fig_caption_set = set(getattr(self, "FIG_CAPTION_STYLES", []) or [])
        self._table_cell_set = set(getattr(self, "TABLE_CELL_STYLES", []) or [])

    # ---------- predicates ----------
    def is_h1(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._h1_set

    def is_h2(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._h2_set

    def is_h3(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._h3_set

    def is_h4(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._h4_set

    def is_title(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._title_set

    def is_body(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._body_set

    def is_table_caption(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._table_caption_set

    def is_fig_caption(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._fig_caption_set

    def is_table_cell(self, style_name: Optional[str]) -> bool:
        return style_name is not None and style_name in self._table_cell_set

    def heading_level(self, style_name: Optional[str]) -> Optional[int]:
        """Return 1/2/3/4 if style is a heading, else None."""
        if self.is_h1(style_name):
            return 1
        if self.is_h2(style_name):
            return 2
        if self.is_h3(style_name):
            return 3
        if self.is_h4(style_name):
            return 4
        return None

    # ---------- formatters ----------
    def format_caption(self, kind: str, **nums: int) -> str:
        """Render caption prefix; kind ∈ {'table','figure','table_h2','figure_h2'}.

        kwargs: H1, H2, N (大小写敏感, 与 YAML 模板里的 {H1} {H2} {N} 占位符匹配)
        """
        fmt = self.CAPTION_NUMBER_FORMAT.get(kind)
        if fmt is None:
            raise KeyError(f"profile has no CAPTION_NUMBER_FORMAT.{kind}")
        return fmt.format(**nums)

    def format_heading(self, level: int, *nums: int) -> str:
        """Render heading prefix; level ∈ {1,2,3,4}, nums 按层级数传 (h1=1个, h4=4个).

        e.g. format_heading(1, 3)         -> "3 "
             format_heading(3, 1, 2, 5)   -> "1.2.5 "
        """
        key = f"h{level}"
        fmt = self.HEADING_NUMBER_FORMAT.get(key)
        if fmt is None:
            raise KeyError(f"profile has no HEADING_NUMBER_FORMAT.{key}")
        kwargs = {f"n{i+1}": nums[i] if i < len(nums) else None for i in range(4)}
        # also support {n} for h1 case (compat)
        if level == 1 and "{n}" in fmt:
            return fmt.format(n=nums[0])
        return fmt.format(**kwargs)

    def pick_style(self, available: set[str], priority_key: str) -> Optional[str]:
        """从 priority list 里返回 docx 实际可用的第一个 style name.

        priority_key ∈ {'TABLE_STYLE_PRIORITY', 'FIGURE_STYLE_PRIORITY'}
        """
        priority = getattr(self, priority_key, None) or []
        for name in priority:
            if name in available:
                return name
        return None


_PROFILE_CACHE: dict[tuple[str, Optional[str]], StylesProfile] = {}


def load_profile(
    name: Optional[str] = None,
    registry_path: Optional[Path] = None,
) -> StylesProfile:
    """Load a profile from styles_registry.yaml.

    Args:
        name: profile key; None → registry['default']
        registry_path: path to yaml; None → DEFAULT_REGISTRY_PATH

    Returns: StylesProfile instance (cached by (path, name))
    """
    path = (registry_path or DEFAULT_REGISTRY_PATH).resolve()
    cache_key = (str(path), name)
    if cache_key in _PROFILE_CACHE:
        return _PROFILE_CACHE[cache_key]

    if not path.is_file():
        raise FileNotFoundError(f"styles_registry.yaml not found: {path}")

    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    profile_name = name or data.get("default")
    if not profile_name:
        raise ValueError(f"no profile name given and no 'default' in {path}")
    profiles = data.get("profiles") or {}
    if profile_name not in profiles:
        available = sorted(profiles.keys())
        raise KeyError(
            f"profile {profile_name!r} not in {path};"
            f" available: {available}"
        )
    prof = StylesProfile(profiles[profile_name])
    # 保留 profile 名作为元数据
    prof._name = profile_name
    prof._registry_path = str(path)
    _PROFILE_CACHE[cache_key] = prof
    return prof


def list_profiles(registry_path: Optional[Path] = None) -> list[str]:
    """List all profile names in the registry."""
    path = (registry_path or DEFAULT_REGISTRY_PATH).resolve()
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    return sorted((data.get("profiles") or {}).keys())
