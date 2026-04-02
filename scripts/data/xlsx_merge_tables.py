#!/usr/bin/env python3
"""
Excel 多表合并工具（AI智能匹配）

功能：
1. 以主表为基准，从多个辅表补充字段
2. 支持AI智能名称匹配（处理同音字、简称等差异）
3. 支持配置驱动的字段映射

使用示例：
    # 使用配置文件
    python merge_tables.py --config merge_config.json

    # 命令行指定
    python merge_tables.py \\
        --master 主表.xlsx --master-key "名称" \\
        --aux 辅表.xlsx --aux-key "工程名称" \\
        --map "目标列=源列" --map "列B=列A" \\
        -o 输出.xlsx
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import pandas as pd

    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("❌ 此工具需要安装pandas: pip install pandas openpyxl")

# 尝试导入 AI 匹配工具
SCRIPTS_LIB = Path(__file__).parent.parent.parent / "lib" / "tools"
sys.path.insert(0, str(SCRIPTS_LIB))

try:
    from ai_name_matcher import ai_batch_match, clean_name, fuzzy_match, get_api_client, normalize_name  # noqa: F401

    HAS_AI_MATCHER = True
except ImportError:
    HAS_AI_MATCHER = False


# ==================== 名称清理（备用，当 ai_name_matcher 不可用时）====================


def _clean_name_local(name):
    """清理名称用于匹配"""
    if pd.isna(name):
        return ""
    name = str(name).strip()
    for suffix in ["水库", "電站", "电站", "工程", "水电站"]:
        if name.endswith(suffix):
            name = name[: -len(suffix)]
    return name


def _normalize_name_local(name):
    """标准化名称（统一同音字）"""
    name = name.replace("阳", "洋").replace("漈", "际").replace("渔", "鱼")
    name = name.replace("藏", "仓").replace("坎", "坑").replace("垅", "垄")
    name = name.replace("砻", "垄").replace("隆", "垄")
    return name


def _fuzzy_match_local(name1, name2):
    """简单模糊匹配"""
    if name1 == name2:
        return True
    norm1 = _normalize_name_local(_clean_name_local(name1))
    norm2 = _normalize_name_local(_clean_name_local(name2))
    if norm1 == norm2:
        return True
    return bool(len(norm1) >= 2 and len(norm2) >= 2 and (norm1 in norm2 or norm2 in norm1))


# ==================== 匹配函数 ====================


def find_match(target_name, source_df, source_key_col, use_ai=True, ai_matches=None):
    """
    在源表中查找匹配的行

    Args:
        target_name: 要查找的名称
        source_df: 源数据DataFrame
        source_key_col: 源表的匹配键列名
        use_ai: 是否使用AI匹配结果
        ai_matches: AI预匹配结果字典 {target: source}

    Returns:
        匹配行的Series，或None
    """
    if pd.isna(target_name):
        return None

    target_name = str(target_name).strip()

    # 1. 精确匹配
    for _, row in source_df.iterrows():
        source_name = str(row.get(source_key_col, "")).strip()
        if source_name == target_name:
            return row

    # 2. 模糊匹配（清理后）
    clean_func = clean_name if HAS_AI_MATCHER else _clean_name_local
    target_clean = clean_func(target_name)

    for _, row in source_df.iterrows():
        source_name = str(row.get(source_key_col, "")).strip()
        source_clean = clean_func(source_name)
        if source_clean and target_clean == source_clean:
            return row

    # 3. 规范化匹配
    fuzzy_func = fuzzy_match if HAS_AI_MATCHER else _fuzzy_match_local

    for _, row in source_df.iterrows():
        source_name = str(row.get(source_key_col, "")).strip()
        if fuzzy_func(target_name, source_name):
            return row

    # 4. AI匹配结果
    if use_ai and ai_matches and target_name in ai_matches:
        matched_name = ai_matches[target_name]
        if matched_name and matched_name != "无匹配":
            for _, row in source_df.iterrows():
                source_name = str(row.get(source_key_col, "")).strip()
                if source_name == matched_name or clean_func(source_name) == clean_func(matched_name):
                    return row

    return None


def ai_prematch(master_names, source_names, source_label="辅表"):
    """
    使用AI预匹配名称

    Args:
        master_names: 主表名称列表
        source_names: 源表名称列表
        source_label: 源表标签（用于打印）

    Returns:
        匹配字典 {主表名称: 源表名称}
    """
    if not HAS_AI_MATCHER:
        print(f"⚠️  AI匹配工具不可用，跳过{source_label}的AI预匹配")
        return {}

    client, config = get_api_client()
    if not client:
        print(f"⚠️  未配置API，跳过{source_label}的AI预匹配")
        return {}

    # 找出模糊匹配失败的名称
    unmatched = []
    clean_func = clean_name if HAS_AI_MATCHER else _clean_name_local
    fuzzy_func = fuzzy_match if HAS_AI_MATCHER else _fuzzy_match_local

    for name in master_names:
        found = False
        name_clean = clean_func(name)
        for src_name in source_names:
            src_clean = clean_func(src_name)
            if name_clean == src_clean or fuzzy_func(name, src_name):
                found = True
                break
        if not found:
            unmatched.append(name)

    if not unmatched:
        print(f"✅ {source_label}: 全部名称已通过规则匹配")
        return {}

    print(f"🤖 {source_label}: {len(unmatched)}个名称需要AI匹配")

    # 调用AI批量匹配
    matches = ai_batch_match(unmatched, list(source_names), client, config)
    return matches


# ==================== 单位转换 ====================

CONVERTERS = {
    "wan_to_yi": lambda x: round(float(x) / 10000, 4) if pd.notna(x) and x != "" else "",
    "yi_to_wan": lambda x: round(float(x) * 10000, 2) if pd.notna(x) and x != "" else "",
    "km2_to_m2": lambda x: round(float(x) * 1000000, 0) if pd.notna(x) and x != "" else "",
    "m2_to_km2": lambda x: round(float(x) / 1000000, 4) if pd.notna(x) and x != "" else "",
}


def apply_converter(value, converter_name):
    """应用转换器"""
    if converter_name not in CONVERTERS:
        return value
    try:
        return CONVERTERS[converter_name](value)
    except Exception:
        return value


# ==================== 合并函数 ====================


def merge_tables(master_df, master_key, sources, output_columns, use_ai=True):
    """
    合并多个表格

    Args:
        master_df: 主表 DataFrame
        master_key: 主表匹配键列名
        sources: 辅表配置列表，每个元素为 dict:
            {
                "df": DataFrame,
                "key": 匹配键列名,
                "label": 标签（打印用）,
                "mappings": [
                    {"target": 目标列, "source": 源列, "converter": 转换器(可选)}
                ]
            }
        output_columns: 输出列顺序列表
        use_ai: 是否使用AI匹配

    Returns:
        合并后的 DataFrame
    """
    # 初始化结果
    result = pd.DataFrame(index=master_df.index)

    # 复制主表列
    for col in master_df.columns:
        result[col] = master_df[col]

    # 为每个辅表进行AI预匹配
    master_names = master_df[master_key].dropna().unique().tolist()
    ai_matches_all = {}

    for src in sources:
        src_df = src["df"]
        src_key = src["key"]
        src_label = src.get("label", "辅表")

        source_names = src_df[src_key].dropna().unique().tolist()

        if use_ai:
            ai_matches = ai_prematch(master_names, source_names, src_label)
            ai_matches_all[src_label] = ai_matches
        else:
            ai_matches_all[src_label] = {}

    # 遍历主表每行进行匹配
    stats = {m["target"]: 0 for src in sources for m in src.get("mappings", [])}

    for idx, row in master_df.iterrows():
        target_name = row.get(master_key)

        for src in sources:
            src_df = src["df"]
            src_key = src["key"]
            src_label = src.get("label", "辅表")
            mappings = src.get("mappings", [])

            # 查找匹配
            matched_row = find_match(
                target_name, src_df, src_key, use_ai=use_ai, ai_matches=ai_matches_all.get(src_label, {})
            )

            if matched_row is None:
                continue

            # 应用字段映射
            for m in mappings:
                target_col = m["target"]
                source_col = m["source"]
                converter = m.get("converter")

                if source_col not in matched_row.index:
                    continue

                value = matched_row[source_col]

                if pd.isna(value) or str(value).strip() in ["", "/", "-", "nan"]:
                    continue

                # 检查目标列是否已有值
                if target_col in result.columns:
                    existing = result.at[idx, target_col]
                    if pd.notna(existing) and str(existing).strip() not in ["", "/", "-", "nan"]:
                        continue

                # 应用转换器
                if converter:
                    value = apply_converter(value, converter)

                result.at[idx, target_col] = value
                stats[target_col] = stats.get(target_col, 0) + 1

    # 重排列顺序
    final_columns = []
    for col in output_columns:
        if col in result.columns:
            final_columns.append(col)

    # 添加其他列
    for col in result.columns:
        if col not in final_columns:
            final_columns.append(col)

    result = result[final_columns]

    # 打印统计
    print("\n📊 字段匹配统计:")
    total_rows = len(master_df)
    for field, count in stats.items():
        pct = (count / total_rows) * 100 if total_rows > 0 else 0
        status = "✅" if pct > 80 else "⚠️" if pct > 50 else "❌"
        print(f"   {status} {field}: {count}/{total_rows} ({pct:.1f}%)")

    return result


# ==================== 配置解析 ====================


def load_config(config_path):
    """加载配置文件"""
    with open(config_path, encoding="utf-8") as f:
        return json.load(f)


def run_from_config(config):
    """
    从配置运行合并

    配置格式:
    {
        "master": {
            "file": "主表.xlsx",
            "sheet": "Sheet1",
            "key": "名称"
        },
        "sources": [
            {
                "file": "辅表.xlsx",
                "sheet": "Sheet1",
                "key": "工程名称",
                "label": "辅表1",
                "mappings": [
                    {"target": "目标列", "source": "源列"},
                    {"target": "库容", "source": "总库容", "converter": "yi_to_wan"}
                ]
            }
        ],
        "output": {
            "file": "输出.xlsx",
            "columns": ["序号", "名称", "库容", ...]
        },
        "use_ai": true
    }
    """
    # 加载主表
    master_file = config["master"]["file"]
    master_sheet = config["master"].get("sheet", 0)
    master_key = config["master"]["key"]

    print(f"📖 读取主表: {master_file}")
    master_df = pd.read_excel(master_file, sheet_name=master_sheet)
    print(f"   {len(master_df)} 行")

    # 加载辅表
    sources = []
    for src_cfg in config.get("sources", []):
        src_file = src_cfg["file"]
        src_sheet = src_cfg.get("sheet", 0)
        src_key = src_cfg["key"]
        src_label = src_cfg.get("label", Path(src_file).stem)
        mappings = src_cfg.get("mappings", [])

        print(f"📖 读取辅表: {src_file}")
        src_df = pd.read_excel(src_file, sheet_name=src_sheet)
        print(f"   {len(src_df)} 行")

        sources.append({"df": src_df, "key": src_key, "label": src_label, "mappings": mappings})

    # 合并
    output_columns = config.get("output", {}).get("columns", [])
    use_ai = config.get("use_ai", True)

    result = merge_tables(master_df, master_key, sources, output_columns, use_ai)

    # 保存
    output_file = config.get("output", {}).get("file", "merged_output.xlsx")
    result.to_excel(output_file, index=False)
    print(f"\n💾 已保存: {output_file}")

    return result


# ==================== 主函数 ====================


def main():
    parser = argparse.ArgumentParser(
        description="Excel多表合并工具（AI智能匹配）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例：

1. 使用配置文件:
   python merge_tables.py --config merge_config.json

2. 命令行指定:
   python merge_tables.py \\
       --master 主表.xlsx --master-key "名称" \\
       --aux 辅表.xlsx --aux-key "工程名称" \\
       --map "目标列=源列" \\
       -o 输出.xlsx

配置文件格式见源码文档。
""",
    )

    parser.add_argument("--config", "-c", help="配置文件路径")
    parser.add_argument("--master", help="主表文件")
    parser.add_argument("--master-key", help="主表匹配键列名")
    parser.add_argument("--aux", action="append", help="辅表文件（可多次指定）")
    parser.add_argument("--aux-key", action="append", help="辅表匹配键列名（与--aux对应）")
    parser.add_argument("--map", action="append", help='字段映射 "目标列=源列[:转换器]"')
    parser.add_argument("-o", "--output", help="输出文件")
    parser.add_argument("--no-ai", action="store_true", help="禁用AI匹配")

    args = parser.parse_args()

    if not HAS_PANDAS:
        sys.exit(1)

    # 配置文件模式
    if args.config:
        config = load_config(args.config)
        run_from_config(config)
        return

    # 命令行模式
    if not args.master or not args.master_key:
        parser.error("需要指定 --master 和 --master-key")

    if not args.aux or not args.aux_key:
        parser.error("需要指定至少一个 --aux 和对应的 --aux-key")

    if len(args.aux) != len(args.aux_key):
        parser.error("--aux 和 --aux-key 数量不匹配")

    # 构造配置
    config = {
        "master": {"file": args.master, "key": args.master_key},
        "sources": [],
        "output": {"file": args.output or "merged_output.xlsx", "columns": []},
        "use_ai": not args.no_ai,
    }

    # 解析映射
    mappings = []
    if args.map:
        for m in args.map:
            if "=" not in m:
                continue
            parts = m.split("=", 1)
            target = parts[0].strip()
            source_part = parts[1].strip()

            if ":" in source_part:
                source, converter = source_part.rsplit(":", 1)
            else:
                source = source_part
                converter = None

            mapping = {"target": target, "source": source}
            if converter:
                mapping["converter"] = converter
            mappings.append(mapping)

    # 添加辅表
    for i, aux_file in enumerate(args.aux):
        config["sources"].append(
            {
                "file": aux_file,
                "key": args.aux_key[i],
                "label": Path(aux_file).stem,
                "mappings": mappings,  # 所有辅表共用映射（简化版）
            }
        )

    run_from_config(config)


if __name__ == "__main__":
    main()
