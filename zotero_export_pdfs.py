"""
Zotero Collection PDF Exporter
===============================

把 Zotero 某个分类(collection)下的所有 PDF 附件复制到一个文件夹里。

用法(在你自己的电脑上运行,需要先关闭 Zotero 以避免数据库被锁):

    python zotero_export_pdfs.py "体育赛事文库"

    # 指定输出文件夹
    python zotero_export_pdfs.py "体育赛事文库" -o "D:\\体育赛事PDF"

    # 同时包含子分类里的 PDF
    python zotero_export_pdfs.py "体育赛事文库" --recursive

    # 手动指定 Zotero 数据目录(默认会自动查找)
    python zotero_export_pdfs.py "体育赛事文库" --data-dir "D:\\Zotero"

说明:
- 只用 Python 标准库,无需安装第三方包。
- 读取时会先把 zotero.sqlite 复制到临时文件,不会改动你的原始数据库。
- 文件名重复时会自动加序号,不会互相覆盖。
"""

import argparse
import os
import re
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path


# ---------------------------------------------------------------------------
# 定位 Zotero 数据目录
# ---------------------------------------------------------------------------
def find_data_dir(explicit=None):
    """返回包含 zotero.sqlite 的目录。"""
    if explicit:
        p = Path(explicit)
        if (p / "zotero.sqlite").exists():
            return p
        raise SystemExit(f"在 {p} 下找不到 zotero.sqlite")

    candidates = [
        Path.home() / "Zotero",                       # Windows / Linux 默认
        Path.home() / "Library" / "Application Support" / "Zotero",  # 旧 macOS
        Path(os.environ.get("USERPROFILE", "")) / "Zotero",
    ]
    for c in candidates:
        if c and (c / "zotero.sqlite").exists():
            return c

    raise SystemExit(
        "未能自动找到 Zotero 数据目录。\n"
        "请用 --data-dir 指定(在 Zotero 的 设置 → 高级 → 文件和文件夹 → "
        "数据目录位置 里可以看到)。"
    )


# ---------------------------------------------------------------------------
# 数据库查询
# ---------------------------------------------------------------------------
def open_db_readonly(sqlite_path):
    """复制一份数据库再打开,避免锁和误改。"""
    tmp = tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False)
    tmp.close()
    shutil.copy2(sqlite_path, tmp.name)
    conn = sqlite3.connect(tmp.name)
    return conn, tmp.name


def find_collection_ids(conn, name, recursive):
    """根据名字找到 collectionID,可选地包含所有子分类。"""
    cur = conn.cursor()
    rows = cur.execute(
        "SELECT collectionID FROM collections WHERE collectionName = ?", (name,)
    ).fetchall()
    if not rows:
        # 列出所有分类名帮助排查
        all_names = [
            r[0] for r in cur.execute(
                "SELECT collectionName FROM collections ORDER BY collectionName"
            ).fetchall()
        ]
        raise SystemExit(
            f'找不到名为 "{name}" 的分类。\n现有分类有:\n  '
            + "\n  ".join(all_names)
        )

    ids = {r[0] for r in rows}
    if recursive:
        # 反复展开子分类
        changed = True
        while changed:
            changed = False
            placeholders = ",".join("?" * len(ids))
            children = cur.execute(
                f"SELECT collectionID FROM collections "
                f"WHERE parentCollectionID IN ({placeholders})",
                tuple(ids),
            ).fetchall()
            for (cid,) in children:
                if cid not in ids:
                    ids.add(cid)
                    changed = True
    return ids


def get_pdf_attachments(conn, collection_ids):
    """
    返回该分类下所有 PDF 附件的列表,每项为 dict:
      {key, path, linkMode, contentType, title}
    包含:直接放进分类的 PDF 附件,以及条目(item)下作为子附件的 PDF。
    """
    cur = conn.cursor()
    placeholders = ",".join("?" * len(collection_ids))
    item_ids = {
        r[0] for r in cur.execute(
            f"SELECT itemID FROM collectionItems "
            f"WHERE collectionID IN ({placeholders})",
            tuple(collection_ids),
        ).fetchall()
    }
    if not item_ids:
        return []

    ip = ",".join("?" * len(item_ids))
    # 直接附件 (itemID 在分类里) 或 子附件 (parentItemID 在分类里)
    rows = cur.execute(
        f"""
        SELECT ia.itemID, i.key, ia.path, ia.linkMode, ia.contentType
        FROM itemAttachments ia
        JOIN items i ON i.itemID = ia.itemID
        WHERE (ia.itemID IN ({ip}) OR ia.parentItemID IN ({ip}))
          AND ia.contentType = 'application/pdf'
        """,
        tuple(item_ids) + tuple(item_ids),
    ).fetchall()

    results = []
    for item_id, key, path, link_mode, content_type in rows:
        results.append(
            {
                "itemID": item_id,
                "key": key,
                "path": path,
                "linkMode": link_mode,
                "contentType": content_type,
            }
        )
    return results


# ---------------------------------------------------------------------------
# 解析附件实际文件路径
# ---------------------------------------------------------------------------
def resolve_file_path(att, data_dir):
    """把一条附件记录解析成磁盘上的真实文件路径(Path)或 None。"""
    path = att["path"] or ""
    link_mode = att["linkMode"]

    # linkMode: 0/1 = 导入(存在 storage/<key>/ 下),2 = 链接到本地文件
    if path.startswith("storage:"):
        filename = path[len("storage:"):]
        return data_dir / "storage" / att["key"] / filename

    if link_mode == 2:  # 链接的本地文件,path 可能是绝对路径
        p = Path(path)
        if p.is_absolute():
            return p
        # 相对路径(相对附件基目录),退而求其次相对数据目录
        return data_dir / path

    # 兜底:storage/<key>/ 里找唯一的 pdf
    folder = data_dir / "storage" / att["key"]
    if folder.is_dir():
        pdfs = list(folder.glob("*.pdf"))
        if len(pdfs) == 1:
            return pdfs[0]
    return None


def safe_filename(name):
    """去掉文件名里 Windows 不允许的字符。"""
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name.strip() or "untitled.pdf"


def unique_path(folder, filename):
    """若目标已存在则加 _1 _2 ... 避免覆盖。"""
    target = folder / filename
    if not target.exists():
        return target
    stem, suffix = target.stem, target.suffix
    n = 1
    while True:
        cand = folder / f"{stem}_{n}{suffix}"
        if not cand.exists():
            return cand
        n += 1


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="把 Zotero 某分类下所有 PDF 复制到一个文件夹"
    )
    parser.add_argument("collection", help='分类名称,例如 "体育赛事文库"')
    parser.add_argument(
        "-o", "--output", default=None,
        help="输出文件夹(默认:当前目录下 <分类名>_PDF)",
    )
    parser.add_argument(
        "--data-dir", default=None,
        help="Zotero 数据目录(含 zotero.sqlite),不填则自动查找",
    )
    parser.add_argument(
        "--recursive", action="store_true",
        help="同时包含子分类里的 PDF",
    )
    args = parser.parse_args()

    data_dir = find_data_dir(args.data_dir)
    print(f"Zotero 数据目录: {data_dir}")

    out_dir = Path(args.output) if args.output else Path.cwd() / f"{args.collection}_PDF"
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"输出文件夹: {out_dir}")

    conn, tmp_db = open_db_readonly(data_dir / "zotero.sqlite")
    try:
        ids = find_collection_ids(conn, args.collection, args.recursive)
        print(f"匹配到 {len(ids)} 个分类(含子分类={args.recursive})")
        attachments = get_pdf_attachments(conn, ids)
    finally:
        conn.close()
        try:
            os.unlink(tmp_db)
        except OSError:
            pass

    print(f"找到 {len(attachments)} 个 PDF 附件,开始复制...\n")

    copied, missing = 0, 0
    for att in attachments:
        src = resolve_file_path(att, data_dir)
        if not src or not src.exists():
            missing += 1
            print(f"  [缺失] key={att['key']} path={att['path']}")
            continue
        dest = unique_path(out_dir, safe_filename(src.name))
        shutil.copy2(src, dest)
        copied += 1
        print(f"  [复制] {dest.name}")

    print(
        f"\n完成:成功复制 {copied} 个,缺失/无法定位 {missing} 个。"
        f"\nPDF 已保存到: {out_dir}"
    )


if __name__ == "__main__":
    main()
