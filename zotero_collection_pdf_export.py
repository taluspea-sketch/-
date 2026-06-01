#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Zotero 文库分类 PDF 整合工具 (Zotero Collection PDF Exporter)

把 Zotero 文库中某个分类(文件夹/Collection)下的所有 PDF 附件
整合复制到一个目标文件夹中。

工作原理:
- Zotero 把 PDF 分散存放在 数据目录/storage/<随机KEY>/*.pdf 下;
- 分类(文件夹)的归属关系记录在 zotero.sqlite 数据库里。
本工具只读 zotero.sqlite(会先拷贝到临时文件,避免 Zotero 正在运行时的锁冲突),
找出指定分类(及其子分类)下所有文献的 PDF 附件,复制到目标文件夹。

默认行为:
- 包含子分类(递归)
- 保留原始文件名(重名时自动追加 _1, _2 ...)
- 复制(不动 Zotero 原文件,安全无损)
- 默认 dry-run(只预览不复制),加 --run 才真正执行

用法示例:
    # 预览将要复制哪些文件
    python zotero_collection_pdf_export.py --collection "体育赛事"

    # 真正执行复制到指定文件夹
    python zotero_collection_pdf_export.py --collection "体育赛事" \
        --output "D:/体育赛事PDF" --run

作者: Claude
"""

import argparse
import os
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path


# PDF 附件的 linkMode 含义(Zotero):
#   0 = imported_file  (导入文件, 存放在 storage/<KEY>/ 下)
#   1 = imported_url   (网页快照等, 也存放在 storage/<KEY>/ 下)
#   2 = linked_file    (链接文件, path 为本机绝对路径)
#   3 = linked_url     (仅链接, 无本地文件)
LINKMODE_IMPORTED = (0, 1)
LINKMODE_LINKED_FILE = 2


def default_zotero_dir():
    """猜测默认的 Zotero 数据目录(Windows / macOS / Linux 均为 ~/Zotero)。"""
    return Path.home() / "Zotero"


def open_db_readonly(sqlite_path):
    """以只读方式打开 zotero.sqlite。

    Zotero 运行时会锁库,这里先把数据库拷贝到临时文件再打开,
    既避免锁冲突,也保证绝不会写到原库。
    """
    sqlite_path = Path(sqlite_path)
    if not sqlite_path.exists():
        raise FileNotFoundError(f"找不到 zotero.sqlite: {sqlite_path}")

    tmp = tempfile.NamedTemporaryFile(prefix="zotero_copy_", suffix=".sqlite",
                                      delete=False)
    tmp.close()
    shutil.copy2(sqlite_path, tmp.name)
    conn = sqlite3.connect(tmp.name)
    conn.row_factory = sqlite3.Row
    return conn, tmp.name


def find_collection_ids(conn, name, include_subcollections=True):
    """根据分类名找到 collectionID(可能有重名),并可递归包含子分类。

    返回 (匹配到的根分类信息列表, 全部要导出的 collectionID 集合)。
    """
    cur = conn.cursor()
    rows = cur.execute(
        "SELECT collectionID, collectionName FROM collections "
        "WHERE collectionName = ?",
        (name,),
    ).fetchall()

    roots = [(r["collectionID"], r["collectionName"]) for r in rows]
    all_ids = set(cid for cid, _ in roots)

    if include_subcollections and all_ids:
        # 逐层向下展开子分类
        frontier = set(all_ids)
        while frontier:
            placeholders = ",".join("?" * len(frontier))
            children = cur.execute(
                f"SELECT collectionID FROM collections "
                f"WHERE parentCollectionID IN ({placeholders})",
                tuple(frontier),
            ).fetchall()
            child_ids = {c["collectionID"] for c in children} - all_ids
            all_ids |= child_ids
            frontier = child_ids

    return roots, all_ids


def collect_pdf_attachments(conn, collection_ids, storage_dir):
    """收集指定分类集合下所有 PDF 附件的源路径。

    返回 (找到的 PDF 源路径列表, 跳过/缺失信息列表)。
    覆盖两种情况:
      1. 文献条目(itemID)下挂着 PDF 附件;
      2. 直接被加进分类的 PDF 附件本身。
    """
    cur = conn.cursor()
    placeholders = ",".join("?" * len(collection_ids))
    item_ids = {
        r["itemID"]
        for r in cur.execute(
            f"SELECT DISTINCT itemID FROM collectionItems "
            f"WHERE collectionID IN ({placeholders})",
            tuple(collection_ids),
        ).fetchall()
    }

    if not item_ids:
        return [], []

    # 找出这些条目下的附件,以及条目本身就是附件的情况
    ph = ",".join("?" * len(item_ids))
    rows = cur.execute(
        f"""
        SELECT ia.itemID      AS attachmentItemID,
               ia.parentItemID AS parentItemID,
               ia.linkMode     AS linkMode,
               ia.contentType  AS contentType,
               ia.path         AS path,
               i.key           AS attachmentKey
        FROM itemAttachments ia
        JOIN items i ON i.itemID = ia.itemID
        WHERE ia.parentItemID IN ({ph}) OR ia.itemID IN ({ph})
        """,
        tuple(item_ids) + tuple(item_ids),
    ).fetchall()

    found = []
    missing = []
    seen = set()
    for r in rows:
        if r["attachmentItemID"] in seen:
            continue
        seen.add(r["attachmentItemID"])

        content_type = (r["contentType"] or "")
        path = r["path"] or ""
        is_pdf = content_type == "application/pdf" or path.lower().endswith(".pdf")
        if not is_pdf:
            continue

        link_mode = r["linkMode"]
        if link_mode in LINKMODE_IMPORTED:
            # path 形如 "storage:文件名.pdf"
            filename = path.split(":", 1)[1] if path.startswith("storage:") else path
            src = Path(storage_dir) / r["attachmentKey"] / filename
        elif link_mode == LINKMODE_LINKED_FILE:
            # 链接文件: path 为绝对路径(可能含 attachments: 前缀)
            raw = path.split(":", 1)[1] if ":" in path[:12] else path
            src = Path(raw)
        else:
            continue  # 纯链接 URL, 无本地文件

        if src.exists():
            found.append(src)
        else:
            missing.append((r["attachmentKey"], str(src)))

    return found, missing


def unique_destination(dest_dir, filename):
    """目标文件夹内若重名,自动追加 _1, _2 ... 保证不覆盖。"""
    dest = dest_dir / filename
    if not dest.exists():
        return dest
    stem, suffix = os.path.splitext(filename)
    i = 1
    while True:
        candidate = dest_dir / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def main():
    parser = argparse.ArgumentParser(
        description="把 Zotero 某分类下所有 PDF 整合复制到一个文件夹",
    )
    parser.add_argument("--collection", "-c", required=True,
                        help="Zotero 分类(文件夹)名称, 例如: 体育赛事")
    parser.add_argument("--zotero-dir", "-z", default=str(default_zotero_dir()),
                        help="Zotero 数据目录 (默认: ~/Zotero)")
    parser.add_argument("--output", "-o", default=None,
                        help="目标文件夹 (默认: ./<分类名>_PDF)")
    parser.add_argument("--no-subcollections", action="store_true",
                        help="不包含子分类, 仅当前分类")
    parser.add_argument("--run", action="store_true",
                        help="真正执行复制 (默认只预览 dry-run)")
    args = parser.parse_args()

    zotero_dir = Path(args.zotero_dir)
    sqlite_path = zotero_dir / "zotero.sqlite"
    storage_dir = zotero_dir / "storage"
    output_dir = Path(args.output) if args.output else Path.cwd() / f"{args.collection}_PDF"
    include_sub = not args.no_subcollections

    print("=" * 60)
    print("Zotero 分类 PDF 整合工具")
    print("=" * 60)
    print(f"数据目录 : {zotero_dir}")
    print(f"目标分类 : {args.collection}  (包含子分类: {'是' if include_sub else '否'})")
    print(f"输出文件夹: {output_dir}")
    print(f"模式     : {'执行复制' if args.run else '预览 (dry-run, 加 --run 才真正复制)'}")
    print("-" * 60)

    try:
        conn, tmp_db = open_db_readonly(sqlite_path)
    except FileNotFoundError as e:
        print(f"[错误] {e}")
        print("提示: 用 --zotero-dir 指定你的 Zotero 数据目录"
              "(里面应有 zotero.sqlite 和 storage 文件夹)。")
        sys.exit(1)

    try:
        roots, all_ids = find_collection_ids(conn, args.collection, include_sub)
        if not roots:
            print(f"[错误] 没找到名为「{args.collection}」的分类。")
            print("请确认分类名是否准确(区分大小写/空格)。")
            sys.exit(1)
        if len(roots) > 1:
            print(f"[提示] 找到 {len(roots)} 个同名分类, 将全部一起导出。")
        print(f"涉及分类数(含子分类): {len(all_ids)}")

        found, missing = collect_pdf_attachments(conn, all_ids, storage_dir)
    finally:
        conn.close()
        try:
            os.unlink(tmp_db)
        except OSError:
            pass

    print(f"找到 PDF 数量: {len(found)}")
    if missing:
        print(f"[警告] 有 {len(missing)} 个附件在数据库中存在但文件缺失(可能未同步):")
        for key, p in missing[:10]:
            print(f"   - [{key}] {p}")
        if len(missing) > 10:
            print(f"   ... 以及其它 {len(missing) - 10} 个")

    if not found:
        print("没有可复制的 PDF, 退出。")
        return

    if not args.run:
        print("-" * 60)
        print("以下是将被复制的文件 (预览):")
        for src in found:
            print(f"   {src.name}")
        print("-" * 60)
        print(f"共 {len(found)} 个。确认无误后加 --run 真正执行。")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    copied = 0
    for src in found:
        dest = unique_destination(output_dir, src.name)
        try:
            shutil.copy2(src, dest)
            copied += 1
            print(f"   已复制: {dest.name}")
        except OSError as e:
            print(f"   [失败] {src} -> {e}")

    print("-" * 60)
    print(f"完成! 共复制 {copied}/{len(found)} 个 PDF 到: {output_dir}")


if __name__ == "__main__":
    main()
