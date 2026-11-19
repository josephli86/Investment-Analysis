from pathlib import Path

src_dir = Path("Src")  # 改为你本地解压后的 Src 文件夹路径
out_file = Path("Src_merged.txt")

with out_file.open("w", encoding="utf-8") as out:
    for p in sorted(src_dir.rglob("*")):
        if p.suffix.lower() in {".cls", ".bas"}:
            raw = p.read_bytes()
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("gb18030", errors="replace")
            out.write(f"\n{'='*60}\nFILE: {p.relative_to(src_dir)}\n{'='*60}\n{text}\n")

print(f"已生成: {out_file.resolve()}")