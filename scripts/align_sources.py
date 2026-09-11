#!/usr/bin/env python3
"""双源转录对齐工具 —— 让程序比对，主上下文只看汇总表和低置信轮次。

子命令
  parse   FILE                  检查解析结果（轮数、发言人分布、前几轮）
  align   BASE REF [--out DIR]  跑三张表 + 逐轮投票，逐轮明细写入 DIR/turns.tsv
  sample  BASE REF --speaker X  并排打印 X 的长句（基准稿 vs 参考稿），用来判断哪份文本更准
  reverse-check BASE CLEANED    收尾：列出原文 ≥N 字但成稿覆盖 <阈值 的轮次（抓漏段）

输入格式
  .docx  —— python-docx 按段落读
  .md/.txt —— 按行读
  .json  —— [{"speaker": "...", "text": "..."}, ...]（解析器认不出格式时的兜底）
  发言人行识别：独占一行的短标签（可带编号/时间戳），或行内 `【标签】内容` / `[13]标签 内容` / `标签：内容`。
  认不出时用 --speaker-regex 传一个带 (?P<spk>...) 的正则（匹配整行=发言人行）。
"""
import argparse
import csv
import difflib
import json
import os
import re
import sys
from collections import Counter, defaultdict

# ---------------------------------------------------------------- parsing
TS = r"(?:\(?\d{1,2}:\d{2}(?::\d{2})?\)?)"
SPK_LINE = re.compile(
    rf"^\s*(?:\[\d+\]\s*)?(?P<spk>[^\s\d:：，。,.!?！？【】\[\]#>*][^:：，。,.!?！？]{{0,11}}?)\s*{TS}?\s*[:：]?\s*$"
)
SPK_ANON = re.compile(rf"^\s*(?P<spk>(?:发言人|说话人|Speaker|speaker)\s*[\w]+)\s*{TS}?\s*[:：]?\s*$")
INLINE = [
    re.compile(r"^\s*【(?P<spk>[^】]{1,12})】\s*(?P<text>.+)$"),
    re.compile(r"^\s*\[\d+\]\s*(?P<spk>[^\s:：]{1,12})[\s:：]+(?P<text>.+)$"),
    re.compile(r"^\s*(?P<spk>[^\s:：，。,]{1,8})\s*[:：]\s*(?P<text>.{6,})$"),
]
SKIP = re.compile(r"^\s*(?:>|#|---|\*\*|关键词|章节|摘要|会议时间|时长|录音)")


def read_lines(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        import docx  # python-docx

        return [p.text for p in docx.Document(path).paragraphs]
    with open(path, encoding="utf-8-sig") as f:
        return f.read().splitlines()


def _speaker_match(line, custom):
    """返回 (label, 是否强信号) 或 None。强信号 = 自定义正则 / 匿名编号 / 带时间戳。"""
    if custom:
        m = custom.match(line)
        return (m.group("spk").strip(), True) if m else None
    m = SPK_ANON.match(line)
    if m:
        return m.group("spk").strip(), True
    if len(line) <= 24:
        m = SPK_LINE.match(line)
        if m:
            return m.group("spk").strip(), bool(re.search(TS + r"\s*[:：]?\s*$", line))
    return None


def parse(path, speaker_regex=None):
    """返回 [(speaker, text), ...]"""
    if path.lower().endswith(".json"):
        data = json.load(open(path, encoding="utf-8"))
        return [(d["speaker"], d["text"]) for d in data]
    custom = re.compile(speaker_regex) if speaker_regex else None
    lines = [l.strip() for l in read_lines(path)]
    lines = [l for l in lines if l and not SKIP.match(l)]
    # 第一遍：弱信号标签（无时间戳的短行）至少出现 2 次才算发言人，避免把残句当标签
    freq = Counter(m[0] for m in (_speaker_match(l, custom) for l in lines) if m)
    turns, cur_spk, buf = [], None, []

    def flush():
        if cur_spk is not None and "".join(buf).strip():
            turns.append((cur_spk, "".join(buf)))

    for line in lines:
        m = _speaker_match(line, custom)
        if m and not m[1]:
            # 弱信号：只出现一次，或紧跟在发言人行后面（该轮还没内容）→ 当正文
            if freq[m[0]] < 2 or (cur_spk is not None and not buf):
                m = None
        if m:
            flush()
            cur_spk, buf = m[0], []
            continue
        if not custom:
            im = next((pat.match(line) for pat in INLINE if pat.match(line)), None)
            if im:
                flush()
                cur_spk, buf = im.group("spk").strip(), [im.group("text")]
                continue
        buf.append(line)
    flush()
    return turns


# ---------------------------------------------------------------- alignment core
KEEP = re.compile(r"[\w一-鿿]")


def stream(turns):
    """拼成去标点字符流，并记录每个字符属于第几轮。"""
    chars, owner = [], []
    for i, (_, t) in enumerate(turns):
        for ch in t:
            if KEEP.match(ch):
                chars.append(ch.lower())
                owner.append(i)
    return "".join(chars), owner


def turn_len(turns):
    return [sum(1 for ch in t if KEEP.match(ch)) for _, t in turns]


def match_blocks(a, b):
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    return [blk for blk in sm.get_matching_blocks() if blk.size >= 4]


def votes(base, ref, blocks):
    """votes[i] = Counter{ref_turn_idx: 匹配字符数}，i 为基准稿轮次"""
    _, bo = stream(base)
    _, ro = stream(ref)
    v = defaultdict(Counter)
    for blk in blocks:
        for k in range(blk.size):
            v[bo[blk.a + k]][ro[blk.b + k]] += 1
    return v


# ---------------------------------------------------------------- reports
def pct(x):
    return f"{100 * x:.0f}%"


def cmd_parse(args):
    t = parse(args.file, args.speaker_regex)
    c = Counter(s for s, _ in t)
    print(f"{len(t)} 轮，{sum(turn_len(t))} 字")
    for s, n in c.most_common():
        print(f"  {s}: {n} 轮")
    for s, x in t[: args.head]:
        print(f"  [{s}] {x[:60]}")


def label_ranges(turns):
    r = {}
    for i, (s, _) in enumerate(turns):
        lo, hi, n = r.get(s, (i, i, 0))
        r[s] = (min(lo, i), max(hi, i), n + 1)
    return r


def alias_candidates(turns, name):
    """零重叠的标签对 → 可能是同一人换麦/声纹丢失后的第二标签。"""
    r = label_ranges(turns)
    labs = [s for s in r if r[s][2] >= 3]
    out = []
    for i, x in enumerate(labs):
        for y in labs[i + 1:]:
            (a0, a1, _), (b0, b1, _) = r[x], r[y]
            if a1 < b0 or b1 < a0:
                out.append((x, y, r[x], r[y]))
    if out:
        print(f"\n## {name}：零重叠标签对（疑似同一人两个标签，需人工确认）")
        for x, y, rx, ry in out:
            print(f"  {x} 轮{rx[0]}–{rx[1]}（{rx[2]}轮）  ⟷  {y} 轮{ry[0]}–{ry[1]}（{ry[2]}轮）")


def near_dups(turns, name, thr):
    hits = []
    for i in range(len(turns) - 1):
        a, b = turns[i][1], turns[i + 1][1]
        if min(len(a), len(b)) < 6:
            continue
        r = difflib.SequenceMatcher(None, a, b).ratio()
        if r > thr:
            hits.append((i, r))
    if hits:
        print(f"\n## {name}：相邻近重复段（多麦串音常见“残缺版+完整版”两轮，比值>{thr}）")
        for i, r in hits[:40]:
            print(f"  轮{i}[{turns[i][0]}] / 轮{i+1}[{turns[i+1][0]}]  ratio={r:.2f}  {turns[i+1][1][:30]}")
        if len(hits) > 40:
            print(f"  …共 {len(hits)} 处")


def find_runs(turns, covered, min_chars):
    """连续未被对方覆盖、且总字数较长的区段 → 没录到 / 外来声音 / 电话接入。"""
    lens = turn_len(turns)
    runs, start = [], None
    for i in range(len(turns) + 1):
        bad = i < len(turns) and covered[i] < 0.2 * max(lens[i], 1)
        if bad and start is None:
            start = i
        if not bad and start is not None:
            n = sum(lens[start:i])
            if n >= min_chars:
                runs.append((start, i - 1, n))
            start = None
    return runs


def print_runs(turns, runs, name, min_chars):
    if runs:
        print(f"\n## {name}：对方未覆盖的长片段（≥{min_chars}字）")
        for s, e, n in runs:
            spk = Counter(turns[k][0] for k in range(s, e + 1)).most_common(3)
            print(f"  轮{s}–{e}，{n}字，标签 {spk}  起句：{turns[s][1][:30]}")


def cmd_align(args):
    base = parse(args.base, args.speaker_regex)
    ref = parse(args.ref, args.speaker_regex)
    a, bo = stream(base)
    b, ro = stream(ref)
    print(f"基准稿 {len(base)} 轮 {len(a)} 字；参考稿 {len(ref)} 轮 {len(b)} 字")
    blocks = match_blocks(a, b)
    v = votes(base, ref, blocks)
    blens = turn_len(base)

    # ---- 表1 标签交叉表
    cross, rtot = defaultdict(Counter), defaultdict(Counter)
    for i, c in v.items():
        for j, n in c.items():
            cross[base[i][0]][ref[j][0]] += n
            rtot[ref[j][0]][base[i][0]] += n
    print("\n## 表1 标签交叉表（按匹配字符数）")
    print("  基准标签 → 参考标签：")
    for bs in sorted(cross, key=lambda x: -sum(cross[x].values())):
        c, tot = cross[bs], sum(cross[bs].values())
        print(f"    {bs}（{tot}字）→ " + "  ".join(f"{rs} {pct(n / tot)}" for rs, n in c.most_common(4)))
    print("  参考标签 → 基准标签：")
    for rs in sorted(rtot, key=lambda x: -sum(rtot[x].values())):
        c, tot = rtot[rs], sum(rtot[rs].values())
        print(f"    {rs}（{tot}字）→ " + "  ".join(f"{bs} {pct(n / tot)}" for bs, n in c.most_common(3)))
    if blocks:
        f, l = blocks[0], blocks[-1]
        print(f"  匹配区间：基准稿 轮{bo[f.a]}–{bo[l.a + l.size - 1]}，参考稿 轮{ro[f.b]}–{ro[l.b + l.size - 1]}")

    # 某参考标签分散在多个基准标签上、且集中在一段连续轮次 → 电话/免提接入者被剁碎
    span = defaultdict(list)
    for i, c in v.items():
        for j in c:
            span[ref[j][0]].append(i)
    for rs, c in rtot.items():
        tot = sum(c.values())
        if len(c) >= 2 and c.most_common(1)[0][1] / tot < 0.7:
            lo, hi = min(span[rs]), max(span[rs])
            where = f"集中在基准轮{lo}–{hi}" if hi - lo <= 3 * len(set(span[rs])) else f"散布在基准轮{lo}–{hi}"
            print(f"  ⚠ 参考「{rs}」分散在 {len(c)} 个基准标签上（{'、'.join(c)}），{where}："
                  f"若紧跟在“某人被点名打电话”之后，就是电话/免提那头的人被剁碎了。")
    # 参考标签 → 基准名：一对一时才改名；多个参考标签落在同一基准标签上时，只有最大的那个改名
    by_base = defaultdict(list)
    for rs, c in rtot.items():
        by_base[c.most_common(1)[0][0]].append(rs)
    mapping = {}
    for bs, rss in by_base.items():
        rss.sort(key=lambda x: -rtot[x][bs])
        mapping[rss[0]] = bs
        for extra in rss[1:]:
            mapping[extra] = extra
        if len(rss) > 1:
            print(f"  ⚠ 参考稿 {'、'.join(rss)} 都主要落在基准「{bs}」上：要么「{bs}」是共享设备主人的标签"
                  f"（屋里未匹配声纹的人被并进来），要么 {'、'.join(rss[1:])} 是同一人的第二个声纹桶。判据见 reference。")
    base_to_ref = defaultdict(list)
    for bs, c in cross.items():
        base_to_ref[c.most_common(1)[0][0]].append(bs)
    for rs, bss in base_to_ref.items():
        if len(bss) > 1:
            print(f"  ⚠ 基准稿 {'、'.join(bss)} 都主要落在参考「{rs}」上：疑似同一人换麦/声纹丢失后的两个标签（看下方零重叠检查）。")

    # ---- 表2 分桶覆盖率
    covered = [sum(v[i].values()) if i in v else 0 for i in range(len(base))]
    print(f"\n## 表2 分桶覆盖率（每 {args.bucket} 轮，基准稿字符被参考稿匹配的比例）")
    print("  整齐掉到 0% 的连续桶 = 参考稿没录到这段，不是对齐失败，不要去调参数。")
    for st in range(0, len(base), args.bucket):
        e = min(st + args.bucket, len(base))
        cov = sum(covered[st:e]) / (sum(blens[st:e]) or 1)
        print(f"  轮{st:>4}–{e - 1:<4} {pct(cov):>4} {'█' * int(cov * 20)}")

    # ---- 逐轮投票
    runs = find_runs(base, covered, args.run_chars)
    in_run = {k for s0, e0, _ in runs for k in range(s0, e0 + 1)}
    rows, pairs = [], Counter()
    for i, (bs, t) in enumerate(base):
        c = v.get(i, Counter())
        tot = sum(c.values())
        lab = Counter()
        for j, n in c.items():
            lab[mapping.get(ref[j][0], ref[j][0])] += n
        top, n = lab.most_common(1)[0] if lab else ("", 0)
        conf, cov = (n / tot if tot else 0), tot / (blens[i] or 1)
        flag = ""
        if i in in_run:
            flag = "参考稿未录到"
        elif blens[i] >= args.min_chars:
            if cov < 0.2:
                flag = "未覆盖"
            elif conf < args.conf:
                flag = "票数混杂"
            elif top != bs:
                flag = "标签不一致"
                pairs[(bs, top)] += 1
        rows.append([i, bs, top, f"{conf:.2f}", f"{cov:.2f}", ",".join(map(str, sorted(c))), flag, t])
    systematic = {p for p, k in pairs.items() if k >= args.sys_min}
    for r in rows:
        if r[6] == "标签不一致" and (r[1], r[2]) in systematic:
            r[6] = "系统性改标"
    os.makedirs(args.out, exist_ok=True)
    path = os.path.join(args.out, "turns.tsv")
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh, delimiter="\t")
        w.writerow(["轮", "基准标签", "投票标签", "置信度", "覆盖率", "参考稿轮次", "标记", "文本"])
        w.writerows(rows)

    fc = Counter(r[6] for r in rows)
    judged = len(base) - fc["参考稿未录到"]
    manual = fc["票数混杂"] + fc["标签不一致"] + fc["未覆盖"]
    print(f"\n## 逐轮投票（参考稿覆盖到的 {judged} 轮里）")
    print(f"  自动定下 {judged - manual} 轮（{pct((judged - manual) / max(judged, 1))}），需人工看 {manual} 轮："
          f"票数混杂 {fc['票数混杂']} / 零星标签不一致 {fc['标签不一致']} / 零星未覆盖 {fc['未覆盖']}")
    if systematic:
        print(f"  系统性改标（同一对 ≥{args.sys_min} 轮，先对照表1确认映射，确认后批量改）：")
        for (bs, top), k in sorted(pairs.items(), key=lambda x: -x[1]):
            if (bs, top) in systematic:
                print(f"    {bs} → {top}：{k} 轮")
    if fc["参考稿未录到"]:
        print(f"  参考稿未录到：{fc['参考稿未录到']} 轮（这些轮只能靠基准稿＋上下文判断）")
    print(f"  明细：{path}。只筛“标记”为 票数混杂/标签不一致/未覆盖 的行来看，不要整表读进上下文。")

    # ---- 纯投票会漏掉的坑：辅助线索
    alias_candidates(base, "基准稿")
    alias_candidates(ref, "参考稿")
    near_dups(base, "基准稿", args.dup)
    print_runs(base, runs, "基准稿", args.run_chars)
    rv = defaultdict(int)
    for c in v.values():
        for j, n in c.items():
            rv[j] += n
    print_runs(ref, find_runs(ref, [rv.get(j, 0) for j in range(len(ref))], args.run_chars), "参考稿", args.run_chars)


def cmd_sample(args):
    base = parse(args.base, args.speaker_regex)
    ref = parse(args.ref, args.speaker_regex)
    a, _ = stream(base)
    b, _ = stream(ref)
    v = votes(base, ref, match_blocks(a, b))
    cands = [i for i, (s, t) in enumerate(base) if s == args.speaker and i in v and len(t) >= 30]
    cands.sort(key=lambda i: -len(base[i][1]))
    if not cands:
        print(f"基准稿里没有 {args.speaker} 的长句被匹配到。可用标签：{sorted(set(s for s, _ in base))}")
        return
    print("看专有名词和数字哪份更对；哪份更准就用哪份当该发言人的措辞来源。\n")
    for i in sorted(cands[: args.n]):
        js = sorted(v[i])
        print(f"— 基准 轮{i} [{base[i][0]}]：{base[i][1]}")
        print(f"  参考 轮{js[0]}–{js[-1]} [{'/'.join(sorted(set(ref[j][0] for j in js)))}]："
              f"{''.join(ref[j][1] for j in range(js[0], js[-1] + 1))}\n")


def cmd_reverse(args):
    base = parse(args.base, args.speaker_regex)
    out = parse(args.cleaned, args.speaker_regex)
    a, _ = stream(base)
    b, _ = stream(out)
    v = votes(base, out, match_blocks(a, b))
    lens = turn_len(base)
    miss = []
    for i, (s, t) in enumerate(base):
        if lens[i] < args.min_chars:
            continue
        cov = sum(v.get(i, Counter()).values()) / lens[i]
        if cov < args.max_cov:
            miss.append((i, s, cov, t))
    print(f"原文 {len(base)} 轮；≥{args.min_chars}字 且成稿覆盖 <{pct(args.max_cov)} 的：{len(miss)} 轮")
    print("逐条判断：是有意删除（闲聊/串音/重复）还是写稿漏掉。\n")
    for i, s, cov, t in miss:
        print(f"  轮{i} [{s}] 覆盖{pct(cov)}：{t[:80]}")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--speaker-regex", help="发言人整行正则，需含 (?P<spk>...)")
    sp = p.add_subparsers(dest="cmd", required=True)

    x = sp.add_parser("parse", parents=[common]); x.add_argument("file"); x.add_argument("--head", type=int, default=5)
    x.set_defaults(fn=cmd_parse)

    x = sp.add_parser("align", parents=[common]); x.add_argument("base"); x.add_argument("ref")
    x.add_argument("--out", default="align_out")
    x.add_argument("--bucket", type=int, default=25)
    x.add_argument("--conf", type=float, default=0.7, help="低于此票数占比 → 票数混杂")
    x.add_argument("--min-chars", type=int, default=6, help="短于此字数的轮不标人工")
    x.add_argument("--dup", type=float, default=0.35, help="相邻近重复阈值")
    x.add_argument("--run-chars", type=int, default=150, help="未覆盖长片段的最小字数")
    x.add_argument("--sys-min", type=int, default=3, help="同一对标签不一致达到此轮数 → 系统性改标")
    x.set_defaults(fn=cmd_align)

    x = sp.add_parser("sample", parents=[common]); x.add_argument("base"); x.add_argument("ref")
    x.add_argument("--speaker", required=True); x.add_argument("--n", type=int, default=10)
    x.set_defaults(fn=cmd_sample)

    x = sp.add_parser("reverse-check", parents=[common]); x.add_argument("base"); x.add_argument("cleaned")
    x.add_argument("--min-chars", type=int, default=8); x.add_argument("--max-cov", type=float, default=0.5)
    x.set_defaults(fn=cmd_reverse)

    args = p.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
