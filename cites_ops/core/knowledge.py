"""Evidence-preserving local announcement extraction from WhatsApp exports."""

import hashlib
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import yaml

from .chat_parser import ChatParser
from .entity_matcher import EntityMatcher


def extract_candidates(paths, issues, start, end, date_order="dmy"):
    config = yaml.safe_load((Path(__file__).parents[1] / "config/knowledge.yaml").read_text(encoding="utf-8"))
    patterns = {kind: [re.compile(p, re.I | re.S) for p in pats] for kind, pats in config["types"].items()}
    exclusions = [re.compile(p, re.I | re.S) for p in config["exclude"]]
    matcher = EntityMatcher()
    frame = pd.DataFrame([{"Id": i["id"], "Summary": i["summary"], "Description": i["description"]} for i in issues])
    index = matcher.build_issue_index(frame)
    by_id = {i["id"]: i for i in issues}
    functions = sorted({i["functionality"] for i in issues})
    function_patterns = {f: re.compile(r"\b" + r"[\s_\-]*".join(re.escape(s) for s in re.findall(r"[A-Za-z]+|\d+", f)) + r"\b", re.I) for f in functions}

    def references(text):
        ids = set()
        for ent in matcher.extract_from_text(text):
            ids.update(index.get((ent["entity_type"], ent["value"]), []))
        for match in re.finditer(r"\b(?:ticket|issue|mantis)(?:\s*(?:id|no\.?|number))?\s*[:#-]?\s*(\d+)\b", text, re.I):
            key = str(int(match[1]))
            if key in by_id:
                ids.add(key)
        funcs = {f for f, pattern in function_patterns.items() if pattern.search(text)}
        funcs.update(by_id[i]["functionality"] for i in ids)
        return ids, funcs

    def sanitize(text):
        text = matcher.mask_pii(text)
        text = re.sub(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", "[email masked]", text)
        return re.sub(r"(?<!\w)(?:\+\d{1,3}[ -]?)?\d(?:[ -]?\d){9,11}(?!\w)", "[identifier masked]", text)

    entries, warnings = {}, []
    for path in paths:
        messages = ChatParser.parse_file(path)
        if messages.empty:
            warnings.append(f"No readable messages in {Path(path).name}.")
            continue
        recent, invalid, last_source = [], 0, None
        for number, row in enumerate(messages.to_dict("records"), 1):
            source = Path(path).name + (" / " + row["source_file"] if row.get("source_file") else "")
            if last_source != source:
                recent = []
            last_source = source
            raw_date = re.sub(r"[.\-]", "/", row["date_str"])
            stamp = None
            for year in ("%Y", "%y"):
                try:
                    prefix = "%d/%m/" if date_order == "dmy" else "%m/%d/"
                    parsed = datetime.strptime(raw_date, prefix + year)
                    time = re.sub(r"\s+", " ", row["time_str"].strip()).upper()
                    for fmt in ("%I:%M %p", "%I:%M:%S %p", "%H:%M", "%H:%M:%S"):
                        try:
                            t = datetime.strptime(time, fmt).time()
                            stamp = datetime.combine(parsed.date(), t)
                            break
                        except ValueError:
                            pass
                    break
                except ValueError:
                    pass
            if stamp is None:
                invalid += 1
                continue
            if not start <= stamp.date() <= end:
                continue
            msg = row["message"]
            ids, funcs = references(msg)
            recent = [r for r in recent if timedelta(0) <= stamp - r[0] <= timedelta(hours=1)]
            kind = next((k for k, pats in patterns.items() if any(p.search(msg) for p in pats)), None)
            if kind and not any(p.search(msg) for p in exclusions):
                context = [r for r in recent if (ids & r[2]) or (funcs & r[3])][-2:]
                evidence = [{"source": source, "message": number, "date": stamp.isoformat(sep=" "),
                             "text": sanitize(msg)}]
                evidence[:0] = [r[1] for r in context]
                key = hashlib.sha256((stamp.isoformat() + row["sender"] + re.sub(r"\s+", " ", msg).casefold()).encode()).hexdigest()
                candidate = {
                    "id": key, "type": kind, "date": stamp.date().isoformat(),
                    "functionalities": sorted(funcs) or ["Needs identification"],
                    "title": sanitize(msg.splitlines()[0])[:160],
                    "text": sanitize(msg), "tickets": sorted(ids),
                    "state": "Candidate — needs review", "evidence": evidence,
                }
                entries.setdefault(key, candidate)
            recent.append((stamp, {"source": source, "message": number,
                                   "date": stamp.isoformat(sep=" "), "text": sanitize(msg)}, ids, funcs))
        if invalid:
            warnings.append(f"Skipped {invalid} messages with unreadable timestamps in {Path(path).name}.")
    return sorted(entries.values(), key=lambda e: (e["date"], e["id"]), reverse=True), warnings
