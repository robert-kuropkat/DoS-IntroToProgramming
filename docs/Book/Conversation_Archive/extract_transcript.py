import json, sys

SRC = sys.argv[1]
OUT = sys.argv[2]

def text_from_tool_input(name, inp):
    try:
        if not isinstance(inp, dict):
            return str(inp)[:200]
        # pick a few representative fields, truncate hard
        parts = []
        for k in ("command", "pattern", "query", "url", "file_path", "subject", "description", "prompt", "title", "skill"):
            if k in inp:
                v = str(inp[k])
                if len(v) > 220:
                    v = v[:220] + "...[truncated]"
                parts.append(f"{k}={v}")
        if not parts:
            s = json.dumps(inp)
            if len(s) > 220:
                s = s[:220] + "...[truncated]"
            parts.append(s)
        return "; ".join(parts)
    except Exception:
        return "<unrenderable input>"

def flush(turn, out_f, turn_id):
    if turn is None:
        return
    if not turn["human_text"].strip() and not turn["assistant_segments"]:
        return
    rec = {
        "record_type": "turn",
        "turn_id": turn_id,
        "timestamp_start": turn["ts_start"],
        "timestamp_end": turn["ts_end"],
        "human_text": turn["human_text"],
        "assistant_text": "\n\n".join(turn["assistant_segments"]),
        "tool_calls": turn["tool_calls"],
        "tool_call_count": len(turn["tool_calls"]),
        "commentary": "",
        "tags": [],
    }
    out_f.write(json.dumps(rec, ensure_ascii=False) + "\n")

def main():
    turn = None
    turn_id = 0
    n = 0
    compaction_count = 0
    with open(SRC, "r", encoding="utf-8", errors="replace") as f, \
         open(OUT, "w", encoding="utf-8") as out_f:
        for line in f:
            n += 1
            line = line.strip()
            if not line:
                continue
            try:
                d = json.loads(line)
            except Exception:
                continue
            t = d.get("type")
            ts = d.get("timestamp")

            if t == "system" and d.get("subtype") == "compact_boundary":
                pass  # fallthrough below handles generically

            if t == "system":
                content = d.get("content")
                if content == "Conversation compacted" or "compact" in (d.get("subtype") or ""):
                    compaction_count += 1
                    out_f.write(json.dumps({
                        "record_type": "compaction_boundary",
                        "index": compaction_count,
                        "timestamp": ts,
                    }, ensure_ascii=False) + "\n")
                continue

            if t == "user":
                msg = d.get("message", {})
                content = msg.get("content")
                if isinstance(content, str):
                    # a REAL human turn starts here
                    flush(turn, out_f, turn_id)
                    turn_id += 1
                    turn = {
                        "ts_start": ts, "ts_end": ts,
                        "human_text": content,
                        "assistant_segments": [],
                        "tool_calls": [],
                    }
                # else: tool_result feedback, not a human message -> skip
                continue

            if t == "assistant":
                if turn is None:
                    # assistant activity before any captured human turn (e.g. startup) - skip
                    continue
                turn["ts_end"] = ts
                msg = d.get("message", {})
                content = msg.get("content")
                if isinstance(content, list):
                    for item in content:
                        if not isinstance(item, dict):
                            continue
                        it = item.get("type")
                        if it == "text":
                            txt = item.get("text", "")
                            if txt.strip():
                                turn["assistant_segments"].append(txt)
                        elif it == "tool_use":
                            turn["tool_calls"].append({
                                "name": item.get("name"),
                                "input_brief": text_from_tool_input(item.get("name"), item.get("input")),
                            })
                        # thinking: intentionally skipped (internal scratchpad, not shown to user)
                continue

            # all other types (queue-operation, attachment, last-prompt, mode, atis-latch, cost-state): skip
        flush(turn, out_f, turn_id)
    print("lines read:", n)
    print("turns written:", turn_id)
    print("compaction boundaries:", compaction_count)

if __name__ == "__main__":
    main()
