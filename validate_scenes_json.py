# -*- coding: utf-8 -*-
import json, importlib.util, os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, "scenes.json"), encoding="utf-8") as f:
    data = json.load(f)

spec = importlib.util.spec_from_file_location("simulate_endings", os.path.join(BASE_DIR, "simulate_endings.py"))
sim = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sim)

errors = []

# Layer 1: structural check against scenes.schema.json (field names, types,
# required scenes, enum values, etc.) -- the same check build_scenes_json.py
# runs automatically on every regeneration; re-run here to re-verify an
# existing file on demand without regenerating it.
try:
    import jsonschema
except ImportError:
    import subprocess
    subprocess.run(["pip", "install", "-q", "jsonschema"], check=True)
    import jsonschema
from jsonschema import Draft202012Validator

with open(os.path.join(BASE_DIR, "scenes.schema.json"), encoding="utf-8") as f:
    schema = json.load(f)
Draft202012Validator.check_schema(schema)
schema_errors = list(Draft202012Validator(schema).iter_errors(
    {"meta": data["meta"], "scenes": data["scenes"]}
))
if schema_errors:
    for e in schema_errors[:30]:
        loc = "/".join(str(p) for p in e.path) or "(root)"
        errors.append(f"[schema] {loc}: {e.message}")
    print(f"SCHEMA CHECK: FAILED ({len(schema_errors)} error(s))")
else:
    print("SCHEMA CHECK: passed (scenes.json conforms to scenes.schema.json)")

# Layer 2: business-logic checks (effects/flags reuse, ending conditions,
# full re-simulation) -- unchanged from before.
scenes = data["scenes"]
meta = data["meta"]

expected_ids = ["P0"] + sim.ORDER + ["EVAL", "E1", "E2", "E3", "E4", "GROWTH_REPORT"]
missing = [s for s in expected_ids if s not in scenes]
extra = [s for s in scenes if s not in expected_ids]
if missing:
    errors.append(f"Missing scenes: {missing}")
if extra:
    errors.append(f"Unexpected scenes: {extra}")

# Check every chapter's decision options match CHAPTERS deltas & flag exactly
FLAG_NORMALIZE = {"اول_عمل": "أول_عمل"}
for chapter in sim.ORDER:
    sc = scenes.get(chapter)
    if not sc:
        continue
    opts = {o["code"]: o for o in sc["decision"]["options"]}
    truth = {code: (delta, flag) for code, delta, flag in sim.CHAPTERS[chapter]}
    if set(opts.keys()) != set(truth.keys()):
        errors.append(f"{chapter}: option codes mismatch {set(opts.keys())} vs {set(truth.keys())}")
        continue
    for code, (delta, flag) in truth.items():
        o = opts[code]
        if o["effects"] != delta:
            errors.append(f"{chapter}/{code}: effects mismatch. json={o['effects']} truth={delta}")
        if flag and "=" in flag:
            var, val = flag.split("=", 1)
            if o.get("setsVariable") != {var: val}:
                errors.append(f"{chapter}/{code}: setsVariable mismatch. json={o.get('setsVariable')} expected={{'{var}':'{val}'}}")
        else:
            expected_flag = FLAG_NORMALIZE.get(flag, flag) if flag else None
            if o.get("flag") != expected_flag:
                errors.append(f"{chapter}/{code}: flag mismatch. json={o.get('flag')} expected={expected_flag}")

# Check C10 mitigation flags match MITIGATION_FLAGS
c10 = scenes["C10"]
mod_flags = set(c10["modifiers"][0]["appliesWhenAnyFlag"])
if mod_flags != sim.MITIGATION_FLAGS:
    errors.append(f"C10 mitigation flags mismatch: {mod_flags} vs {sim.MITIGATION_FLAGS}")
if c10["modifiers"][0]["extraEffects"] != {"M": 4, "P": -5}:
    errors.append("C10 mitigation extraEffects mismatch vs simulate_endings.py (+4 M, -5 P)")

# Check EVAL resolution conditions match resolve_ending() thresholds
eval_res = scenes["EVAL"]["resolution"]["conditions"]
e1 = eval_res["E1"]["all"]
assert {"stat": "R", "gte": 60} in e1
assert {"stat": "B", "gte": 55} in e1
assert {"stat": "D", "gte": 55} in e1
assert {"flag": "مسؤولية"} in e1
assert {"anyFlag": ["استدامة", "تجربة_محدودة", "إعادة_هيكلة"]} in e1
e2 = eval_res["E2"]["all"]
assert {"stat": "C", "gte": 65} in e2
assert {"stat": "M", "gte": 55} in e2
assert {"anyFlag": ["توسع_سريع", "شراكة_سريعة", "دين"]} in e2
e3 = eval_res["E3"]["all"]
assert {"stat": "K", "gte": 60} in e3
assert {"stat": "D", "gte": 55} in e3
assert eval_res["E4"]["default"] is True

# ---------------------------------------------------------------------------
# Full re-simulation using ONLY scenes.json data (not simulate_endings.py's
# CHAPTERS constant) to prove the JSON file alone reproduces the same ending
# distribution as the authoritative simulation.
# ---------------------------------------------------------------------------
def clamp(v):
    return max(0, min(100, v))

def resolve_ending_from_json(stats, flags):
    K, D, C, B, R, M, P = stats["K"], stats["D"], stats["C"], stats["B"], stats["R"], stats["M"], stats["P"]
    if R >= 60 and B >= 55 and D >= 55 and "مسؤولية" in flags and ({"استدامة", "تجربة_محدودة", "إعادة_هيكلة"} & flags):
        return "E1"
    if C >= 65 and M >= 55 and ({"توسع_سريع", "شراكة_سريعة", "دين"} & flags):
        return "E2"
    if K >= 60 and D >= 55:
        return "E3"
    return "E4"

START = {k: v["start"] for k, v in meta["stats"].items()}
mitigation_flags = set(meta["mitigationFlags"])
order = meta["chapterOrder"]
option_lists = [scenes[ch]["decision"]["options"] for ch in order]

import itertools
counts = {"E1": 0, "E2": 0, "E3": 0, "E4": 0}
total = 1
for opts in option_lists:
    total *= len(opts)

for combo in itertools.product(*[range(len(o)) for o in option_lists]):
    stats = dict(START)
    flags = set()
    for ch, idx in zip(order, combo):
        opt = option_lists[order.index(ch)][idx]
        d = dict(opt["effects"])
        if ch == "C10" and (mitigation_flags & flags):
            d["M"] = d.get("M", 0) + 4
            d["P"] = d.get("P", 0) - 5
        for k, v in d.items():
            stats[k] = clamp(stats[k] + v)
        if opt.get("flag"):
            flags.add(opt["flag"])
        if opt.get("setsVariable"):
            for var, val in opt["setsVariable"].items():
                flags.add(f"{var}={val}")
    ending = resolve_ending_from_json(stats, flags)
    counts[ending] += 1

print("Re-simulation from scenes.json ONLY:")
for e in ["E1", "E2", "E3", "E4"]:
    print(f"  {e}: {counts[e]} / {total} ({counts[e]/total*100:.3f}%)")

# Compare against the authoritative simulate_endings.py output
ref_counts = {"E1": 0, "E2": 0, "E3": 0, "E4": 0}
option_counts_ref = [len(sim.CHAPTERS[c]) for c in sim.ORDER]
for combo in itertools.product(*[range(n) for n in option_counts_ref]):
    _, stats, flags, ending = sim.simulate(combo)
    ref_counts[ending] += 1

if counts != ref_counts:
    errors.append(f"Re-simulation mismatch! json-derived={counts} vs simulate_endings.py={ref_counts}")
else:
    print("MATCH: scenes.json reproduces the exact same ending distribution as simulate_endings.py")

if errors:
    print("\nVALIDATION ERRORS:")
    for e in errors:
        print(" -", e)
else:
    print("\nAll validation checks passed.")
