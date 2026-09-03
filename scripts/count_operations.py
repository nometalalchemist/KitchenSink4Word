"""Count distinct document OPERATIONS across the consolidated v2 surface.

The v2 build folds many v1 tools into action-dispatching multiplexers
(manage_note action=insert|edit|delete|..., list_elements type=...,
validate checks=..., modify_table_structure action=..., convert_notes
direction=..., etc.). The public HUMAN-facing headline is the number of
distinct operations, not the tool count: one multiplexer tool performs
several distinct operations, and this script sums them so the published
figure is script-generated, never hand-math.

Definition of an operation (documented, so the number is defensible):
each of the 108 document tools contributes its distinct operation count,
computed as the MAX of two independent, committed-source detectors:

  1. SOURCE DISPATCH SCAN: the distinct values of the tool's primary
     dispatch parameter (action / operation / check / direction / type /
     mode / kind), extracted automatically from the tool's own validation
     code (the tuple, set, dict, membership test, equality chain, or
     table lookup the tool checks its dispatch argument against). When a
     tool passes its dispatch parameter straight through to an ops-module
     function (e.g. manage_backups -> ops.backups.manage_backups), the
     scan follows the call and reads the validation there, so dispatch
     values validated one layer down are still counted. A single-purpose
     tool contributes 1.
  2. MIGRATION CAPABILITY MAP: the number of distinct injected-argument
     signatures that migration/v1_to_v2.json folds into that tool (each
     is one preserved v1.6 capability). This is the project's vetted,
     test-guarded machine-readable capability enumeration.

Taking the per-tool MAX unions both views: the source scan catches v2
dispatch values that expand past v1 (extra break types, reference-list
kinds), and the migration map catches secondary discriminators the scan
does not multiply (footnote-vs-endnote, rows-vs-columns). The result is
a defensible lower bound on the distinct operations the v2 surface can
perform, drawn entirely from two committed sources with zero hand-math.

The two pack-toggle tools (enable_tools, disable_tools) are surface
plumbing, not document operations, and are excluded.

Run:  .venv/Scripts/python.exe -X utf8 scripts/count_operations.py
"""

from __future__ import annotations

import ast
import json
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from word_mcp import server as server_mod  # noqa: E402

SERVER_PY = SRC / "word_mcp" / "server.py"
MIGRATION_JSON = ROOT / "migration" / "v1_to_v2.json"

# Parameter names that name a PRIMARY dispatch dimension. The first one of
# these that a tool actually validates against a value-set is the dimension
# whose distinct values are counted as operations.
DISPATCH_NAMES = (
    "action", "operation", "op", "check", "direction", "mode", "kind", "type",
)

# Plumbing, not document operations.
EXCLUDE = {"enable_tools", "disable_tools"}


def _decorated_with_tool(node: ast.FunctionDef) -> bool:
    for dec in node.decorator_list:
        target = dec.func if isinstance(dec, ast.Call) else dec
        if isinstance(target, ast.Name) and target.id == "_tool":
            return True
        if isinstance(target, ast.Attribute) and target.attr == "_tool":
            return True
    return False


def _literal_values(node: ast.AST):
    """String values from a tuple/list/set literal or a dict literal's keys."""
    if isinstance(node, (ast.Tuple, ast.List, ast.Set)):
        vals = [el.value for el in node.elts
                if isinstance(el, ast.Constant) and isinstance(el.value, str)]
        return vals if vals and len(vals) == len(node.elts) else None
    if isinstance(node, ast.Dict):
        vals = [k.value for k in node.keys
                if isinstance(k, ast.Constant) and isinstance(k.value, str)]
        return vals if vals and len(vals) == len(node.keys) else None
    return None


def _local_map(fn: ast.FunctionDef):
    """Map local names assigned a tuple/list/set/dict of string constants."""
    out: dict[str, list[str]] = {}
    for node in ast.walk(fn):
        if isinstance(node, ast.Assign):
            vals = _literal_values(node.value)
            if vals:
                for tgt in node.targets:
                    if isinstance(tgt, ast.Name):
                        out[tgt.id] = vals
    return out


def _scan_param_values(fn: ast.FunctionDef, wanted: tuple[str, ...],
                       module_obj=server_mod) -> dict[str, set[str]]:
    """Union every validation signal for each wanted parameter name inside
    one function:
      - `X in/not in <set|Name>`            (membership tests)
      - `X == "const"` / `"const" == X`     (if/elif dispatch chains)
      - `DICT.get(X)`, `DICT[X]`, `X in DICT` (table-driven dispatch)
    Name comparators resolve against local assignments first, then the
    given module's constants."""
    local = _local_map(fn)
    found: dict[str, set[str]] = {}

    def module_values(name: str):
        val = getattr(module_obj, name, None)
        if isinstance(val, dict):
            return [k for k in val.keys() if isinstance(k, str)]
        if isinstance(val, (tuple, list, set)):
            return [v for v in val if isinstance(v, str)]
        return None

    def resolve(node: ast.AST):
        lit = _literal_values(node)
        if lit is not None:
            return lit
        if isinstance(node, ast.Name):
            return local.get(node.id) or module_values(node.id)
        return None

    def add(param: str, vals):
        if vals:
            found.setdefault(param, set()).update(vals)

    for sub in ast.walk(fn):
        # Membership: X in/not in <set|Name|DICT>
        if isinstance(sub, ast.Compare) and len(sub.ops) == 1:
            op = sub.ops[0]
            left = sub.left
            comp = sub.comparators[0]
            if isinstance(op, (ast.In, ast.NotIn)) and isinstance(left, ast.Name) \
                    and left.id in wanted:
                add(left.id, resolve(comp))
            # Equality dispatch: X == "const"  /  "const" == X
            if isinstance(op, (ast.Eq,)):
                for a, b in ((left, comp), (comp, left)):
                    if isinstance(a, ast.Name) and a.id in wanted \
                            and isinstance(b, ast.Constant) \
                            and isinstance(b.value, str):
                        add(a.id, [b.value])
        # Table-driven: DICT.get(X) or DICT[X]
        if isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute) \
                and sub.func.attr == "get" and sub.args:
            arg = sub.args[0]
            base = sub.func.value
            if isinstance(arg, ast.Name) and arg.id in wanted \
                    and isinstance(base, ast.Name):
                add(arg.id, resolve(base))
        if isinstance(sub, ast.Subscript) and isinstance(sub.value, ast.Name):
            idx = sub.slice
            if isinstance(idx, ast.Name) and idx.id in wanted:
                add(idx.id, resolve(sub.value))

    return found


def _ops_index():
    """(module_name -> (ast func map, imported module)) for word_mcp.ops.*,
    plus the alias map server.py imports them under (backups as _bk, ...)."""
    import importlib

    tree = ast.parse(SERVER_PY.read_text(encoding="utf-8"))
    aliases: dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module == "ops":
            for a in node.names:
                aliases[a.asname or a.name] = a.name
    modules: dict[str, tuple[dict[str, ast.FunctionDef], object]] = {}
    for modname in set(aliases.values()):
        path = SRC / "word_mcp" / "ops" / f"{modname}.py"
        if not path.exists():
            continue
        mtree = ast.parse(path.read_text(encoding="utf-8"))
        fns = {n.name: n for n in ast.walk(mtree)
               if isinstance(n, ast.FunctionDef)}
        modules[modname] = (fns, importlib.import_module(
            f"word_mcp.ops.{modname}"))
    return aliases, modules


def _delegate_values(fn: ast.FunctionDef, aliases, modules):
    """Follow calls that pass a dispatch-named parameter through to an
    ops-module function (positionally or as a keyword) and scan the callee's
    validation for that parameter. Returns {tool_param: set[values]}."""
    out: dict[str, set[str]] = defaultdict(set)
    for sub in ast.walk(fn):
        if not (isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute)
                and isinstance(sub.func.value, ast.Name)
                and sub.func.value.id in aliases):
            continue
        modname = aliases[sub.func.value.id]
        if modname not in modules:
            continue
        fns, mod_obj = modules[modname]
        callee = fns.get(sub.func.attr)
        if callee is None:
            continue
        pos = [a.arg for a in callee.args.args]
        pairs = []  # (callee_param_name, passed node)
        for i, arg in enumerate(sub.args):
            if i < len(pos):
                pairs.append((pos[i], arg))
        for kw in sub.keywords:
            if kw.arg:
                pairs.append((kw.arg, kw.value))
        for callee_param, node in pairs:
            if isinstance(node, ast.Name) and node.id in DISPATCH_NAMES:
                vals = _scan_param_values(
                    callee, (callee_param,), mod_obj).get(callee_param)
                if vals:
                    out[node.id] |= vals
    return out


def _dispatch_values(fn: ast.FunctionDef, aliases, modules):
    """Return (param_name, sorted[values]) for the primary dispatch dimension,
    or None. Unions the tool function's own validation with any validation
    found by following pass-through calls into ops modules. Picks the
    dispatch param with the largest value union; ties break by
    DISPATCH_NAMES priority so `action` beats a minor secondary `type`."""
    found = _scan_param_values(fn, DISPATCH_NAMES)
    for param, vals in _delegate_values(fn, aliases, modules).items():
        found.setdefault(param, set()).update(vals)
    if not found:
        return None
    best = max(found, key=lambda p: (len(found[p]),
                                     -DISPATCH_NAMES.index(p)))
    return best, sorted(found[best])


def _migration_ops():
    """Distinct injected-argument signatures folded into each v2 tool."""
    data = json.loads(MIGRATION_JSON.read_text(encoding="utf-8"))
    per_dest: dict[str, set] = defaultdict(set)
    for entry in data["tools"].values():
        inj = entry.get("inject") or {}
        per_dest[entry["to"]].add(json.dumps(inj, sort_keys=True))
    return {dest: len(sigs) for dest, sigs in per_dest.items()}


def main() -> None:
    tree = ast.parse(SERVER_PY.read_text(encoding="utf-8"))
    tool_fns = {
        n.name: n
        for n in ast.walk(tree)
        if isinstance(n, ast.FunctionDef) and _decorated_with_tool(n)
    }
    mig = _migration_ops()
    aliases, modules = _ops_index()

    total = 0
    multiplex = 0
    single = 0
    rows = []
    for name in sorted(tool_fns):
        if name in EXCLUDE:
            continue
        disp = _dispatch_values(tool_fns[name], aliases, modules)
        scan = len(disp[1]) if disp else 1
        migc = mig.get(name, 1)
        ops = max(scan, migc)
        detail = ""
        if disp:
            detail = f"{disp[0]}: {', '.join(disp[1])}"
        if ops > 1:
            multiplex += 1
        else:
            single += 1
        rows.append((name, scan, migc, ops, detail))
        total += ops

    counted_tools = len(tool_fns) - len(EXCLUDE & set(tool_fns))
    print(f"{'tool':<28} {'scan':>4} {'mig':>4} {'ops':>4}  dispatch values")
    print("-" * 90)
    for name, scan, migc, ops, detail in rows:
        d = detail if len(detail) <= 40 else detail[:37] + "..."
        print(f"{name:<28} {scan:>4} {migc:>4} {ops:>4}  {d}")
    print("-" * 90)
    print(f"tools counted (excl. enable/disable_tools): {counted_tools}")
    print(f"  single-operation tools:                   {single}")
    print(f"  multi-operation tools:                    {multiplex}")
    print(f"TOTAL DISTINCT OPERATIONS:                  {total}")


if __name__ == "__main__":
    main()
