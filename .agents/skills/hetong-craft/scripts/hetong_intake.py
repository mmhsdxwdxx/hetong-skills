#!/usr/bin/env python3
"""Contract intake plans, explicit UI receipts, and localhost forms (stdlib only)."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
from pathlib import Path
import secrets
import sys
import time
from urllib.parse import parse_qs, urlsplit
import webbrowser

# Some hosts run Python with -I; load only this verified skill's own sibling module.
SCRIPT_ROOT = Path(__file__).resolve().parent
if str(SCRIPT_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRIPT_ROOT))
from buyer_clause_select import DEFAULT_CATALOG, SELLER_CATALOG, load_catalog, list_categories, select_categories, render_modules

SCHEMA_VERSION = 1
NONE = "__none__"
# These values select clause families, never payment or guarantee terms.
BUYER_SCOPE = {
    "A": ("保交条款",),
    "B": (),
    "C": ("保交条款", "共管条款"),
    "D": ("共管条款",),
}
BUYER_SCOPE_LABELS = {
    "A": "A、保交付 + 不资金共管",
    "B": "B、不保交付 + 不资金共管",
    "C": "C、保交付 + 资金共管",
    "D": "D、不保交付 + 资金共管",
}
CHOICES = {
    "perspective": [("buyer", "我方是购买方", "向上游购买货物"), ("seller", "我方是销售方", "向下游销售货物"), ("other", "其他合同立场", "本次不直接使用买卖条款库")],
    "stage": [("unsigned", "签署前", "尚未签署"), ("signed", "已签署，尚未履行", "修改只能作为拟变更建议"), ("performing", "履行中或已有争议", "结合实际履行审查"), ("unknown", "尚不确定", "保留状态待核实，不推定未签署")],
    "action": [("review", "只给审查意见", "不改合同正文或批注"), ("comment", "只加批注", "另存副本，正文不改"), ("revise", "按已确认范围修订", "未确认的商业条件仍须另行确认")],
    "delivery": [("tracked", "修订批注版", "保留本轮真实修订和批注"), ("clean", "清洁版", "另存无痕成稿，历史修订取舍另行确认"), ("both", "两份都要", "同一底稿派生修订版和清洁版")],
}
TITLES = {"perspective": "本次代表哪一方？", "stage": "合同目前处于哪个阶段？", "action": "本次如何处理合同？", "delivery": "修改后的 Word 需要哪种版本？", "buyer_scope": "本次纳入哪组保交付、资金共管条款？（仅选择条款范围）", "categories": "除此之外，本次还需纳入哪些条款类别？（可多选）"}
HEADERS = {"perspective": "我方立场", "stage": "合同阶段", "action": "处理方式", "delivery": "交付版本", "buyer_scope": "保交与共管", "categories": "条款类别"}


def catalog_for(answers):
    return load_catalog(SELLER_CATALOG if answers.get("perspective") == "seller" else DEFAULT_CATALOG)


def now():
    return datetime.now(timezone.utc).isoformat()


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + "." + secrets.token_hex(4) + ".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def fingerprint(sources):
    if not sources:
        raise ValueError("At least one current source file is required")
    result = []
    for source in sources:
        path = Path(source).resolve(strict=True)
        if not path.is_file():
            raise ValueError("Source must be a file")
        result.append({"path": str(path), "sha256": digest(path)})
    if len({x["path"] for x in result}) != len(result):
        raise ValueError("Duplicate source file")
    return result


def validate_values(values, catalog):
    allowed = set(CHOICES) | {"categories", "author"}
    if not isinstance(values, dict) or set(values) - allowed:
        raise ValueError("Unknown intake field; commercial parameters do not belong in this form")
    for key, value in values.items():
        if key in CHOICES and value not in {x[0] for x in CHOICES[key]}:
            raise ValueError("Invalid value for " + key)
        if key == "categories":
            if not isinstance(value, list) or any(not isinstance(x, str) for x in value):
                raise ValueError("Categories must be an explicit array")
            if len(value) != len(set(value)) or set(value) - set(catalog["categories"]):
                raise ValueError("Unknown or duplicate category")
        if key == "author" and (not isinstance(value, str) or not value.strip() or len(value) > 80):
            raise ValueError("Invalid review author")
    if "categories" in values and values.get("perspective") not in (None, "buyer", "seller"):
        raise ValueError("Trade categories cannot be assigned to another perspective")


def load_state(path, check_sources=True):
    state = read_json(path)
    if not isinstance(state, dict):
        raise ValueError("Intake receipt must be a JSON object")
    if state.get("schema_version") != SCHEMA_VERSION:
        raise ValueError("Unsupported intake receipt schema")
    catalog, sha = catalog_for(state["answers"])
    if sha != state["catalog_sha256"]:
        raise ValueError("Clause catalog changed; create a new intake receipt")
    if check_sources and fingerprint([x["path"] for x in state["sources"]]) != state["sources"]:
        raise ValueError("Source changed; do not reuse this intake receipt")
    validate_values(state["answers"], catalog)
    return state, catalog


def init_state(path, sources, known=None, evidence=""):
    known = known or {}
    catalog, sha = catalog_for(known)
    validate_values(known, catalog)
    if known and not evidence.strip():
        raise ValueError("Known answers require the user's actual instruction as evidence")
    source_set = fingerprint(sources)
    if Path(path).exists():
        state, _ = load_state(path)
        if state["sources"] != source_set:
            raise ValueError("Existing receipt belongs to different source files")
        if any(state["answers"].get(k) != v for k, v in known.items()):
            if "perspective" in known and known["perspective"] != state["answers"].get("perspective"):
                state["answers"].pop("categories", None)
            if "action" in known and known["action"] != "revise":
                state["answers"].pop("delivery", None)
            state["answers"].update(known)
            catalog, sha = catalog_for(state["answers"])
            validate_values(state["answers"], catalog)
            state.update(catalog_id=catalog["catalog_id"], catalog_sha256=sha)
            state["revision"] += 1
            state["active_plan"] = None
            state["cancelled"] = False
            state["events"].append({"at": now(), "channel": "user_text", "answers": known, "evidence": evidence})
    else:
        state = {"schema_version": SCHEMA_VERSION, "task_id": secrets.token_hex(12), "revision": 0,
                 "sources": source_set, "catalog_id": catalog["catalog_id"], "catalog_sha256": sha,
                 "answers": known, "events": [], "active_plan": None, "cancelled": False}
        if known:
            state["events"].append({"at": now(), "channel": "user_text", "answers": known, "evidence": evidence})
    state["status"] = status(state)
    write_json(path, state)
    return state


def missing(state):
    answers = state["answers"]
    result = [k for k in ("perspective", "stage", "action") if k not in answers]
    if answers.get("action") == "revise" and "delivery" not in answers:
        result.append("delivery")
    if answers.get("perspective") in ("buyer", "seller") and "categories" not in answers:
        result.append("categories")
    return result


def status(state):
    if state.get("cancelled"):
        return "cancelled"
    return "pending" if missing(state) else "submitted"


def public_state(state):
    value = {k: state[k] for k in ("schema_version", "task_id", "revision", "sources", "catalog_id", "catalog_sha256", "answers")}
    value.update(status=status(state), missing=missing(state), commercial_conditions_confirmed=False)
    if state["answers"].get("action") in ("comment", "revise"):
        value["review_author"] = state["answers"].get("author", "SR")
    return value


def make_plan(path, host="browser", phase="auto"):
    state, catalog = load_state(path)
    pending = missing(state)
    if phase == "intake":
        pending = [x for x in pending if x != "categories"]
    elif phase == "categories":
        if state["answers"].get("perspective") not in ("buyer", "seller"):
            raise ValueError("Confirm the buyer or seller perspective before choosing clause categories")
        pending = [x for x in pending if x == "categories"]
    # Establish the task before asking buyer-only questions; no speculative branch.
    if phase == "auto" and any(x != "categories" for x in pending):
        pending = [x for x in pending if x != "categories"]
    questions = []
    for key in pending:
        if key == "categories":
            if state["answers"].get("perspective") == "seller":
                options = [{"value": x["name"], "label": x["name"], "description": f'{x["module_count"]}项条款'} for x in list_categories(catalog) if x["module_count"]]
                options.append({"value": NONE, "label": "本次不用乙方预设条款", "description": "明确不调用此库；仍检查合同自身问题"})
                questions.append({"id": key, "type": "multi_select", "question": "本次采用哪些乙方预设条款类别？（可多选）", "options": options})
                continue
            grouped = set(BUYER_SCOPE["C"])
            if not grouped.issubset(catalog["categories"]):
                raise ValueError("Buyer scope categories changed; update the question layout explicitly")
            questions.append({"id": "buyer_scope", "type": "select", "question": TITLES["buyer_scope"],
                              "options": [{"value": value, "label": BUYER_SCOPE_LABELS[value],
                                           "description": "仅决定是否纳入相应预设条款，不确认具体商业条件"}
                                          for value in BUYER_SCOPE]})
            options = [{"value": x["name"], "label": x["name"], "description": f'{x["module_count"]}项整合条款 / {x["source_count"]}项来源'} for x in list_categories(catalog) if x["name"] not in grouped]
            options.append({"value": NONE, "label": "其余类别均不选", "description": "保留前题选择，不追加本题其他类别"})
        else:
            options = [{"value": v, "label": label, "description": description} for v, label, description in CHOICES[key]]
        questions.append({"id": key, "type": "multi_select" if key == "categories" else "select", "question": TITLES[key], "options": options})
    if host == "claude" and (len(questions) > 4 or any(not 2 <= len(q["options"]) <= 4 for q in questions)):
        raise ValueError("Claude native question limit exceeded; revise the layout instead of silently opening a browser")
    plan = {"request_id": secrets.token_hex(12), "task_id": state["task_id"], "revision": state["revision"],
            "host": host, "phase": phase, "questions": questions,
            "sources": state["sources"], "known": state["answers"], "catalog_sha256": state["catalog_sha256"]}
    if host == "pideck":
        plan["tool"] = "ask_question"
        plan["arguments"] = {"questions": questions, "review": True}
    elif host == "claude":
        plan["tool"] = "AskUserQuestion"
        plan["arguments"] = {"questions": [{"question": q["question"], "header": HEADERS[q["id"]], "multiSelect": q["type"] == "multi_select",
                                                   "options": [{"label": o["label"], "description": o["description"]} for o in q["options"]]} for q in questions]}
    else:
        plan["tool"] = "localhost_form"
    state["active_plan"] = plan
    write_json(path, state)
    return plan


def accept_response(path, plan, response, channel, tool_call_id=None):
    state, catalog = load_state(path)
    if not isinstance(plan, dict) or not isinstance(response, dict):
        raise ValueError("Question plan and response must be JSON objects")
    active = state.get("active_plan")
    if not active or active != plan or state["revision"] != plan["revision"]:
        raise ValueError("Stale or already-consumed question plan")
    if plan["task_id"] != state["task_id"] or plan["catalog_sha256"] != state["catalog_sha256"]:
        raise ValueError("Question plan identity mismatch")
    if channel not in ("pideck", "claude", "browser") or channel != plan["host"]:
        raise ValueError("Response channel does not match the planned interface")
    raw = response.get("details", response)
    if not isinstance(raw, dict):
        raise ValueError("Interface result must be a JSON object")
    if raw.get("cancelled") is True:
        state["cancelled"] = True
        state["active_plan"] = None
        state["events"].append({"at": now(), "channel": channel, "request_id": plan["request_id"], "status": "cancelled"})
        state["status"] = status(state)
        write_json(path, state)
        return public_state(state)
    annotations = raw.get("annotations") or {}
    if not isinstance(annotations, dict):
        raise ValueError("Invalid interface annotations")
    if (any(k in raw for k in ("afkTimeoutMs", "followUp"))
            or isinstance(raw.get("response"), str) and raw["response"].strip()
            or any(isinstance(v, dict) and v.get("unanswered") for v in annotations.values())):
        return {**public_state(state), "status": "pending", "reason": "Interface did not supply an explicit complete submission"}
    if channel in ("pideck", "browser"):
        if raw.get("cancelled") is not False or not isinstance(raw.get("answers"), list):
            raise ValueError("Missing explicit submission result")
        rows = raw["answers"]
        if len({row.get("id") for row in rows}) != len(rows):
            raise ValueError("Duplicate answer ID")
        keyed = {row.get("id"): row for row in rows}
    else:
        answers = raw.get("answers")
        if not isinstance(answers, dict):
            raise ValueError("Missing Claude answer map")
        if set(answers) != {q["question"] for q in plan["questions"]}:
            raise ValueError("Submission must answer exactly the displayed questions")
        keyed = {}
        for question in plan["questions"]:
            native = answers[question["question"]]
            # Native Claude returns multi-select labels as a list or comma-separated text.
            labels = native if isinstance(native, list) else ([x.strip() for x in native.split(",")] if isinstance(native, str) else [])
            label_map = {o["label"]: o["value"] for o in question["options"]}
            if (not labels or any(not isinstance(x, str) or x not in label_map for x in labels)
                    or len(labels) != len(set(labels)) or question["type"] == "select" and len(labels) != 1):
                raise ValueError("Missing or ambiguous native answer")
            values = [label_map[label] for label in labels]
            keyed[question["id"]] = {"id": question["id"], "type": question["type"],
                                      "value": values if question["type"] == "multi_select" else values[0]}
    if set(keyed) != {q["id"] for q in plan["questions"]}:
        raise ValueError("Submission must answer exactly the displayed questions")
    selected = {}
    for question in plan["questions"]:
        row = keyed[question["id"]]
        if row.get("type") != question["type"]:
            raise ValueError("Answer type mismatch")
        value = row.get("value")
        allowed = {x["value"] for x in question["options"]}
        if question["type"] == "multi_select":
            if not isinstance(value, list) or not value or any(not isinstance(x, str) for x in value) or len(value) != len(set(value)) or set(value) - allowed:
                raise ValueError("Choose categories or explicitly select the displayed opt-out option")
            if NONE in value and value != [NONE]:
                raise ValueError("Opt-out cannot be combined with categories")
            value = [] if value == [NONE] else [x for x in catalog["categories"] if x in value]
        elif not isinstance(value, str) or value not in allowed:
            raise ValueError("Answer is not a displayed option")
        selected[question["id"]] = value
    category_selection = None
    if "buyer_scope" in selected:
        scope = selected.pop("buyer_scope")
        category_selection = {"buyer_scope": scope, "additional_categories": selected["categories"][:]}
        included = set(BUYER_SCOPE[scope]) | set(selected["categories"])
        selected["categories"] = [x for x in catalog["categories"] if x in included]
    merged = {**state["answers"], **selected}
    discarded = []
    old_side = state["answers"].get("perspective") or "buyer"
    if "perspective" in selected and selected["perspective"] != old_side and "categories" in merged:
        merged.pop("categories")
        discarded.append("categories")
    catalog, sha = catalog_for(merged)
    validate_values(merged, catalog)
    state["answers"] = merged
    state.update(catalog_id=catalog["catalog_id"], catalog_sha256=sha)
    state["revision"] += 1
    state["active_plan"] = None
    state["cancelled"] = False
    state["events"].append({"at": now(), "channel": channel, "request_id": plan["request_id"], "tool_call_id": tool_call_id,
                            "answers": selected, "category_selection": category_selection, "discarded_inapplicable_answers": discarded,
                            "raw_response_sha256": hashlib.sha256(json.dumps(response, ensure_ascii=False, sort_keys=True).encode()).hexdigest()})
    state["status"] = status(state)
    write_json(path, state)
    return public_state(state)


def gate(path, operation="review"):
    state, _ = load_state(path)
    if status(state) != "submitted":
        raise ValueError("Intake is incomplete or cancelled; continue only source inventory and read-only checks")
    permission = {"review": {"review"}, "comment": {"review", "comment"}, "revise": {"review", "comment", "revise"}}
    if operation not in permission[state["answers"]["action"]]:
        raise ValueError("Operation exceeds the user's selected scope")
    value = public_state(state)
    value["operation"] = operation
    value["signed_document_changes_are_proposals"] = state["answers"]["stage"] != "unsigned"
    return value


def collect(path, phase="auto", timeout=900, open_browser=True):
    state, _ = load_state(path)
    if not missing(state):
        return public_state(state)
    token = secrets.token_urlsafe(24)
    shared = {"plan": make_plan(path, "browser", phase), "done": False, "result": None}
    if not shared["plan"]["questions"]:
        return {**public_state(state), "reason": "No unanswered questions in the requested phase"}
    template = (SCRIPT_ROOT.parent / "assets" / "intake-form.html").read_text(encoding="utf-8")

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass

        def send(self, code, body, content_type="application/json; charset=utf-8"):
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(body.encode("utf-8"))

        def authorized(self):
            return parse_qs(urlsplit(self.path).query).get("token") == [token]

        def do_GET(self):
            if not self.authorized() or urlsplit(self.path).path != "/":
                self.send(404, "{}")
                return
            payload = json.dumps(shared["plan"], ensure_ascii=False).replace("<", "\\u003c")
            self.send(200, template.replace("__HETONG_PLAN__", payload), "text/html; charset=utf-8")

        def do_POST(self):
            if not self.authorized() or urlsplit(self.path).path != "/submit":
                self.send(404, "{}")
                return
            origin = self.headers.get("Origin")
            if origin and origin != f"http://127.0.0.1:{self.server.server_port}":
                self.send(403, "{}")
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if not 0 < length <= 65536:
                    raise ValueError("Invalid request size")
                raw = json.loads(self.rfile.read(length))
                if not isinstance(raw, dict):
                    raise ValueError("Submission must be a JSON object")
                if raw.get("request_id") != shared["plan"]["request_id"]:
                    raise ValueError("Stale browser form")
                result = accept_response(path, shared["plan"], raw, "browser")
                if result["status"] in ("submitted", "cancelled"):
                    shared.update(done=True, result=result)
                    self.send(200, json.dumps(result, ensure_ascii=False))
                else:
                    next_plan = make_plan(path, "browser", phase)
                    if not next_plan["questions"]:
                        shared.update(done=True, result=result)
                        self.send(200, json.dumps(result, ensure_ascii=False))
                    else:
                        shared["plan"] = next_plan
                        self.send(200, json.dumps({"status": "next", "plan": next_plan}, ensure_ascii=False))
            except (ValueError, KeyError, TypeError, OSError) as exc:
                self.send(422, json.dumps({"error": str(exc)}, ensure_ascii=False))

    server = HTTPServer(("127.0.0.1", 0), Handler)
    server.timeout = 0.5
    url = f"http://127.0.0.1:{server.server_port}/?token={token}"
    print(json.dumps({"event": "form_opened", "url": url, "state_path": str(Path(path).resolve())}, ensure_ascii=False), flush=True)
    if open_browser:
        try:
            webbrowser.open(url, new=2)
        except webbrowser.Error:
            print("Open the form URL shown above; automatic browser launch was unavailable.", file=sys.stderr)
    deadline = time.monotonic() + timeout
    try:
        while not shared["done"] and time.monotonic() < deadline:
            server.handle_request()
    finally:
        server.server_close()
    return shared["result"] or {**public_state(load_state(path)[0]), "status": "pending", "reason": "No explicit submission before timeout; no default answer was applied"}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    init = sub.add_parser("init")
    init.add_argument("state", type=Path)
    init.add_argument("--source", action="append", required=True)
    init.add_argument("--known", type=Path)
    init.add_argument("--evidence", default="")
    for name in ("status", "plan", "accept", "collect", "gate", "export"):
        item = sub.add_parser(name)
        item.add_argument("state", type=Path)
        if name in ("plan", "collect"):
            item.add_argument("--phase", choices=("auto", "intake", "categories"), default="auto")
        if name == "plan":
            item.add_argument("--host", choices=("pideck", "claude", "browser"), default="browser")
            item.add_argument("--out", type=Path)
        if name == "accept":
            item.add_argument("--plan", type=Path, required=True)
            item.add_argument("--response", type=Path, required=True)
            item.add_argument("--host", choices=("pideck", "claude", "browser"), required=True)
            item.add_argument("--tool-call-id")
        if name == "collect":
            item.add_argument("--timeout", type=float, default=900)
            item.add_argument("--no-open", action="store_true")
        if name == "gate":
            item.add_argument("--operation", choices=("review", "comment", "revise"), default="review")
        if name == "export":
            item.add_argument("--out", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "init":
            result = public_state(init_state(args.state, args.source, read_json(args.known) if args.known else {}, args.evidence))
        elif args.command == "plan":
            result = make_plan(args.state, args.host, args.phase)
            if args.out:
                write_json(args.out, result)
        elif args.command == "accept":
            result = accept_response(args.state, read_json(args.plan), read_json(args.response), args.host, args.tool_call_id)
        elif args.command == "collect":
            result = collect(args.state, args.phase, args.timeout, not args.no_open)
        elif args.command == "gate":
            result = gate(args.state, args.operation)
        elif args.command == "export":
            result = gate(args.state)
            state, catalog = load_state(args.state)
            if state["answers"].get("perspective") not in ("buyer", "seller"):
                raise ValueError("No trade preset selection for this perspective")
            modules, selection = select_categories(catalog, state["answers"]["categories"], state["catalog_sha256"])
            if args.out.exists():
                raise ValueError("Do not overwrite an existing clause export")
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(render_modules(modules, selection), encoding="utf-8")
            result["selected_clauses_path"] = str(args.out.resolve())
        else:
            result = public_state(load_state(args.state)[0])
        print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
        return 3 if result.get("status") in ("pending", "cancelled") and args.command == "collect" else 0
    except (ValueError, KeyError, TypeError, OSError) as exc:
        print("hetong_intake: " + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
