#!/usr/bin/env python3
"""Offline behavior tests for intake receipts; every fixture is fictional.

Run with: python3 -B -m unittest -v test_hetong_intake.py
The HTTP tests bind loopback only and never open a browser or call a model.
"""
from __future__ import annotations

from contextlib import contextmanager
import copy
import importlib.util
import itertools
import json
from pathlib import Path
import queue
import subprocess
import sys
import tempfile
import threading
import unittest
from unittest import mock
from urllib.error import HTTPError
from urllib.parse import urlsplit, urlunsplit
from urllib.request import ProxyHandler, Request, build_opener


SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
SPEC = importlib.util.spec_from_file_location("hetong_intake_under_test", SCRIPTS / "hetong_intake.py")
core = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(core)


class IntakeBehaviorTests(unittest.TestCase):
    def test_malformed_interface_result_cannot_change_receipt(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state, "pideck")
        before = self.snapshot()
        for invalid in ([], None, {"details": []}):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                core.accept_response(self.state, plan, invalid, "pideck")
            self.assertEqual(self.snapshot(), before)

    def test_core_starts_under_python_isolated_mode(self):
        result = subprocess.run([sys.executable, "-I", str(SCRIPTS / "hetong_intake.py"), "--help"], cwd=self.root, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("collect", result.stdout)

    def test_new_seller_answer_discards_inapplicable_unbound_buyer_categories(self):
        self.init({"categories": [self.categories[0]], "stage": "unsigned", "action": "review"})
        plan = core.make_plan(self.state, "browser")
        result = core.accept_response(self.state, plan, self.response(plan, {"perspective": "seller"}), "browser")
        self.assertEqual(result["status"], "pending")
        self.assertEqual(result["missing"], ["categories"])
        self.assertNotIn("categories", result["answers"])
        self.assertEqual(core.read_json(self.state)["events"][-1]["discarded_inapplicable_answers"], ["categories"])

    def test_empty_phase_never_opens_an_unanswerable_form(self):
        self.init({"perspective": "buyer", "categories": []})
        with mock.patch.object(core.webbrowser, "open") as browser:
            result = core.collect(self.state, phase="categories", timeout=0.1)
        self.assertEqual(result["status"], "pending")
        self.assertEqual(result["missing"], ["stage", "action"])
        browser.assert_not_called()

    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="hetong-intake-test-")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "fictional-source.txt"
        self.source.write_text("虚构测试材料，不包含真实合同。\n", encoding="utf-8")
        self.catalog_path = self.root / "fictional-catalog.json"
        self.categories = ["保交条款", "必要条款", "一般条款", "其他条款", "共管条款"]
        catalog = {
            "schema_version": 1,
            "catalog_id": "fictional-offline-catalog",
            "categories": self.categories,
            "sources": [
                {"id": f"T{i}", "category": category}
                for i, category in enumerate(self.categories, 1)
            ],
            "modules": [
                {
                    "id": f"M{i}", "title": "虚构测试条款",
                    "category": category, "source_ids": [f"T{i}"],
                    "essential": "仅供行为测试。", "example": "不用于任何真实交易。",
                    "action": "clause_example",
                }
                for i, category in enumerate(self.categories, 1)
            ],
            "relations": [],
        }
        self.catalog_path.write_text(json.dumps(catalog, ensure_ascii=False), encoding="utf-8")
        seller_catalog = copy.deepcopy(catalog)
        seller_catalog["library_side"] = "seller"
        seller_catalog["catalog_id"] = "fictional-seller-catalog"
        seller_catalog["sources"] = [x for x in seller_catalog["sources"] if x["category"] in ("必要条款", "其他条款")]
        seller_catalog["modules"] = [x for x in seller_catalog["modules"] if x["category"] in ("必要条款", "其他条款")]
        self.seller_catalog_path = self.root / "fictional-seller-catalog.json"
        self.seller_catalog_path.write_text(json.dumps(seller_catalog, ensure_ascii=False), encoding="utf-8")
        seller_patcher = mock.patch.object(core, "SELLER_CATALOG", self.seller_catalog_path)
        seller_patcher.start()
        self.addCleanup(seller_patcher.stop)
        patcher = mock.patch.object(core, "DEFAULT_CATALOG", self.catalog_path)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.state = self.root / "state.json"

    def init(self, known=None, path=None):
        return core.init_state(
            path or self.state, [self.source], known or {},
            "测试用户明确给出的已知范围。" if known else "",
        )

    def buyer_categories_pending(self):
        return self.init({"perspective": "buyer", "stage": "unsigned", "action": "review"})

    def response(self, plan, overrides=None, cancelled=False):
        overrides = overrides or {}
        return {
            "request_id": plan["request_id"], "cancelled": cancelled,
            "answers": [] if cancelled else [
                {
                    "id": q["id"], "type": q["type"],
                    "value": overrides.get(
                        q["id"], [q["options"][0]["value"]]
                        if q["type"] == "multi_select" else "B" if q["id"] == "buyer_scope" else q["options"][0]["value"],
                    ),
                }
                for q in plan["questions"]
            ],
        }

    def native_response(self, plan):
        return {"answers": {q["question"]: q["options"][0]["label"] for q in plan["questions"]}}

    def snapshot(self):
        return self.state.read_bytes()

    @staticmethod
    def http(url, body=None, path=None, origin=None):
        parts = urlsplit(url)
        target = urlunsplit((parts.scheme, parts.netloc, path or parts.path, parts.query, ""))
        headers = {"Content-Type": "application/json"}
        if origin is not None:
            headers["Origin"] = origin
        data = None if body is None else json.dumps(body).encode("utf-8")
        request = Request(target, data=data, headers=headers, method="GET" if data is None else "POST")
        # Ignore ambient HTTP proxy settings: these tests must stay on loopback.
        with build_opener(ProxyHandler({})).open(request, timeout=3) as result:
            return result.status, result.headers, result.read().decode("utf-8")

    @contextmanager
    def collector(self, phase="categories"):
        opened = queue.Queue()
        finished = queue.Queue()

        def capture(message, **_):
            event = json.loads(message)
            if event.get("event") == "form_opened":
                opened.put(event)

        def run():
            try:
                finished.put((True, core.collect(self.state, phase=phase, timeout=4, open_browser=False)))
            except BaseException as exc:
                finished.put((False, exc))

        with mock.patch.object(core, "print", capture, create=True), mock.patch.object(core.webbrowser, "open") as browser:
            thread = threading.Thread(target=run, daemon=True)
            thread.start()
            event = opened.get(timeout=3)
            try:
                yield event["url"], finished
            finally:
                if thread.is_alive():
                    current = core.read_json(self.state)
                    active = current.get("active_plan")
                    if active:
                        try:
                            self.http(event["url"], self.response(active, cancelled=True), "/submit")
                        except (HTTPError, OSError):
                            pass
                thread.join(timeout=5)
                self.assertFalse(thread.is_alive(), "localhost collector did not stop")
                browser.assert_not_called()

    def test_known_answers_are_not_asked_again(self):
        known = {"perspective": "buyer", "stage": "unsigned", "action": "revise", "delivery": "tracked", "categories": [self.categories[2]]}
        state = self.init(known)
        self.assertEqual(core.missing(state), [])
        self.assertEqual(core.make_plan(self.state)["questions"], [])
        self.assertEqual(core.public_state(state)["answers"], known)

    def test_known_values_require_actual_user_evidence(self):
        with self.assertRaisesRegex(ValueError, "evidence"):
            core.init_state(self.state, [self.source], {"perspective": "buyer"})
        self.assertFalse(self.state.exists())

    def test_buyer_and_seller_request_their_own_categories(self):
        for perspective in ("buyer", "seller", "other"):
            with self.subTest(perspective=perspective):
                state = self.init({"perspective": perspective}, self.root / f"{perspective}.json")
                self.assertEqual("categories" in core.missing(state), perspective in ("buyer", "seller"))

    def test_only_revision_branch_requests_delivery(self):
        for action in ("review", "comment", "revise"):
            with self.subTest(action=action):
                state = self.init({"perspective": "seller", "stage": "unsigned", "action": action}, self.root / f"{action}.json")
                self.assertEqual("delivery" in core.missing(state), action == "revise")

    def test_claude_native_scope_single_then_remaining_multi_select(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state, host="claude", phase="categories")
        self.assertEqual(plan["host"], "claude")
        self.assertEqual(plan["tool"], "AskUserQuestion")
        self.assertEqual([q["id"] for q in plan["questions"]], ["buyer_scope", "categories"])
        self.assertEqual([o["value"] for o in plan["questions"][0]["options"]], ["A", "B", "C", "D"])
        self.assertEqual([o["value"] for o in plan["questions"][1]["options"]], self.categories[1:4] + [core.NONE])
        self.assertEqual([q["multiSelect"] for q in plan["arguments"]["questions"]], [False, True])

    def test_pideck_uses_real_multi_select_and_review_submission(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state, host="pideck", phase="categories")
        self.assertEqual(plan["tool"], "ask_question")
        self.assertTrue(plan["arguments"]["review"])
        self.assertEqual([q["type"] for q in plan["arguments"]["questions"]], ["select", "multi_select"])

    def test_intake_submission_does_not_auto_select_buyer_categories(self):
        self.init()
        plan = core.make_plan(self.state, host="browser", phase="intake")
        result = core.accept_response(self.state, plan, self.response(plan), "browser")
        self.assertEqual(result["status"], "pending")
        self.assertEqual(result["missing"], ["categories"])
        self.assertNotIn("categories", result["answers"])

    def test_cancel_is_not_an_empty_selection_or_permission(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        result = core.accept_response(self.state, plan, self.response(plan, cancelled=True), "browser")
        self.assertEqual(result["status"], "cancelled")
        self.assertNotIn("categories", result["answers"])
        with self.assertRaisesRegex(ValueError, "incomplete or cancelled"):
            core.gate(self.state)

    def test_empty_checkbox_array_is_not_explicit_opt_out(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, "explicitly select"):
            core.accept_response(self.state, plan, self.response(plan, {"categories": []}), "browser")
        self.assertEqual(self.snapshot(), before)

    def test_none_sentinel_explicitly_submits_empty_scope(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        result = core.accept_response(self.state, plan, self.response(plan, {"categories": [core.NONE]}), "browser")
        self.assertEqual(result["status"], "submitted")
        self.assertEqual(result["answers"]["categories"], [])
        self.assertFalse(result["commercial_conditions_confirmed"])

    def test_multiple_categories_are_exact_not_expanded(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state, host="pideck")
        result = core.accept_response(self.state, plan, {"details": self.response(plan, {"buyer_scope": "C", "categories": [core.NONE]})}, "pideck", "fake-tool-1")
        self.assertEqual(result["answers"]["categories"], [self.categories[0], self.categories[-1]])
        self.assertEqual(core.read_json(self.state)["events"][-1]["tool_call_id"], "fake-tool-1")

    def test_none_cannot_be_mixed_with_real_categories(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        with self.assertRaisesRegex(ValueError, "cannot be combined"):
            core.accept_response(self.state, plan, self.response(plan, {"categories": [core.NONE, self.categories[1]]}), "browser")

    def test_ui_submission_requires_flag_exact_ids_and_unique_rows(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        valid = self.response(plan)
        bad_responses = [
            {"answers": valid["answers"]},
            {**valid, "answers": []},
            {**valid, "answers": valid["answers"] * 2},
            {**valid, "answers": valid["answers"] + [{"id": "amount", "type": "select", "value": "1"}]},
        ]
        before = self.snapshot()
        for response in bad_responses:
            with self.subTest(response=response), self.assertRaises(ValueError):
                core.accept_response(self.state, plan, response, "browser")
            self.assertEqual(self.snapshot(), before)

    def test_commercial_fields_are_refused_in_known_answers(self):
        for field in ("amount", "payment_party", "guarantee", "payment_deadline", "advance_repayment"):
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, "commercial parameters"):
                self.init({field: "fictional-value"})
        self.assertFalse(self.state.exists())

    def test_all_32_category_combinations_round_trip_on_each_host(self):
        for host, scope in itertools.product(("claude", "pideck", "browser"), core.BUYER_SCOPE):
            for mask in range(8):
                additional = [x for i, x in enumerate(self.categories[1:4]) if mask & (1 << i)]
                with self.subTest(host=host, scope=scope, additional=additional):
                    path = self.root / f"{host}-{scope}-{mask}.json"
                    self.init({"perspective": "buyer", "stage": "unsigned", "action": "review"}, path)
                    plan = core.make_plan(path, host, "categories")
                    raw = self.response(plan, {"buyer_scope": scope, "categories": additional or [core.NONE]})
                    if host == "claude":
                        labels = {}
                        for question, row in zip(plan["questions"], raw["answers"]):
                            values = row["value"] if isinstance(row["value"], list) else [row["value"]]
                            labels[question["question"]] = [o["label"] for o in question["options"] if o["value"] in values]
                        # Cover both native return shapes, including multiple selected labels.
                        raw = {"answers": {q: (", ".join(values) if mask % 2 else values) for q, values in labels.items()}}
                    result = core.accept_response(path, plan, raw, host)
                    expected = set(core.BUYER_SCOPE[scope]) | set(additional)
                    self.assertEqual(result["status"], "submitted")
                    self.assertEqual(result["answers"]["categories"], [x for x in self.categories if x in expected])
                    self.assertFalse(result["commercial_conditions_confirmed"])
                    self.assertEqual(core.read_json(path)["events"][-1]["category_selection"], {"buyer_scope": scope, "additional_categories": additional})

    def test_front_question_or_checkbox_partial_answer_never_submits(self):
        self.buyer_categories_pending()
        for host in ("claude", "pideck", "browser"):
            plan = core.make_plan(self.state, host)
            for omitted in (0, 1):
                raw = self.response(plan)
                raw["answers"].pop(omitted)
                if host == "claude":
                    raw = {"answers": {q["question"]: q["options"][0]["label"] for i, q in enumerate(plan["questions"]) if i != omitted}}
                before = self.snapshot()
                with self.subTest(host=host, omitted=omitted), self.assertRaises(ValueError):
                    core.accept_response(self.state, plan, raw, host)
                self.assertEqual(self.snapshot(), before)

    def test_native_rejects_mixed_none_free_text_and_duplicate_labels(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state, "claude")
        front, extra = plan["questions"]
        before = self.snapshot()
        for value in ([], "", "请先讨论", [extra["options"][0]["label"]] * 2,
                      [extra["options"][0]["label"], extra["options"][-1]["label"]]):
            with self.subTest(value=value), self.assertRaises(ValueError):
                core.accept_response(self.state, plan, {"answers": {front["question"]: front["options"][0]["label"], extra["question"]: value}}, "claude")
            self.assertEqual(self.snapshot(), before)

    def test_expanded_catalog_never_silently_routes_claude_to_browser(self):
        data = json.loads(self.catalog_path.read_text())
        data["categories"].append("未来类别")
        data["sources"].append({"id": "T6", "category": "未来类别"})
        data["modules"].append({**data["modules"][0], "id": "M6", "category": "未来类别", "source_ids": ["T6"]})
        self.catalog_path.write_text(json.dumps(data, ensure_ascii=False))
        self.buyer_categories_pending()
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, "native question limit"):
            core.make_plan(self.state, "claude")
        self.assertEqual(self.snapshot(), before)

    def test_unknown_stage_and_delivery_are_resolved_before_category_batch(self):
        self.init({"perspective": "buyer", "action": "revise"})
        plan = core.make_plan(self.state, "claude")
        self.assertEqual([q["id"] for q in plan["questions"]], ["stage", "delivery"])
        result = core.accept_response(self.state, plan, self.native_response(plan), "claude")
        self.assertEqual(result["missing"], ["categories"])
        self.assertEqual([q["id"] for q in core.make_plan(self.state, "claude")["questions"]], ["buyer_scope", "categories"])

    def test_trade_categories_cannot_be_assigned_to_other_perspective(self):
        with self.assertRaisesRegex(ValueError, "another perspective"):
            self.init({"perspective": "other", "categories": [self.categories[0]]})
        self.init({"perspective": "buyer", "stage": "unsigned", "action": "review", "categories": [self.categories[0]]})
        before = self.snapshot()
        try:
            self.init({"perspective": "seller"})
        except ValueError:
            # Rejecting an incompatible correction is safe only if atomic.
            self.assertEqual(self.snapshot(), before)
        else:
            try:
                corrected, _ = core.load_state(self.state)
            except ValueError as exc:
                self.fail(f"User correction persisted an unreadable intake state: {exc}")
            self.assertEqual(corrected["answers"]["perspective"], "seller")
            self.assertNotIn("categories", corrected["answers"])

    def test_consumed_plan_cannot_be_replayed(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        response = self.response(plan)
        core.accept_response(self.state, plan, response, "browser")
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, "already-consumed"):
            core.accept_response(self.state, plan, response, "browser")
        self.assertEqual(self.snapshot(), before)

    def test_replacement_plan_and_revision_invalidate_old_request(self):
        self.buyer_categories_pending()
        first = core.make_plan(self.state)
        second = core.make_plan(self.state)
        with self.assertRaisesRegex(ValueError, "Stale"):
            core.accept_response(self.state, first, self.response(first), "browser")
        self.init({"stage": "signed"})
        with self.assertRaisesRegex(ValueError, "Stale"):
            core.accept_response(self.state, second, self.response(second), "browser")

    def test_plan_payload_cannot_relabel_question_to_overwrite_known_value(self):
        self.buyer_categories_pending()
        plan = core.make_plan(self.state)
        forged = copy.deepcopy(plan)
        forged["questions"] = [{"id": "action", "type": "select", "question": "伪造的未展示问题", "options": [{"value": "revise", "label": "伪造写入授权", "description": ""}]}]
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, "plan|Plan|identity|Identity|Stale|displayed"):
            core.accept_response(self.state, forged, self.response(forged), "browser")
        self.assertEqual(self.snapshot(), before)

    def test_source_byte_change_invalidates_receipt_and_gate(self):
        self.init({"perspective": "seller", "stage": "unsigned", "action": "review"})
        self.source.write_text("另一份虚构材料。\n", encoding="utf-8")
        for operation in (lambda: core.load_state(self.state), lambda: core.gate(self.state)):
            with self.subTest(operation=operation), self.assertRaisesRegex(ValueError, "Source changed"):
                operation()

    def test_catalog_byte_change_invalidates_existing_receipt(self):
        self.buyer_categories_pending()
        self.catalog_path.write_text(self.catalog_path.read_text(encoding="utf-8") + "\n ", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "catalog changed"):
            core.load_state(self.state)

    def test_native_timeout_followup_and_skip_do_not_supply_answers(self):
        self.init({"perspective": "seller"})
        plan = core.make_plan(self.state, host="claude")
        valid = self.native_response(plan)
        before = self.snapshot()
        for extra in ({"afkTimeoutMs": 30000}, {"followUp": True}, {"annotations": {plan["questions"][0]["question"]: {"unanswered": "skipped"}}}, {"annotations": {plan["questions"][0]["question"]: {"unanswered": "dismissed"}}}, {"response": "先不要提交这些选项，我要改一下。"}):
            with self.subTest(extra=extra):
                result = core.accept_response(self.state, plan, {**valid, **extra}, "claude")
                self.assertEqual(result["status"], "pending")
                self.assertEqual(self.snapshot(), before)

    def test_complete_native_answer_map_submits_selected_scope(self):
        self.init({"perspective": "seller", "categories": []})
        plan = core.make_plan(self.state, host="claude")
        self.assertEqual(plan["tool"], "AskUserQuestion")
        self.assertTrue(all(len(q["options"]) <= 4 for q in plan["arguments"]["questions"]))
        result = core.accept_response(self.state, plan, self.native_response(plan), "claude", "fake-native-tool")
        self.assertEqual(result["status"], "submitted")
        self.assertEqual(result["answers"], {"perspective": "seller", "stage": "unsigned", "action": "review", "categories": []})

    def test_gate_permission_matrix_and_signed_change_marker(self):
        allowed = {"review": {"review"}, "comment": {"review", "comment"}, "revise": {"review", "comment", "revise"}}
        for action in allowed:
            path = self.root / f"gate-{action}.json"
            known = {"perspective": "seller", "stage": "signed", "action": action, "categories": []}
            if action == "revise":
                known["delivery"] = "tracked"
            self.init(known, path)
            for operation in ("review", "comment", "revise"):
                with self.subTest(action=action, operation=operation):
                    if operation in allowed[action]:
                        result = core.gate(path, operation)
                        self.assertTrue(result["signed_document_changes_are_proposals"])
                        self.assertFalse(result["commercial_conditions_confirmed"])
                    else:
                        with self.assertRaisesRegex(ValueError, "exceeds"):
                            core.gate(path, operation)

    def test_seller_plan_and_submission_use_seller_library(self):
        self.init({"perspective": "seller", "stage": "unsigned", "action": "review"})
        plan = core.make_plan(self.state, "claude")
        self.assertEqual([q["id"] for q in plan["questions"]], ["categories"])
        question = plan["questions"][0]
        self.assertEqual([o["value"] for o in question["options"]], ["必要条款", "其他条款", core.NONE])
        self.assertTrue(plan["arguments"]["questions"][0]["multiSelect"])
        result = core.accept_response(self.state, plan, {"answers": {question["question"]: "必要条款, 其他条款"}}, "claude")
        self.assertEqual(result["status"], "submitted")
        self.assertEqual(result["catalog_id"], "fictional-seller-catalog")
        self.assertEqual(result["answers"]["categories"], ["必要条款", "其他条款"])

    def test_switching_to_seller_discards_buyer_selection_even_same_category_name(self):
        self.init({"perspective": "buyer", "stage": "unsigned", "action": "review", "categories": ["必要条款"]})
        result = self.init({"perspective": "seller"})
        self.assertNotIn("categories", result["answers"])
        self.assertEqual(result["catalog_id"], "fictional-seller-catalog")
        self.assertEqual(result["status"], "pending")
        self.assertEqual(core.missing(result), ["categories"])

    def test_seller_cancel_never_becomes_empty_scope(self):
        self.init({"perspective": "seller", "stage": "unsigned", "action": "comment"})
        plan = core.make_plan(self.state, "pideck")
        result = core.accept_response(self.state, plan, self.response(plan, cancelled=True), "pideck")
        self.assertEqual(result["status"], "cancelled")
        self.assertNotIn("categories", result["answers"])
        with self.assertRaisesRegex(ValueError, "incomplete or cancelled"):
            core.gate(self.state, "comment")

    def test_localhost_get_and_real_post_submit_without_browser(self):
        self.buyer_categories_pending()
        with self.collector() as (url, finished):
            self.assertEqual(urlsplit(url).hostname, "127.0.0.1")
            status, headers, html = self.http(url)
            self.assertEqual(status, 200)
            self.assertIn("no-store", headers["Cache-Control"])
            self.assertIn("'checkbox':'radio'", html)
            self.assertIn("提交选择", html)
            self.assertNotIn("__HETONG_PLAN__", html)
            self.assertNotIn("categories", core.read_json(self.state)["answers"])
            plan = core.read_json(self.state)["active_plan"]
            status, _, text = self.http(url, self.response(plan, {"buyer_scope": "C", "categories": [core.NONE]}), "/submit", f"http://{urlsplit(url).netloc}")
            self.assertEqual(status, 200)
            self.assertEqual(json.loads(text)["status"], "submitted")
            ok, result = finished.get(timeout=3)
            self.assertTrue(ok, result)
            self.assertEqual(result["answers"]["categories"], [self.categories[0], self.categories[-1]])

    def test_localhost_rejects_bad_token_origin_and_stale_form_then_cancels(self):
        self.buyer_categories_pending()
        with self.collector() as (url, finished):
            plan = core.read_json(self.state)["active_plan"]
            before = self.snapshot()
            for target, body, origin, expected in (
                (url.replace("token=", "wrong="), self.response(plan), None, 404),
                (url, self.response(plan), "http://untrusted.invalid", 403),
                (url, {**self.response(plan), "request_id": "stale"}, None, 422),
            ):
                with self.subTest(expected=expected):
                    with self.assertRaises(HTTPError) as raised:
                        self.http(target, body, "/submit", origin)
                    self.assertEqual(raised.exception.code, expected)
                    raised.exception.close()
                    self.assertEqual(self.snapshot(), before)
            self.http(url, self.response(plan, cancelled=True), "/submit")
            ok, result = finished.get(timeout=3)
            self.assertTrue(ok, result)
            self.assertEqual(result["status"], "cancelled")
            self.assertNotIn("categories", result["answers"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
