"""Vendor the built-in Material Symbols icon set into the package
(build-time only — the running server never touches the network).

Downloads a curated business-deck subset of Google Material Symbols
(Apache 2.0) from the @material-symbols/svg-400 npm package via jsDelivr and
writes src/ppt_mcp/assets/icons/material/{icons/*.svg, index.json, LICENSE}.

Run once per icon-list change: uv run python scripts/fetch_icon_set.py
"""

import json
import time
import urllib.request
from pathlib import Path

DEST = Path(__file__).parent.parent / "src" / "ppt_mcp" / "assets" / "icons" / "material"
URL = "https://cdn.jsdelivr.net/npm/@material-symbols/svg-400/outlined/{name}.svg"

# icon name -> extra search tags (name parts are always searchable)
ICONS: dict[str, list[str]] = {
    # status & signals
    "check_circle": ["done", "success", "ok", "complete"],
    "cancel": ["fail", "rejected", "no"],
    "warning": ["risk", "caution", "alert"],
    "error": ["problem", "issue", "critical"],
    "info": ["information", "note"],
    "help": ["question", "support", "faq"],
    "pending": ["waiting", "in progress"],
    "block": ["blocked", "forbidden", "stop"],
    "priority_high": ["urgent", "important", "exclamation"],
    "verified": ["approved", "certified", "quality"],
    "release_alert": ["new", "launch", "release"],
    "flag": ["milestone", "marker", "goal"],
    # people & org
    "person": ["user", "individual", "employee"],
    "group": ["team", "people", "pair"],
    "groups": ["organization", "community", "audience", "workforce"],
    "diversity_3": ["diversity", "inclusion", "collaboration circle"],
    "handshake": ["partnership", "agreement", "deal", "cooperation"],
    "support_agent": ["service", "helpdesk", "call center"],
    "badge": ["id", "identity", "employee card"],
    "person_add": ["hire", "recruit", "onboarding"],
    "engineering": ["engineer", "technical", "construction worker"],
    "psychology": ["mind", "brain", "behavior", "ai thinking"],
    "school": ["education", "training", "learning", "academy"],
    "workspace_premium": ["certificate", "premium", "award ribbon"],
    "trophy": ["trophy", "winner", "achievement", "award"],
    "military_tech": ["medal", "honor", "rank"],
    # business & finance
    "business_center": ["briefcase", "business", "portfolio"],
    "work": ["job", "career", "briefcase"],
    "account_balance": ["bank", "government", "institution", "finance"],
    "payments": ["payment", "money", "transaction"],
    "savings": ["piggy bank", "save", "fund"],
    "credit_card": ["card", "payment", "billing"],
    "attach_money": ["dollar", "cost", "price", "revenue"],
    "euro": ["currency", "cost", "price euro"],
    "shopping_cart": ["purchase", "procurement", "buy", "ecommerce"],
    "inventory_2": ["inventory", "stock", "warehouse", "box"],
    "local_shipping": ["logistics", "delivery", "truck", "transport"],
    "factory": ["manufacturing", "plant", "industry", "production"],
    "storefront": ["shop", "retail", "store", "market"],
    "gavel": ["legal", "law", "compliance", "court", "decision"],
    "balance": ["justice", "scales", "fairness", "tradeoff", "governance"],
    "receipt_long": ["invoice", "receipt", "billing", "audit trail"],
    "request_quote": ["quote", "offer", "estimate"],
    # charts & analysis
    "bar_chart": ["chart", "graph", "statistics", "report"],
    "pie_chart": ["share", "distribution", "segments"],
    "show_chart": ["line chart", "trend", "performance"],
    "trending_up": ["growth", "increase", "improvement"],
    "trending_down": ["decline", "decrease", "loss"],
    "trending_flat": ["stable", "flat", "unchanged"],
    "analytics": ["analysis", "data", "metrics"],
    "chart_data": ["insight", "spark", "discovery"],
    "query_stats": ["investigate", "analysis", "statistics search"],
    "monitoring": ["observability", "dashboard", "tracking"],
    "leaderboard": ["ranking", "comparison", "podium"],
    "dashboard": ["overview", "panel", "kpi"],
    "table_chart": ["table", "spreadsheet", "matrix"],
    "timeline": ["roadmap", "history", "sequence"],
    "schema": ["structure", "diagram", "hierarchy"],
    "account_tree": ["org chart", "tree", "structure", "dependencies"],
    "hub": ["network", "center", "ecosystem", "platform"],
    "speed": ["fast", "performance", "velocity", "gauge"],
    "target": ["goal", "objective", "aim", "okr"],
    # process & workflow
    "task_alt": ["task", "todo", "done", "checkbox"],
    "checklist": ["list", "requirements", "criteria"],
    "fact_check": ["verify", "review", "validation", "audit"],
    "assignment": ["document task", "clipboard", "assignment"],
    "approval": ["stamp", "approve", "sign off"],
    "published_with_changes": ["change", "deployed", "updated cycle"],
    "autorenew": ["cycle", "recurring", "iteration", "agile"],
    "sync": ["synchronize", "integration", "exchange"],
    "swap_horiz": ["swap", "exchange", "migration"],
    "compare_arrows": ["compare", "versus", "benchmark"],
    "alt_route": ["alternative", "detour", "option", "scenario"],
    "call_split": ["split", "branch", "fork", "divide"],
    "merge": ["merge", "consolidate", "combine"],
    "route": ["path", "journey", "roadmap route"],
    "rocket_launch": ["launch", "startup", "go live", "kickoff"],
    "flag_circle": ["finish", "goal flag", "milestone circle"],
    "hourglass_top": ["waiting", "duration", "time running"],
    "schedule": ["clock", "time", "deadline"],
    "calendar_month": ["calendar", "date", "planning", "schedule"],
    "event": ["appointment", "meeting", "date"],
    "history": ["past", "log", "previous", "undo"],
    "update": ["refresh", "new version", "upgrade"],
    "timer": ["stopwatch", "countdown", "sprint"],
    # technology & data
    "cloud": ["cloud computing", "saas", "hosting"],
    "cloud_upload": ["upload", "migration to cloud"],
    "cloud_done": ["cloud ready", "migrated"],
    "storage": ["database", "disk", "server", "data store"],
    "database": ["data", "sql", "repository"],
    "dns": ["server", "infrastructure", "hosting"],
    "lan": ["network", "topology", "intranet"],
    "router": ["network device", "gateway", "connectivity"],
    "memory": ["chip", "processor", "hardware", "compute", "ai chip"],
    "computer": ["desktop", "workstation", "pc"],
    "devices": ["multi device", "responsive", "hardware"],
    "smartphone": ["mobile", "phone", "app"],
    "terminal": ["console", "cli", "shell", "developer"],
    "code": ["development", "programming", "software"],
    "api": ["interface", "integration", "endpoint"],
    "webhook": ["automation", "callback", "integration"],
    "bug_report": ["bug", "defect", "issue", "testing"],
    "build": ["wrench", "tool", "configuration", "maintenance"],
    "settings": ["gear", "configuration", "preferences", "setup"],
    "smart_toy": ["robot", "bot", "ai assistant", "automation"],
    "precision_manufacturing": ["robot arm", "automation", "industry 4.0"],
    "auto_awesome": ["ai", "magic", "sparkle", "genai", "generative"],
    "network_intelligence": ["ai network", "machine learning", "neural"],
    # security & governance
    "security": ["shield", "protection", "defense"],
    "shield": ["security", "guard", "safety"],
    "lock": ["secure", "private", "confidential", "encryption"],
    "lock_open": ["unlocked", "access granted", "open"],
    "key": ["access", "credential", "password", "license"],
    "vpn_key": ["authentication", "secret", "token"],
    "policy": ["governance", "rules", "compliance shield"],
    "admin_panel_settings": ["admin", "access control", "permissions"],
    "visibility": ["eye", "transparency", "watch", "observe"],
    "visibility_off": ["hidden", "private", "blind spot"],
    "fingerprint": ["biometric", "identity", "unique"],
    "health_and_safety": ["safety", "health", "care", "protection"],
    "medical_services": ["medical", "healthcare", "first aid"],
    "verified_user": ["protected", "secure check", "trusted"],
    # communication & docs
    "mail": ["email", "message", "envelope", "contact"],
    "chat": ["conversation", "discussion", "messaging"],
    "forum": ["community", "dialogue", "exchange"],
    "call": ["phone", "telephone", "contact"],
    "notifications": ["bell", "alert", "reminder"],
    "campaign": ["announcement", "megaphone", "marketing", "communication"],
    "share": ["sharing", "distribute", "social"],
    "link": ["url", "connection", "chain"],
    "attach_file": ["attachment", "paperclip", "file"],
    "folder": ["directory", "files", "archive"],
    "description": ["document", "file", "page", "report"],
    "article": ["text", "news", "blog", "documentation"],
    "menu_book": ["handbook", "manual", "guide", "documentation"],
    "library_books": ["library", "knowledge base", "collection"],
    "edit": ["pencil", "modify", "write"],
    "draw": ["design", "sketch", "creative"],
    "search": ["find", "magnifier", "discover", "research"],
    "lightbulb": ["idea", "innovation", "insight", "creativity"],
    "lightbulb_circle": ["tip", "hint", "innovation bulb"],
    # places & world
    "public": ["globe", "world", "global", "international"],
    "language": ["web", "www", "international", "translation"],
    "location_on": ["pin", "place", "map marker", "site"],
    "map": ["geography", "navigation", "territory"],
    "home": ["house", "main", "start"],
    "apartment": ["building", "office", "company", "headquarters"],
    "eco": ["sustainability", "green", "leaf", "environment", "esg"],
    "recycling": ["circular economy", "reuse", "sustainable"],
    "bolt": ["energy", "power", "fast", "electric"],
    "solar_power": ["renewable", "solar", "energy transition"],
    # arrows & actions
    "arrow_forward": ["next", "right", "continue"],
    "arrow_back": ["previous", "left", "return"],
    "arrow_upward": ["up", "increase", "north"],
    "arrow_downward": ["down", "decrease", "south"],
    "chevron_right": ["next", "expand", "breadcrumb"],
    "expand_more": ["dropdown", "more", "down chevron"],
    "open_in_new": ["external", "new window", "link out"],
    "download": ["save", "export", "get"],
    "upload": ["import", "submit", "send up"],
    "add_circle": ["plus", "add", "new", "create"],
    "do_not_disturb_on": ["minus", "remove", "reduce"],
    "delete": ["trash", "bin", "remove", "discard"],
    "content_copy": ["copy", "duplicate", "clone"],
    "filter_alt": ["filter", "funnel", "refine"],
    "sort": ["order", "ranking", "arrange"],
    "zoom_in": ["magnify", "enlarge", "detail"],
    "star": ["favorite", "rating", "highlight"],
    "favorite": ["heart", "like", "love", "engagement"],
    "thumb_up": ["like", "approve", "positive", "agree"],
    "thumb_down": ["dislike", "reject", "negative", "disagree"],
}

LICENSE_TEXT = """Material Symbols (https://fonts.google.com/icons)
Copyright Google LLC.
Licensed under the Apache License, Version 2.0:
https://www.apache.org/licenses/LICENSE-2.0
Vendored subset fetched from the @material-symbols/svg-400 npm package.
"""


def main() -> None:
    icons_dir = DEST / "icons"
    icons_dir.mkdir(parents=True, exist_ok=True)
    index = {
        "set_id": "material",
        "name": "Material Symbols (outlined, curated business subset)",
        "license": "Apache-2.0",
        "recolorable": True,
        "icons": [],
    }
    failed = []
    for i, (name, tags) in enumerate(sorted(ICONS.items())):
        target = icons_dir / f"{name}.svg"
        if not target.is_file():
            try:
                data = urllib.request.urlopen(URL.format(name=name), timeout=30).read()
                target.write_bytes(data)
                time.sleep(0.05)
            except Exception as exc:  # noqa: BLE001 - report and continue
                failed.append((name, str(exc)))
                continue
        index["icons"].append(
            {"id": name, "name": name.replace("_", " "), "tags": tags,
             "file": f"icons/{name}.svg", "format": "svg"}
        )
        if (i + 1) % 25 == 0:
            print(f"{i + 1}/{len(ICONS)} ...")
    (DEST / "index.json").write_text(
        json.dumps(index, indent=1, ensure_ascii=False), encoding="utf-8"
    )
    (DEST / "LICENSE").write_text(LICENSE_TEXT, encoding="utf-8")
    print(f"vendored {len(index['icons'])} icons -> {DEST}")
    if failed:
        print("FAILED:", failed)


if __name__ == "__main__":
    main()
