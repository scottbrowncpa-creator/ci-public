"""blind_guard.py - PreToolUse guard for the blind ("goldfish") reviewer agents: no memory, CLAUDE.md, notes or briefs.

Wired only from the frontmatter of .claude/agents/*-blind.md (hooks: PreToolUse), so it runs for those agents and nobody else.
The agent files already set omitClaudeMd: true and carry no memory field; this guard closes the other door, the tools.

Allows:
  Read            a file that is not on the blocked list below
  Grep / Glob     an explicit directory that is not blocked, does not contain the workspace or ~/.claude, and has no
                  blocked file or folder anywhere under it (walked, capped at WALK_CAP entries; past the cap = deny)
  Bash            python tools/{wb_lint,ship_check,xlsx_probe,close_scan}.py <args>, one command, no shell operators,
                  ship_check only as `xlsx ... --no-stamp`; every path argument checked against the blocked list
Blocks everything else, including any other tool name.
Blocked: any CLAUDE.md / CLAUDE.local.md / MEMORY*.md, anything under a .claude folder (agent-memory, projects = transcripts
and auto memory, skills, settings), any memory folder, and in the workspace: knowledge, import, logs, clients, meetings,
email, scratch, review (except the standards files in REVIEW_OK); the Obsidian Vault; Claude app data and session scratchpads.
Fails closed: unreadable input, unknown tool or any exception = deny. Deny = permissionDecision deny, exit 0 (as path_guard).
Test: python tools/test_blind_review.py
"""
import json
import os
import posixpath
import re
import sys
from pathlib import Path

WS = Path(__file__).resolve().parents[2]
HOME = Path.home()
TOOLS_OK = {"wb_lint.py", "ship_check.py", "xlsx_probe.py", "close_scan.py"}
REVIEW_OK = {"review/close-review-rules.md"}
WS_BLOCKED_DIRS = ("knowledge", "import", "logs", "clients", "meetings", "email", "scratch", "review")
NAME_RX = re.compile(r"^(claude(\.local)?\.md|memory[^/]*\.md)$")
SEG_RX = re.compile(r"(^|/)(\.claude|memory|agent-memory)(/|$)")
OTHER_BLOCKED = ("appdata/local/claude-data", "appdata/local/temp/claude", "appdata/roaming/claude",
                 "local/temp/claude", "desktop/vault")


def norm(p, cwd=None):
    """Lower-case forward-slash absolute form; Git Bash /c/... becomes c:/... ; relative joins cwd (default workspace)."""
    s = str(p).strip().strip('"').strip("'").replace("\\", "/")
    s = os.path.expandvars(os.path.expanduser(s)).replace("\\", "/")
    m = re.match(r"^/([a-zA-Z])(/|$)", s)
    if m:
        s = m.group(1) + ":/" + s[3:]
    if not re.match(r"^[a-zA-Z]:/", s) and not s.startswith("//"):
        base = norm(cwd) if cwd else str(WS).replace("\\", "/").lower()
        s = base.rstrip("/") + "/" + s
    s = posixpath.normpath(s).lower()
    return s


WS_N = norm(str(WS))
HOME_CLAUDE_N = norm(str(HOME / ".claude"))


def blocked(p, cwd=None):
    n = norm(p, cwd)
    name = n.rsplit("/", 1)[-1]
    if NAME_RX.match(name):
        return "CLAUDE.md or memory file"
    if SEG_RX.search(n):
        return ".claude or memory folder"
    if any(o in n for o in OTHER_BLOCKED):
        return "Claude app data, scratchpad or Vault"
    if n == WS_N or n.startswith(WS_N + "/"):
        rel = n[len(WS_N) + 1:]
        top = rel.split("/", 1)[0]
        if top in WS_BLOCKED_DIRS and rel not in REVIEW_OK:
            return f"workspace {top}/ (notes, memory, briefs)"
    return None


def contains_protected(d):
    """A search root that holds the workspace or ~/.claude would reach CLAUDE.md and memory."""
    d = d.rstrip("/")
    for prot in (WS_N, HOME_CLAUDE_N):
        if prot == d or prot.startswith(d + "/"):
            return True
    return False


WALK_CAP = 20000


def holds_blocked(d):
    """Walk a search root; a CLAUDE.md, memory file or .claude folder anywhere under it = the search could read it.
    Past WALK_CAP entries the answer is unknown, so it counts as blocked (and says so)."""
    if not os.path.isdir(d):
        return None
    n = 0
    for root, dirs, files in os.walk(d):
        for name in dirs + files:
            n += 1
            if n > WALK_CAP:
                return f"more than {WALK_CAP} entries, cannot verify; pass a narrower folder"
            why = blocked(os.path.join(root, name))
            if why:
                return f"{why} inside it: {os.path.join(root, name)[-90:]}"
    return None


def check_search(tool, ti, cwd):
    path = ti.get("path") or ""
    if not str(path).strip():
        return f"{tool} needs an explicit path (the default is the workspace root)"
    for pat in (ti.get("pattern") if tool == "Glob" else None, ti.get("glob")):
        if pat and re.search(r"claude|memory|\.\.", str(pat), re.I):
            return f"{tool} pattern reaches CLAUDE.md, memory or a parent folder"
        if pat and re.match(r"^([a-zA-Z]:)?[/\\]", str(pat)):
            return f"{tool} pattern is an absolute path; pass the folder as path"
    d = norm(path, cwd)
    why = blocked(d)
    if why:
        return f"{tool} on a blocked folder ({why})"
    if contains_protected(d):
        return f"{tool} on a folder that contains the workspace or ~/.claude"
    why = holds_blocked(str(path) if os.path.isabs(str(path)) else os.path.join(cwd or str(WS), str(path)))
    if why:
        return f"{tool} on a folder with {why}"
    return None


TOKEN_RX = re.compile(r'"[^"]*"|\'[^\']*\'|\S+')


def check_bash(cmd, cwd):
    if not isinstance(cmd, str) or not cmd.strip():
        return "empty command"
    if re.search(r"[\r\n`$%]", cmd):
        return "newline, backtick, $ or % in the command"
    outside = re.sub(r'"[^"]*"|\'[^\']*\'', "", cmd)
    if re.search(r"[;&|<>()]", outside) or cmd.count('"') % 2 or cmd.count("'") % 2:
        return "shell operator or unbalanced quote"
    toks = TOKEN_RX.findall(cmd)
    if len(toks) < 2 or toks[0].lower() not in ("python", "python3", "py", "python.exe"):
        return "only python tools/<reviewer tool>.py is allowed"
    script = norm(toks[1], cwd)
    if script.rsplit("/", 1)[0] != WS_N + "/tools" or script.rsplit("/", 1)[-1] not in TOOLS_OK:
        return "script is not one of " + ", ".join(sorted(TOOLS_OK))
    args = [t.strip('"').strip("'") for t in toks[2:]]
    if script.endswith("/ship_check.py") and (not args or args[0] != "xlsx" or "--no-stamp" not in args):
        return "ship_check only as: xlsx <file> --no-stamp"
    for a in args:
        if a.startswith("-") and "/" not in a and "\\" not in a:
            continue
        if "/" in a or "\\" in a or "." in a or re.search(r"claude|memory", a, re.I):
            why = blocked(a, cwd)
            if why:
                return f"argument {a[:80]} is blocked ({why})"
            full = a if os.path.isabs(a) else os.path.join(cwd or str(WS), a)
            if os.path.isdir(full):
                if contains_protected(norm(a, cwd)):
                    return f"argument {a[:80]} contains the workspace or ~/.claude"
                inner = holds_blocked(full)
                if inner:
                    return f"argument folder holds {inner}"
    return None


def decide(payload):
    if not isinstance(payload, dict):
        return "unreadable hook input"
    tool = payload.get("tool_name")
    ti = payload.get("tool_input")
    cwd = payload.get("cwd") or None
    if not isinstance(ti, dict):
        return "unreadable tool input"
    if tool == "Read":
        fp = ti.get("file_path")
        if not fp:
            return "Read with no file_path"
        why = blocked(fp, cwd)
        return f"Read of a blocked file ({why})" if why else None
    if tool in ("Grep", "Glob"):
        return check_search(tool, ti, cwd)
    if tool == "Bash":
        return check_bash(ti.get("command"), cwd)
    return f"tool {tool!r} is not open to a blind reviewer"


def emit_deny(reason):
    sys.stdout.write(json.dumps({"hookSpecificOutput": {
        "hookEventName": "PreToolUse", "permissionDecision": "deny",
        "permissionDecisionReason": "blind_guard: " + reason + ". A blind reviewer sees only the deliverable and its primary sources."}}))
    sys.stdout.flush()


def main():
    try:
        raw = sys.stdin.buffer.read().decode("utf-8-sig", errors="replace")
        payload = json.loads(raw) if raw.strip() else None
        why = decide(payload)
    except Exception as e:                      # error path is DENY
        why = f"guard error {type(e).__name__}"
    if why:
        emit_deny(why)
    sys.exit(0)


if __name__ == "__main__":
    main()
