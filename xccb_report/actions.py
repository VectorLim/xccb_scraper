"""Parse the observed server-rendered Actions panel without submitting forms."""

from urllib.parse import parse_qs, urlparse
import re

from bs4 import BeautifulSoup


def workspace_identity(url):
    parsed = urlparse(str(url))
    if (
        parsed.scheme not in {"http", "https"}
        or parsed.hostname != "xccb2.intel.com"
        or parsed.username
        or parsed.password
        or parsed.port not in {None, 80, 443}
        or parsed.path.casefold() != "/documents/documentworkspace.aspx"
    ):
        raise ValueError("Invalid xCCB Document Workspace URL")
    query = parse_qs(parsed.query)
    ids = [
        v for k, values in query.items() if k.casefold() == "documentid" for v in values
    ]
    if len(ids) != 1 or not ids[0].isdigit():
        raise ValueError("Missing or ambiguous DocumentID")
    return (
        ids[0],
        f"https://xccb2.intel.com/Documents/DocumentWorkspace.aspx?DocumentID={ids[0]}",
    )


def text(node):
    return " ".join(node.get_text().split())


def parse_actions(html, document_id, empty_state_text=""):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table", class_="action-panel")
    errors, actions, seen = [], [], set()
    empty_verified = bool(empty_state_text.strip()) and any(
        text(n) == empty_state_text.strip()
        for n in soup.find_all(["span", "div", "p", "td"])
    )
    if not tables and empty_verified:
        return [], []
    if len(tables) != 1:
        return [], [
            "Expected exactly one action-panel table; unloaded/empty layout is not verified"
        ]
    table = tables[0]
    for row in table.find_all("tr"):
        links = [
            n
            for n in row.find_all("a")
            if "/actions/completedocumentaction.aspx" in n.get("href", "").casefold()
        ]
        if not links:
            continue  # Observed second row contains status, never the assigner.
        if len(links) != 1:
            errors.append("Action row has multiple record links")
            continue
        query = parse_qs(urlparse(links[0].attrs["href"]).query)
        ids, docs = query.get("ActionID", []), query.get("DocumentID", [])
        assigners = [
            re.sub(r"^Assigned\s+By\s*:\s*", "", text(n), flags=re.I)
            for n in row.find_all("span")
            if re.match(r"^Assigned\s+By\s*:", text(n), re.I)
        ]
        if (
            len(ids) != 1
            or not ids[0].isdigit()
            or docs != [str(document_id)]
            or len(assigners) != 1
            or not assigners[0]
        ):
            errors.append(
                "Malformed action identity, document reference, or Assigned By field"
            )
            continue
        if ids[0] in seen:
            errors.append(f"Duplicate ActionID {ids[0]}")
            continue
        seen.add(ids[0])
        actions.append({"action_id": ids[0], "assigned_by": assigners[0]})
    # A changed layout must not silently disappear from the extraction.
    labels = [
        n
        for n in table.find_all("span")
        if re.match(r"^Assigned\s+By\s*:", text(n), re.I)
    ]
    if len(labels) != len(actions):
        errors.append("Assigned By count does not reconcile with parsed actions")
    controls = soup.find_all(["a", "button", "select"])
    if any(
        re.search(
            r"Page\$|pager|pagination|show\s+all|next\s*(page|[>»])|load\s+more",
            " ".join([text(n), *map(str, n.attrs.values())]),
            re.I,
        )
        for n in controls
    ):
        errors.append(
            "Possible pagination or incomplete listing; manual validation required"
        )
    if not actions and not empty_verified:
        errors.append(
            "Empty Actions layout requires explicit live empty-state validation"
        )
    return actions, errors
