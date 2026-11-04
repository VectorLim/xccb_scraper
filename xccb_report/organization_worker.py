"""Read-only HR lookup in a bounded source or bundled worker process."""

import ast
import contextlib
import io
import json
from pathlib import Path
import re
import sys


def main():
    request_path = Path(sys.argv[2]) if len(sys.argv) == 4 and sys.argv[1] == "--organization-worker" else None
    response_path = Path(sys.argv[3]) if request_path else None
    def respond(value):
        text = json.dumps(value)
        if response_path is not None:
            response_path.write_text(text, encoding="utf-8")
        else:
            sys.stdout.write(text)
    try:
        request = json.loads(request_path.read_text(encoding="utf-8") if request_path else sys.stdin.read())
        batch = isinstance(request, dict)
        names = request["names"] if batch else [request[0]]
        names = list(dict.fromkeys(" ".join(name.split()).upper() for name in names if name.strip()))
        if not names:
            respond({} if batch else [])
            return 0
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(
            io.StringIO()
        ):
            import datasyncx
            from pandas import isna
            from datasyncx.utils.oracle_db import query_oracle

            helper = Path(datasyncx.__file__).parent / "utils" / "get_user_email.py"
            tree = ast.parse(helper.read_text(encoding="utf-8"))
            account = next(
                node.value.value
                for node in ast.walk(tree)
                if isinstance(node, ast.Assign)
                and isinstance(node.value, ast.Constant)
                and any(
                    isinstance(target, ast.Name) and target.id == "username"
                    for target in node.targets
                )
            )
            literal = lambda value: "'" + value.replace("'", "''") + "'"
            predicate = " OR ".join(
                "REGEXP_REPLACE(UPPER(TRIM(CCMAILNAME)), '[[:space:]]+', ' ') IN ("
                + ", ".join(literal(name) for name in names[start:start + 1000])
                + ")"
                for start in range(0, len(names), 1000)
            )
            rows = query_oracle(
                acc=account,
                pwd=account,
                node="ALL.EMPLOYEE_HRDS",
                sql=f"""SELECT WWID,
                CCMAILNAME, SITECODE, ORGUNITCODE, ORGUNITDESCR, GL_COST_CENTER_DESCR
                FROM ATHRDS_WORKER WHERE GENIEMPSTATUS NOT IN ('R', 'T') AND ({predicate})""",
            ).to_dict(orient="records")

            def text(value):
                return "" if isna(value) else str(value)

            result = [
                {
                    "wwid": text(row["WWID"]),
                    "name": text(row["CCMAILNAME"]),
                    "site": text(row["SITECODE"]),
                    "organization_code": text(row["ORGUNITCODE"]),
                    "organization": text(row["ORGUNITDESCR"]),
                    "cost_center": text(row["GL_COST_CENTER_DESCR"]),
                }
                for row in rows
            ]
        if batch:
            grouped = {name.casefold(): [] for name in names}
            for row in result:
                key = " ".join(row["name"].split()).casefold()
                if key in grouped:
                    grouped[key].append(row)
            respond(grouped)
        else:
            respond(result)
        return 0
    except Exception as exc:
        codes = sorted(set(re.findall(r"(?:ORA|DPI|DPY)-\d+", str(exc))))
        respond(
            {
                    "error": "DataSyncX HR lookup failed: "
                    + type(exc).__name__
                    + (" (" + ", ".join(codes) + ")" if codes else "")
                    + "; check DataSyncX Python, Oracle client/TNS and corporate network access"
                }
        )
        return 2


if __name__ == "__main__":
    sys.exit(main())
