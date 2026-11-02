"""Executed by the user's approved DataSyncX Python, not the bundled Python."""
import ast
import contextlib
import io
import json
from pathlib import Path
import re
import sys


def main():
    try:
        name, wwid = json.loads(sys.stdin.read())
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            import datasyncx
            from pandas import isna
            from datasyncx.utils.oracle_db import query_oracle
            helper = Path(datasyncx.__file__).parent / "utils" / "get_user_email.py"
            tree = ast.parse(helper.read_text(encoding="utf-8"))
            account = next(node.value.value for node in ast.walk(tree)
                           if isinstance(node, ast.Assign) and isinstance(node.value, ast.Constant)
                           and any(isinstance(target, ast.Name) and target.id == "username" for target in node.targets))
            literal = lambda value: "'" + value.replace("'", "''") + "'"
            if wwid:
                predicate = "WWID = " + literal(wwid)
            else:
                predicate = " OR ".join(f"UPPER(TRIM({expression})) = {literal(' '.join(name.split()).upper())}"
                                        for expression in ("CCMAILNAME", "LASTNAME || ', ' || FIRSTNAME", "LASTNAME || ', ' || NICKNAME"))
            rows = query_oracle(acc=account, pwd=account, node="ALL.EMPLOYEE_HRDS", sql=f"""SELECT WWID,
                CCMAILNAME, SITECODE, ORGUNITCODE, ORGUNITDESCR, GL_COST_CENTER_DESCR
                FROM ATHRDS_WORKER WHERE GENIEMPSTATUS NOT IN ('R', 'T') AND ({predicate})""").to_dict(orient="records")
            def text(value):
                return "" if isna(value) else str(value)
            result = [{"wwid": text(row["WWID"]), "name": text(row["CCMAILNAME"]), "site": text(row["SITECODE"]),
                       "organization_code": text(row["ORGUNITCODE"]), "organization": text(row["ORGUNITDESCR"]),
                       "cost_center": text(row["GL_COST_CENTER_DESCR"])} for row in rows]
        sys.stdout.write(json.dumps(result))
        return 0
    except Exception as exc:
        codes = sorted(set(re.findall(r"(?:ORA|DPI|DPY)-\d+", str(exc))))
        sys.stdout.write(json.dumps({"error": "DataSyncX HR lookup failed: " + type(exc).__name__ +
                                    (" (" + ", ".join(codes) + ")" if codes else "") +
                                    "; check DataSyncX Python, Oracle client/TNS and corporate network access"}))
        return 2


if __name__ == "__main__":
    sys.exit(main())