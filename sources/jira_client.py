from __future__ import annotations
import os
from jira import JIRA


def connect(cfg: dict) -> JIRA:
    server = cfg["jira"]["server"]
    pat = os.environ.get("JIRA_PAT")
    if pat:
        return JIRA(server=server, token_auth=pat)
    email = os.environ.get("JIRA_EMAIL")
    api_token = os.environ.get("JIRA_API_TOKEN")
    if email and api_token:
        return JIRA(server=server, basic_auth=(email, api_token))
    raise RuntimeError("Jira 인증 정보 없음 (.env 에 JIRA_PAT 설정).")


def get_objective(jira: JIRA, key: str) -> dict:
    i = jira.issue(key)
    return {
        "key": i.key,
        "summary": i.fields.summary,
        "status": str(i.fields.status),
        "url": f"{jira._options['server']}/browse/{i.key}",
    }


def issues_by_fixversion(jira: JIRA, project: str, fix_version: str) -> list:
    jql = f'project = "{project}" AND fixVersion = "{fix_version}" ORDER BY key ASC'
    return [
        {
            "key": i.key,
            "summary": i.fields.summary,
            "status": str(i.fields.status),
            "type": str(i.fields.issuetype),
        }
        for i in jira.search_issues(jql, maxResults=200)
    ]
