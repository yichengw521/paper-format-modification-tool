# Backend API

Base URL: `http://127.0.0.1:8080/api/v1`

## Health

`GET /health`

## Create formatting task

`POST /tasks` with `multipart/form-data`:

- `template`: format template DOCX
- `document`: document to format DOCX

The response status is `202 Accepted`. This request starts analysis only. Use the returned `links.status` URL to poll the task.

## Task status

`GET /tasks/{taskId}`

Possible states: `ANALYZING`, `AWAITING_CONFIRMATION`, `QUEUED`, `PROCESSING`, `COMPLETED`, and `FAILED`.

## Review and confirm the format plan

- `GET /tasks/{taskId}/analysis`: read the extracted page, cover, abstract, TOC, body, reference, and acknowledgement rules when the state is `AWAITING_CONFIRMATION`
- `POST /tasks/{taskId}/confirm` with JSON such as `{"enabledRuleKeys":["page","cover","abstract-zh","abstract-en","toc","headings","body","captions","references","thanks"]}`: persist the confirmation and begin formatting

No output DOCX is created before the confirmation request.

## Downloads

- `GET /tasks/{taskId}/result`: formatted DOCX
- `GET /tasks/{taskId}/report`: JSON modification report

Downloads become available when the task state is `COMPLETED`.
